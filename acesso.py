"""Login e gestão de senhas do CGAPE.

Histórico: login criado em 25/11/2026 (usuários e hashes escritos aqui);
desde 27/11/2026 (etapa 1 da "gestão de senhas", pedido do usuário) o
cadastro fica FORA do código e do GitHub público, num arquivo da rede:

    X:\\COAPI\\NOTAS TÉCNICAS COAPI\\2024-2026\\PANORAMA\\2026\\SISTEMA CGAPE\\usuarios.json

(variável de ambiente CGAPE_SISTEMA_PASTA troca a pasta, para testes), com
trava entre máquinas, gravação atômica, número de revisão e backups — o
mesmo esquema do Controle de Prazos. Uma cópia de reserva só para leitura
fica em %LOCALAPPDATA%\\CGAPE (login fora da rede; sem troca de senha).

Regras (decididas pelo usuário em 27/11/2026):
  - administrador: só Diego (perfil "admin");
  - o admin define uma senha PROVISÓRIA (vale 7 dias); no primeiro login o
    usuário é obrigado a trocá-la — enquanto não troca, a sessão não tem
    NENHUMA seção liberada (conferido no Python, não só na tela);
  - senha nova: 8+ caracteres, com letra, número e caractere especial,
    diferente da atual e sem o nome de usuário;
  - 5 erros seguidos bloqueiam por 15 minutos;
  - senhas nunca são gravadas: só PBKDF2-SHA256 com sal próprio.

Acesso livre (sem login): só PAC e MONITORAMENTO.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import json
import os
import re
import shutil
import tempfile
import time

ITERACOES = 310_000
SECOES_TODAS = ("PAC", "MON", "PRZ", "DOE")
SECOES_LIVRE = ("PAC", "MON")
PERFIS = ("admin", "usuario")

PASTA_PADRAO = r"X:\COAPI\NOTAS TÉCNICAS COAPI\2024-2026\PANORAMA\2026\SISTEMA CGAPE"
NOME_ARQUIVO = "usuarios.json"
NOME_BACKUPS = "backups usuarios"
MAX_BACKUPS = 100

DIAS_PROVISORIA = 7
MAX_TENTATIVAS = 5
MINUTOS_BLOQUEIO = 15
TAMANHO_MINIMO = 8

# Quem pode criar/editar/excluir ações do CONTROLE DE PRAZOS (25/11/2026):
# "apenas Silvana, os outros apenas a visualização"; Diego também edita
# desde 27/11/2026 (pedido do usuário).
EDITORES_PRAZOS = ("silvana.silva", "diego.figueiredo")


def pode_editar_prazos(sessao) -> bool:
    return str((sessao or {}).get("usuario") or "").lower() in EDITORES_PRAZOS


# Gerador de gráficos do SISTEMA PRISIONAL (sistema_prisional.py), dentro da
# seção PRAZOS — pedido do usuário em 27/11/2026: só Silvana e Diego veem e
# editam.
ACESSO_SISTEMA_PRISIONAL = ("silvana.silva", "diego.figueiredo")


def pode_sistema_prisional(sessao) -> bool:
    s = sessao or {}
    return "PRZ" in (s.get("secoes") or []) and str(s.get("usuario") or "").lower() in ACESSO_SISTEMA_PRISIONAL


def eh_admin(sessao) -> bool:
    return (sessao or {}).get("perfil") == "admin" and bool((sessao or {}).get("secoes"))


# ------------------------------------------------------------ arquivos ----
def pasta_sistema() -> str:
    return os.environ.get("CGAPE_SISTEMA_PASTA") or PASTA_PADRAO


def caminho_usuarios() -> str:
    return os.path.join(pasta_sistema(), NOME_ARQUIVO)


def caminho_reserva() -> str:
    pasta = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), "CGAPE")
    nome = "usuarios_reserva.json" if not os.environ.get("CGAPE_SISTEMA_PASTA") else "usuarios_reserva_teste.json"
    return os.path.join(pasta, nome)


def _agora():
    return dt.datetime.now()


def _iso(d):
    return d.isoformat(timespec="seconds")


def _de_iso(t):
    try:
        return dt.datetime.fromisoformat(t) if t else None
    except (TypeError, ValueError):
        return None


def _ler_json(caminho):
    with open(caminho, encoding="utf-8") as f:
        return json.load(f)


def _gravar_json_atomico(caminho, dados):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".usuarios_", suffix=".tmp", dir=os.path.dirname(caminho))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=1)
        os.replace(tmp, caminho)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def _espelhar_reserva(dados):
    try:
        _gravar_json_atomico(caminho_reserva(), dados)
    except OSError:
        pass


def _backup():
    origem = caminho_usuarios()
    if not os.path.isfile(origem):
        return
    pasta = os.path.join(pasta_sistema(), NOME_BACKUPS)
    try:
        os.makedirs(pasta, exist_ok=True)
        shutil.copy2(origem, os.path.join(pasta, f"usuarios_{_agora():%Y%m%d_%H%M%S_%f}.json"))
        antigos = sorted(f for f in os.listdir(pasta) if f.startswith("usuarios_"))
        for f in antigos[:-MAX_BACKUPS]:
            os.remove(os.path.join(pasta, f))
    except OSError:
        pass


class _Trava:
    """Trava entre máquinas (arquivo .lock exclusivo; expira em 30 s)."""
    def __init__(self, espera=10.0):
        self.caminho = caminho_usuarios() + ".lock"
        self.espera = espera
        self.fd = None

    def __enter__(self):
        limite = time.time() + self.espera
        while True:
            try:
                self.fd = os.open(self.caminho, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(self.fd, str(os.getpid()).encode())
                return self
            except FileExistsError:
                try:
                    idade = time.time() - os.path.getmtime(self.caminho)
                except OSError:
                    idade = 0
                if idade > 30:
                    try:
                        os.remove(self.caminho)
                    except OSError:
                        pass
                    continue
                if time.time() > limite:
                    raise TimeoutError("o cadastro de usuários está sendo gravado em outra máquina — tente de novo")
                time.sleep(0.2)

    def __exit__(self, *exc):
        try:
            os.close(self.fd)
        finally:
            try:
                os.remove(self.caminho)
            except OSError:
                pass


def _carregar():
    """(dados, gravavel). Rede primeiro; sem rede, a cópia de reserva (só leitura)."""
    principal = caminho_usuarios()
    if os.path.isfile(principal):
        try:
            dados = _ler_json(principal)
            _espelhar_reserva(dados)
            return dados, True
        except (OSError, ValueError):
            pass
    reserva = caminho_reserva()
    if os.path.isfile(reserva):
        try:
            return _ler_json(reserva), False
        except (OSError, ValueError):
            pass
    return None, False


def _alterar(funcao):
    """Aplica funcao(usuarios) -> resultado sob a trava e grava (com backup)."""
    if not os.path.isfile(caminho_usuarios()):
        raise OSError("cadastro de usuários inacessível (pasta da rede fora do ar)")
    with _Trava():
        dados = _ler_json(caminho_usuarios())
        resultado = funcao(dados.setdefault("usuarios", {}))
        _backup()
        dados["revisao"] = int(dados.get("revisao") or 0) + 1
        dados["alterado_em"] = _agora().strftime("%d/%m/%Y %H:%M")
        _gravar_json_atomico(caminho_usuarios(), dados)
        _espelhar_reserva(dados)
    return resultado


# --------------------------------------------------------------- senha ----
def _hash(senha: str, sal_hex: str, iteracoes: int = ITERACOES) -> str:
    return hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), bytes.fromhex(sal_hex), iteracoes).hex()


def _novo_hash(senha):
    sal = os.urandom(16).hex()
    return {"sal": sal, "hash": _hash(senha, sal), "iteracoes": ITERACOES}


def _confere(registro, senha) -> bool:
    try:
        esperado = registro["hash"]
        calculado = _hash(str(senha), registro["sal"], int(registro.get("iteracoes") or ITERACOES))
    except (KeyError, ValueError, TypeError):
        return False
    return hmac.compare_digest(calculado, esperado)


def problemas_senha(nova, usuario="", atual=None):
    """Lista de regras não atendidas (vazia = senha aceita)."""
    nova = str(nova or "")
    erros = []
    if len(nova) < TAMANHO_MINIMO:
        erros.append(f"ter pelo menos {TAMANHO_MINIMO} caracteres")
    if not re.search(r"[^\W\d_]", nova):
        erros.append("ter pelo menos uma letra")
    if not re.search(r"\d", nova):
        erros.append("ter pelo menos um número")
    if not re.search(r"[^\w\s]|_", nova):
        erros.append("ter pelo menos um caractere especial (ex.: @ # $ % & * ! ?)")
    if re.search(r"\s", nova):
        erros.append("não ter espaços")
    u = str(usuario or "").lower()
    partes = [p for p in re.split(r"[._\-]", u) if len(p) >= 4]
    if u and (u in nova.lower() or any(p in nova.lower() for p in partes)):
        erros.append("não conter o seu nome de usuário")
    if atual is not None and nova == atual:
        erros.append("ser diferente da senha atual")
    return erros


# ------------------------------------------------------------- sessões ----
def _sessao_de(chave, reg, liberar=True):
    return {"usuario": chave, "nome": reg.get("nome") or chave,
            "perfil": reg.get("perfil") or "usuario",
            "secoes": [s for s in (reg.get("secoes") or SECOES_TODAS) if s in SECOES_TODAS] if liberar else []}


def sessao_usuario(chave):
    """Sessão de um usuário já autenticado (relançamento do programa).
    Senha pendente de troca ou usuário inativo -> sem seções."""
    chave = str(chave or "").strip().lower()
    dados, _g = _carregar()
    reg = ((dados or {}).get("usuarios") or {}).get(chave)
    if not reg or not reg.get("ativo", True):
        return None
    return _sessao_de(chave, reg, liberar=not reg.get("trocar_senha"))


def sessao_livre():
    return {"usuario": "", "nome": "Acesso livre", "perfil": "", "secoes": list(SECOES_LIVRE)}


def entrar(usuario, senha):
    """Login. Devolve dict:
      {"ok": True, "sessao": {...}}                          -> entra
      {"ok": False, "trocar_senha": True, "usuario", "nome"} -> senha certa, mas provisória
      {"ok": False, "mensagem": "..."}                       -> recusado
    """
    chave = str(usuario or "").strip().lower()
    dados, gravavel = _carregar()
    if dados is None:
        _hash(str(senha or "x"), "00" * 16)
        return {"ok": False, "mensagem": "Cadastro de usuários inacessível (pasta da rede fora do ar). "
                                         "Use o ACESSO LIVRE ou tente de novo dentro da rede."}
    reg = (dados.get("usuarios") or {}).get(chave)
    if not reg or not senha or not reg.get("ativo", True):
        _hash(str(senha or "x"), "00" * 16)   # mesmo custo: não revela quem existe
        return {"ok": False, "mensagem": "Usuário ou senha incorretos."}
    bloqueio = _de_iso(reg.get("bloqueado_ate"))
    if bloqueio and bloqueio > _agora():
        minutos = max(1, int((bloqueio - _agora()).total_seconds() // 60) + 1)
        return {"ok": False, "mensagem": f"Acesso bloqueado por excesso de tentativas. Tente de novo em {minutos} min "
                                         "ou peça ao administrador para desbloquear."}
    if not _confere(reg, senha):
        if gravavel:
            try:
                def falha(us):
                    r = us[chave]
                    r["tentativas_falhas"] = int(r.get("tentativas_falhas") or 0) + 1
                    if r["tentativas_falhas"] >= MAX_TENTATIVAS:
                        r["bloqueado_ate"] = _iso(_agora() + dt.timedelta(minutes=MINUTOS_BLOQUEIO))
                        r["tentativas_falhas"] = 0
                        return True
                    return False
                if _alterar(falha):
                    return {"ok": False, "mensagem": f"Senha incorreta {MAX_TENTATIVAS} vezes seguidas: acesso "
                                                     f"bloqueado por {MINUTOS_BLOQUEIO} minutos."}
            except (OSError, TimeoutError, ValueError):
                pass
        return {"ok": False, "mensagem": "Usuário ou senha incorretos."}
    # senha certa
    if reg.get("trocar_senha"):
        expira = _de_iso(reg.get("provisoria_expira_em"))
        if expira and expira < _agora():
            return {"ok": False, "mensagem": "Sua senha provisória expirou. Peça ao administrador uma nova."}
        if not gravavel:
            return {"ok": False, "mensagem": "Você precisa definir uma senha nova, e isso só é possível dentro "
                                             "da rede (pasta X:). Conecte-se e tente de novo."}
        return {"ok": False, "trocar_senha": True, "usuario": chave, "nome": reg.get("nome") or chave}
    if gravavel and (reg.get("tentativas_falhas") or reg.get("bloqueado_ate")):
        try:
            def zera(us):
                us[chave]["tentativas_falhas"] = 0
                us[chave].pop("bloqueado_ate", None)
                us[chave]["ultimo_login"] = _iso(_agora())
            _alterar(zera)
        except (OSError, TimeoutError, ValueError):
            pass
    return {"ok": True, "sessao": _sessao_de(chave, reg)}


def autenticar(usuario, senha):
    """Compatibilidade: sessão ou None (senha provisória conta como None)."""
    r = entrar(usuario, senha)
    return r.get("sessao") if r.get("ok") else None


def trocar_senha(usuario, atual, nova, confirmacao):
    """Troca obrigatória (senha provisória) ou voluntária. Confere a senha
    atual de novo no servidor. Devolve {"ok", "mensagem", "sessao"?}."""
    chave = str(usuario or "").strip().lower()
    if str(nova or "") != str(confirmacao or ""):
        return {"ok": False, "mensagem": "A confirmação não é igual à senha nova."}
    dados, gravavel = _carregar()
    reg = ((dados or {}).get("usuarios") or {}).get(chave)
    if not reg or not reg.get("ativo", True) or not _confere(reg, atual):
        _hash(str(atual or "x"), "00" * 16)
        return {"ok": False, "mensagem": "Usuário ou senha atual incorretos."}
    if not gravavel:
        return {"ok": False, "mensagem": "Troca de senha indisponível fora da rede (pasta X:)."}
    if reg.get("trocar_senha"):
        expira = _de_iso(reg.get("provisoria_expira_em"))
        if expira and expira < _agora():
            return {"ok": False, "mensagem": "Sua senha provisória expirou. Peça ao administrador uma nova."}
    erros = problemas_senha(nova, chave, atual)
    if erros:
        return {"ok": False, "mensagem": "A senha nova precisa " + "; ".join(erros) + "."}

    def aplicar(us):
        r = us[chave]
        r.update(_novo_hash(nova))
        r["trocar_senha"] = False
        r.pop("provisoria_expira_em", None)
        r["senha_alterada_em"] = _iso(_agora())
        r["tentativas_falhas"] = 0
        r.pop("bloqueado_ate", None)
        r["ultimo_login"] = _iso(_agora())
        return dict(r)
    try:
        novo = _alterar(aplicar)
    except (OSError, TimeoutError, ValueError) as exc:
        return {"ok": False, "mensagem": f"Não foi possível gravar a senha: {exc}"}
    return {"ok": True, "mensagem": "Senha alterada.", "sessao": _sessao_de(chave, novo)}


# ------------------------------------------------- administração (Diego) ----
# Usadas pela tela "Gestão de usuários" (etapa 4) e, até lá, pela linha de
# comando:  python -c "import acesso; print(acesso.criar_usuario(...))"
def _validar_secoes(secoes):
    secoes = [str(s).upper() for s in (secoes or SECOES_TODAS)]
    ruins = [s for s in secoes if s not in SECOES_TODAS]
    if ruins:
        raise ValueError(f"seção desconhecida: {', '.join(ruins)} (use {', '.join(SECOES_TODAS)})")
    return secoes


def criar_usuario(usuario, nome, senha_provisoria, secoes=None, perfil="usuario", por="admin"):
    chave = str(usuario or "").strip().lower()
    if not re.fullmatch(r"[a-z0-9][a-z0-9._-]{2,40}", chave):
        raise ValueError("usuário inválido (use nome.sobrenome, sem acentos)")
    if perfil not in PERFIS:
        raise ValueError("perfil inválido")
    if not senha_provisoria or len(str(senha_provisoria)) < 6:
        raise ValueError("senha provisória muito curta (mínimo 6)")
    secoes = _validar_secoes(secoes)

    def aplicar(us):
        if chave in us:
            raise ValueError("usuário já existe")
        us[chave] = {"nome": nome or chave, "perfil": perfil, "secoes": secoes, "ativo": True,
                     **_novo_hash(senha_provisoria), "trocar_senha": True,
                     "provisoria_expira_em": _iso(_agora() + dt.timedelta(days=DIAS_PROVISORIA)),
                     "tentativas_falhas": 0, "criado_em": _iso(_agora()), "criado_por": por}
        return chave
    return _alterar(aplicar)


def redefinir_senha(usuario, senha_provisoria, por="admin"):
    """Nova senha provisória (vale 7 dias) + troca obrigatória; desbloqueia."""
    chave = str(usuario or "").strip().lower()
    if not senha_provisoria or len(str(senha_provisoria)) < 6:
        raise ValueError("senha provisória muito curta (mínimo 6)")

    def aplicar(us):
        if chave not in us:
            raise ValueError("usuário não encontrado")
        r = us[chave]
        r.update(_novo_hash(senha_provisoria))
        r["trocar_senha"] = True
        r["provisoria_expira_em"] = _iso(_agora() + dt.timedelta(days=DIAS_PROVISORIA))
        r["tentativas_falhas"] = 0
        r.pop("bloqueado_ate", None)
        r["redefinida_em"] = _iso(_agora())
        r["redefinida_por"] = por
        return chave
    return _alterar(aplicar)


def alterar_usuario(usuario, nome=None, secoes=None, perfil=None, ativo=None, por="admin"):
    chave = str(usuario or "").strip().lower()

    def aplicar(us):
        if chave not in us:
            raise ValueError("usuário não encontrado")
        r = us[chave]
        if nome is not None:
            r["nome"] = nome
        if secoes is not None:
            r["secoes"] = _validar_secoes(secoes)
        if perfil is not None:
            if perfil not in PERFIS:
                raise ValueError("perfil inválido")
            r["perfil"] = perfil
        if ativo is not None:
            r["ativo"] = bool(ativo)
        r["alterado_em"] = _iso(_agora())
        r["alterado_por"] = por
        return chave
    return _alterar(aplicar)


def desbloquear(usuario, por="admin"):
    chave = str(usuario or "").strip().lower()

    def aplicar(us):
        if chave not in us:
            raise ValueError("usuário não encontrado")
        us[chave]["tentativas_falhas"] = 0
        us[chave].pop("bloqueado_ate", None)
        return chave
    return _alterar(aplicar)


def listar_usuarios():
    """Resumo SEM hash/sal (para telas)."""
    dados, _g = _carregar()
    saida = []
    for chave, r in sorted(((dados or {}).get("usuarios") or {}).items()):
        saida.append({"usuario": chave, "nome": r.get("nome"), "perfil": r.get("perfil"),
                      "secoes": r.get("secoes"), "ativo": r.get("ativo", True),
                      "trocar_senha": bool(r.get("trocar_senha")),
                      "provisoria_expira_em": r.get("provisoria_expira_em"),
                      "bloqueado_ate": r.get("bloqueado_ate"), "ultimo_login": r.get("ultimo_login"),
                      "senha_alterada_em": r.get("senha_alterada_em")})
    return saida
