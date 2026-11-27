# -*- coding: utf-8 -*-
"""
SISTEMA PRISIONAL - GRÁFICOS SEMANAIS, dentro do CGAPE (seção PRAZOS).

Incorporado em 27/11/2026 a pedido do usuário a partir do app avulso
"X:/COAPI 2/2026/20. SISTEMA PRISIONAL/GRÁFICOS - SISTEMA PRISIONAL/
SISTEMA PRISIONAL - GRAFICOS.py" (que continua existindo). Interface (HTML/JS)
e leitura da planilha são as do app original; aqui mudam só:
  - a planilha é sempre a da pasta da rede (CAMINHO_PLANILHA abaixo;
    variável de ambiente CGAPE_PRISIONAL_PLANILHA troca, para testes);
  - abre numa JANELA PRÓPRIA do pywebview, a partir do botão SISTEMA
    PRISIONAL do Controle de Prazos (abrir_janela);
  - só Silvana e Diego (acesso.ACESSO_SISTEMA_PRISIONAL) — conferido no
    Python ao abrir e em CADA chamada da API da janela.
A planilha (dados) nunca vai para o GitHub: só este código.
"""
import os
import sys
import io
import json
import time
import base64
import shutil
import subprocess
import tempfile
import unicodedata
import webbrowser
from datetime import datetime, date

import openpyxl

ERRO_IMPORT_WEBVIEW = ""
try:
    import webview
except Exception as _e:
    webview = None
    ERRO_IMPORT_WEBVIEW = f"{type(_e).__name__}: {_e}"

NOME_APP = "Sistema Prisional - Gráficos semanais"
NOME_PLANILHA = "BASE - SISTEMA PRISIONAL.xlsx"

PASTA_PADRAO = r"X:\COAPI 2\2026\20. SISTEMA PRISIONAL\GRÁFICOS - SISTEMA PRISIONAL"
CAMINHO_PLANILHA = os.environ.get("CGAPE_PRISIONAL_PLANILHA") or os.path.join(PASTA_PADRAO, NOME_PLANILHA)
PASTA_BASE = os.path.dirname(CAMINHO_PLANILHA)

LINK_SEAP = "https://www.ba.gov.br/seap/publicacoes/dadosestatisticos"
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")


# ---------------------------------------------------------------------------
# LEITURA DA PLANILHA
# ---------------------------------------------------------------------------
def _norm(txt):
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return " ".join(txt.upper().split())


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(".", "").replace(",", ".").strip())
    except ValueError:
        return "ERRO"


def _data(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _linha_cabecalho(ws, chave):
    for r in range(1, 15):
        for c in range(1, 30):
            if _norm(ws.cell(row=r, column=c).value).startswith(chave):
                return r
    return None


def ler_planilha():
    """Leitura da planilha do app avulso — desde 27/11/2026 usada só para a
    IMPORTAÇÃO INICIAL do cadastro (ver CADASTRO abaixo)."""
    if not os.path.isfile(CAMINHO_PLANILHA):
        return {"ok": False, "erro": f"Planilha não encontrada:\n{CAMINHO_PLANILHA}\n\n"
                                      f"Coloque o arquivo \"{NOME_PLANILHA}\" na mesma pasta do aplicativo."}

    # Copia para um temporário antes de ler: assim funciona mesmo com a
    # planilha aberta no Excel por outra pessoa na rede.
    temp = os.path.join(tempfile.gettempdir(), f"sp_base_{os.getpid()}.xlsx")
    try:
        shutil.copy2(CAMINHO_PLANILHA, temp)
        wb = openpyxl.load_workbook(temp, data_only=True, read_only=False)
    except Exception as e:
        return {"ok": False, "erro": f"Não consegui abrir a planilha: {e}"}

    avisos = []

    # --- UNIDADES
    unidades = []
    if "UNIDADES" in wb.sheetnames:
        ws = wb["UNIDADES"]
        r0 = _linha_cabecalho(ws, "UNIDADE") or 4
        cab = {_norm(ws.cell(row=r0, column=c).value): c for c in range(1, ws.max_column + 1)}

        def col(prefixo):
            for k, c in cab.items():
                if k.startswith(prefixo):
                    return c
            return None

        c_nome, c_exib, c_cap = col("UNIDADE"), col("NOME NO GRAFICO"), col("CAPACIDADE")
        c_mostrar, c_cor, c_reg = col("EXIBIR"), col("COR"), col("REGIAO")
        for r in range(r0 + 1, ws.max_row + 1):
            nome = ws.cell(row=r, column=c_nome).value
            if not nome:
                continue
            cap = _num(ws.cell(row=r, column=c_cap).value) if c_cap else None
            if cap == "ERRO":
                avisos.append(f"UNIDADES, linha {r}: capacidade não numérica.")
                cap = None
            unidades.append({
                "nome": str(nome).strip(),
                "exibicao": str(ws.cell(row=r, column=c_exib).value or "").strip() if c_exib else "",
                "regiao": str(ws.cell(row=r, column=c_reg).value or "").strip() if c_reg else "",
                "capacidade": cap,
                "exibir": _norm(ws.cell(row=r, column=c_mostrar).value) != "NAO" if c_mostrar else True,
                "cor": str(ws.cell(row=r, column=c_cor).value or "").strip() if c_cor else "",
            })
    else:
        avisos.append("Aba UNIDADES não encontrada.")

    # --- SALVADOR E RMS
    ssa = {"datas": [], "series": {}, "fontes": [], "obs": []}
    if "SSA_RMS_SEMANAL" in wb.sheetnames:
        ws = wb["SSA_RMS_SEMANAL"]
        r0 = _linha_cabecalho(ws, "DATA") or 4
        cols_unid, c_fonte, c_obs = [], None, None
        for c in range(2, ws.max_column + 1):
            h = ws.cell(row=r0, column=c).value
            if not h:
                continue
            n = _norm(h)
            if n.startswith("FONTE"):
                c_fonte = c
            elif n.startswith("OBSERVA"):
                c_obs = c
            else:
                cols_unid.append((c, str(h).strip()))
        nomes_cad = {_norm(u["nome"]) for u in unidades}
        for _, h in cols_unid:
            if _norm(h) not in nomes_cad:
                avisos.append(f"A coluna \"{h}\" da aba SSA_RMS_SEMANAL não está cadastrada em UNIDADES "
                              f"(sem capacidade; aparece só se marcada no painel).")
                unidades.append({"nome": h, "exibicao": "", "regiao": "SALVADOR E RMS",
                                 "capacidade": None, "exibir": False, "cor": ""})
        linhas = []
        for r in range(r0 + 1, ws.max_row + 1):
            bruto = ws.cell(row=r, column=1).value
            vals = {h: _num(ws.cell(row=r, column=c).value) for c, h in cols_unid}
            if bruto is None and all(v is None for v in vals.values()):
                continue
            d = _data(bruto)
            if d is None:
                avisos.append(f"SSA_RMS_SEMANAL, linha {r}: data inválida ou vazia. Linha ignorada.")
                continue
            for h, v in vals.items():
                if v == "ERRO":
                    avisos.append(f"SSA_RMS_SEMANAL, linha {r}, \"{h}\": valor não numérico. Ignorado.")
                    vals[h] = None
            linhas.append((d, vals,
                           str(ws.cell(row=r, column=c_fonte).value or "").strip() if c_fonte else "",
                           str(ws.cell(row=r, column=c_obs).value or "").strip() if c_obs else ""))
        linhas.sort(key=lambda x: x[0])
        vistos = set()
        for d, vals, fonte, obs in linhas:
            if d in vistos:
                avisos.append(f"SSA_RMS_SEMANAL: a data {d:%d/%m/%Y} aparece mais de uma vez.")
            vistos.add(d)
            ssa["datas"].append(d.isoformat())
            ssa["fontes"].append(fonte)
            ssa["obs"].append(obs)
            for _, h in cols_unid:
                ssa["series"].setdefault(h, []).append(vals[h])
    else:
        avisos.append("Aba SSA_RMS_SEMANAL não encontrada.")

    # --- INTERIOR
    interior = []
    if "INTERIOR_SEMANAL" in wb.sheetnames:
        ws = wb["INTERIOR_SEMANAL"]
        r0 = _linha_cabecalho(ws, "DATA") or 4
        cab = {}
        for c in range(1, ws.max_column + 1):
            n = _norm(ws.cell(row=r0, column=c).value)
            if not n:
                continue
            if n.startswith("DATA"):
                cab["data"] = c
            elif "DELEGACIA" in n and "TOTAL" in n:
                cab["del_total"] = c
            elif "AUDIENCIA" in n and not n.startswith("%"):
                cab["del_aud"] = c
            elif "INTERNOS" in n:
                cab["up_internos"] = c
            elif "CAPACIDADE" in n and not n.startswith("%"):
                cab["up_capacidade"] = c
            elif "CONJUNTOS" in n:
                cab["up_conjuntos"] = c
            elif n.startswith("REFERENCIA") or n.startswith("FONTE"):
                cab["fonte"] = c
            elif n.startswith("OBSERVA"):
                cab["obs"] = c
        ult_cap, ult_conj = None, None
        for r in range(r0 + 1, ws.max_row + 1):
            bruto = ws.cell(row=r, column=cab.get("data", 1)).value
            reg = {}
            for k in ("del_total", "del_aud", "up_internos", "up_capacidade", "up_conjuntos"):
                v = _num(ws.cell(row=r, column=cab[k]).value) if k in cab else None
                if v == "ERRO":
                    avisos.append(f"INTERIOR_SEMANAL, linha {r}: valor não numérico em {k}. Ignorado.")
                    v = None
                reg[k] = v
            if bruto is None and all(v is None for v in reg.values()):
                continue
            d = _data(bruto)
            if d is None:
                avisos.append(f"INTERIOR_SEMANAL, linha {r}: data inválida ou vazia. Linha ignorada.")
                continue
            # capacidade e nº de conjuntos repetem a última semana se ficarem em branco
            if reg["up_capacidade"] is None:
                reg["up_capacidade"] = ult_cap
            if reg["up_conjuntos"] is None:
                reg["up_conjuntos"] = ult_conj
            ult_cap, ult_conj = reg["up_capacidade"], reg["up_conjuntos"]
            if reg["del_total"] is not None and reg["del_aud"] is not None and reg["del_aud"] > reg["del_total"]:
                avisos.append(f"INTERIOR_SEMANAL, linha {r}: com audiência realizada maior que o total.")
            reg["data"] = d.isoformat()
            reg["fonte"] = str(ws.cell(row=r, column=cab["fonte"]).value or "").strip() if "fonte" in cab else ""
            reg["obs"] = str(ws.cell(row=r, column=cab["obs"]).value or "").strip() if "obs" in cab else ""
            interior.append(reg)
        interior.sort(key=lambda x: x["data"])
    else:
        avisos.append("Aba INTERIOR_SEMANAL não encontrada.")

    wb.close()
    try:
        os.remove(temp)
    except Exception:
        pass

    return {
        "ok": True,
        "arquivo": CAMINHO_PLANILHA,
        "modificado_em": datetime.fromtimestamp(os.path.getmtime(CAMINHO_PLANILHA)).strftime("%d/%m/%Y %H:%M"),
        "unidades": unidades,
        "ssa": ssa,
        "interior": interior,
        "avisos": avisos,
    }


# ---------------------------------------------------------------------------
# CADASTRO (27/11/2026, pedido do usuário): os lançamentos são feitos DENTRO
# do gerador (aba "Lançamentos"), gravados num JSON na pasta da rede com
# trava entre máquinas, revisão (evita sobrescrever o que outra pessoa
# gravou), backups e histórico — o mesmo esquema do Controle de Prazos. A
# planilha antiga só serve para a importação inicial; o Excel agora é só
# exportação (EXPORTAR EXCEL, mesmo layout das abas antigas).
#
# carregar_base() devolve o MESMO formato da leitura da planilha, então os
# slides (JS) não mudaram.
# ---------------------------------------------------------------------------
NOME_DADOS = "SISTEMA PRISIONAL - dados.json"
CAMINHO_DADOS = os.environ.get("CGAPE_PRISIONAL_DADOS") or os.path.join(PASTA_BASE, NOME_DADOS)
NOME_HISTORICO = "SISTEMA PRISIONAL - historico.jsonl"
NOME_BACKUPS = "SISTEMA PRISIONAL - backups"
MAX_BACKUPS = 200
CAMPOS_INTERIOR = ("del_total", "del_aud", "up_internos", "up_capacidade", "up_conjuntos")
LIMITE_TEXTO = 500


def _pasta_dados():
    return os.path.dirname(CAMINHO_DADOS)


def _ler_json(caminho):
    with open(caminho, encoding="utf-8") as f:
        return json.load(f)


def _gravar_json_atomico(caminho, dados):
    fd, tmp = tempfile.mkstemp(prefix=".sp_", suffix=".tmp", dir=os.path.dirname(caminho))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=1)
        os.replace(tmp, caminho)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


class _Trava:
    """Trava entre máquinas (arquivo .lock exclusivo; expira em 30 s)."""
    def __init__(self, espera=10.0):
        self.caminho = CAMINHO_DADOS + ".lock"
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
                    raise TimeoutError("outra pessoa está gravando agora — tente de novo em instantes")
                time.sleep(0.2)

    def __exit__(self, *exc):
        try:
            os.close(self.fd)
        finally:
            try:
                os.remove(self.caminho)
            except OSError:
                pass


def _backup():
    if not os.path.isfile(CAMINHO_DADOS):
        return
    pasta = os.path.join(_pasta_dados(), NOME_BACKUPS)
    try:
        os.makedirs(pasta, exist_ok=True)
        shutil.copy2(CAMINHO_DADOS, os.path.join(pasta, f"dados_{datetime.now():%Y%m%d_%H%M%S_%f}.json"))
        antigos = sorted(f for f in os.listdir(pasta) if f.startswith("dados_"))
        for f in antigos[:-MAX_BACKUPS]:
            os.remove(os.path.join(pasta, f))
    except OSError:
        pass


def _registrar_historico(acao, usuario, chave, antes, depois):
    linha = {"quando": datetime.now().isoformat(timespec="seconds"), "usuario": usuario,
             "acao": acao, "chave": chave, "antes": antes, "depois": depois}
    try:
        with open(os.path.join(_pasta_dados(), NOME_HISTORICO), "a", encoding="utf-8") as f:
            f.write(json.dumps(linha, ensure_ascii=False) + "\n")
    except OSError:
        pass


def importar_da_planilha():
    """Cria o cadastro a partir da planilha antiga (uma vez só)."""
    b = ler_planilha()
    if not b.get("ok"):
        raise RuntimeError(b.get("erro") or "não foi possível ler a planilha")

    def inteiro(v):
        return int(v) if isinstance(v, float) and v.is_integer() else v
    for u in b["unidades"]:
        u["capacidade"] = inteiro(u.get("capacidade"))
    b["ssa"]["series"] = {n: [inteiro(v) for v in vals] for n, vals in b["ssa"]["series"].items()}
    for r in b["interior"]:
        for k in CAMPOS_INTERIOR:
            r[k] = inteiro(r.get(k))
    unidades = [{"nome": u["nome"], "exibicao": u.get("exibicao") or "", "regiao": u.get("regiao") or "",
                 "capacidade": u.get("capacidade"), "exibir": bool(u.get("exibir")), "cor": u.get("cor") or "",
                 "obs": ""} for u in b["unidades"]]
    ssa = []
    for i, d in enumerate(b["ssa"]["datas"]):
        ssa.append({"data": d, "valores": {n: vals[i] for n, vals in b["ssa"]["series"].items()},
                    "fonte": b["ssa"]["fontes"][i], "obs": b["ssa"]["obs"][i]})
    interior = [{k: r.get(k) for k in ("data",) + CAMPOS_INTERIOR + ("fonte", "obs")} for r in b["interior"]]
    dados = {"revisao": 1, "alterado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
             "alterado_por": "importação da planilha", "importado_de": CAMINHO_PLANILHA,
             "unidades": unidades, "ssa": ssa, "interior": interior}
    _gravar_json_atomico(CAMINHO_DADOS, dados)
    return dados


def _carregar_dados():
    if not os.path.isfile(CAMINHO_DADOS):
        with _Trava():
            if not os.path.isfile(CAMINHO_DADOS):
                return importar_da_planilha()
    return _ler_json(CAMINHO_DADOS)


def _chave_un(nome):
    return _norm(nome)


def montar_base(dados):
    """Cadastro -> formato que os slides usam (o mesmo da leitura da planilha)."""
    avisos = []
    unidades = [dict(u) for u in dados.get("unidades", [])]
    nomes = [u["nome"] for u in unidades]
    ssa = {"datas": [], "series": {n: [] for n in nomes}, "fontes": [], "obs": []}
    for reg in sorted(dados.get("ssa", []), key=lambda r: r["data"]):
        ssa["datas"].append(reg["data"])
        ssa["fontes"].append(reg.get("fonte") or "")
        ssa["obs"].append(reg.get("obs") or "")
        vals = reg.get("valores") or {}
        for n in nomes:
            ssa["series"][n].append(vals.get(n))
    interior, ult_cap, ult_conj = [], None, None
    for reg in sorted(dados.get("interior", []), key=lambda r: r["data"]):
        r = dict(reg)
        r["bruto_capacidade"], r["bruto_conjuntos"] = r.get("up_capacidade"), r.get("up_conjuntos")
        # capacidade e nº de conjuntos repetem a semana anterior se ficarem em branco
        if r.get("up_capacidade") is None:
            r["up_capacidade"] = ult_cap
        if r.get("up_conjuntos") is None:
            r["up_conjuntos"] = ult_conj
        ult_cap, ult_conj = r["up_capacidade"], r["up_conjuntos"]
        if r.get("del_total") is not None and r.get("del_aud") is not None and r["del_aud"] > r["del_total"]:
            avisos.append(f"Interior, semana {r['data'][8:10]}/{r['data'][5:7]}/{r['data'][:4]}: "
                          "com audiência realizada maior que o total.")
        r["fonte"], r["obs"] = r.get("fonte") or "", r.get("obs") or ""
        interior.append(r)
    for u in unidades:
        if u.get("capacidade") is None:
            avisos.append(f"Unidade \"{u['nome']}\" sem capacidade cadastrada.")
    try:
        mod = datetime.fromtimestamp(os.path.getmtime(CAMINHO_DADOS)).strftime("%d/%m/%Y %H:%M")
    except OSError:
        mod = dados.get("alterado_em") or ""
    return {"ok": True, "arquivo": CAMINHO_DADOS, "modificado_em": mod, "revisao": dados.get("revisao"),
            "alterado_por": dados.get("alterado_por") or "", "unidades": unidades, "ssa": ssa,
            "interior": interior, "avisos": avisos}


def carregar_base():
    try:
        return montar_base(_carregar_dados())
    except Exception as e:
        return {"ok": False, "erro": f"Não consegui ler o cadastro do Sistema Prisional:\n{CAMINHO_DADOS}\n\n{e}"}


# --- validação dos lançamentos
def _inteiro(v, rotulo, obrigatorio=False):
    if v is None or (isinstance(v, str) and not v.strip()):
        if obrigatorio:
            raise ValueError(f"preencha {rotulo}")
        return None
    if isinstance(v, bool):
        raise ValueError(f"{rotulo}: valor inválido")
    if isinstance(v, (int, float)):
        n = float(v)
    else:
        t = str(v).strip().replace(" ", "").replace(".", "")
        if "," in t:
            t = t.replace(",", ".")
        try:
            n = float(t)
        except ValueError:
            raise ValueError(f"{rotulo}: use só números (ex.: 5.657)")
    if n < 0:
        raise ValueError(f"{rotulo}: não pode ser negativo")
    return int(round(n))


def _data_iso(v):
    try:
        return date.fromisoformat(str(v or "")[:10]).isoformat()
    except ValueError:
        raise ValueError("data de referência inválida")


def _texto(v):
    return " ".join(str(v or "").split())[:LIMITE_TEXTO]


def _alterar(revisao_tela, usuario, funcao):
    """Aplica funcao(dados) -> (acao, chave, antes, depois) sob a trava,
    conferindo a revisão; devolve a resposta com a base nova."""
    if not os.path.isdir(_pasta_dados()):
        return {"ok": False, "erro": "Pasta do Sistema Prisional inacessível (rede fora do ar)."}
    try:
        with _Trava():
            dados = _carregar_dados()
            if revisao_tela is not None and dados.get("revisao") != revisao_tela:
                return {"ok": False, "conflito": True, "base": montar_base(dados),
                        "erro": "Os dados foram alterados em outra máquina desde que a tela foi aberta. "
                                "A tela foi atualizada — confira e grave de novo."}
            acao, chave, antes, depois = funcao(dados)
            _backup()
            dados["revisao"] = int(dados.get("revisao") or 0) + 1
            dados["alterado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            dados["alterado_por"] = usuario
            _gravar_json_atomico(CAMINHO_DADOS, dados)
            _registrar_historico(acao, usuario, chave, antes, depois)
        return {"ok": True, "base": montar_base(dados)}
    except (ValueError, TimeoutError, OSError) as e:
        return {"ok": False, "erro": str(e)}


def _seguro(f):
    """Erro de validação (ValueError etc.) vira {"ok": False, "erro": ...}."""
    def g(*a, **k):
        try:
            return f(*a, **k)
        except (ValueError, TimeoutError, OSError) as e:
            return {"ok": False, "erro": str(e)}
    g.__name__ = f.__name__
    return g


@_seguro
def salvar_ssa(reg, data_original, revisao, usuario):
    reg = reg or {}
    novo = {"data": _data_iso(reg.get("data")), "fonte": _texto(reg.get("fonte")), "obs": _texto(reg.get("obs"))}

    def aplicar(dados):
        nomes = [u["nome"] for u in dados.get("unidades", [])]
        vals = reg.get("valores") or {}
        novo["valores"] = {n: _inteiro(vals.get(n), n) for n in nomes}
        if all(v is None for v in novo["valores"].values()):
            raise ValueError("preencha o número de custodiados de pelo menos uma unidade")
        lst = dados.setdefault("ssa", [])
        orig = str(data_original or "")[:10]
        if any(r["data"] == novo["data"] and r["data"] != orig for r in lst):
            raise ValueError(f"já existe lançamento de Salvador e RMS na data {novo['data'][8:10]}/{novo['data'][5:7]}/{novo['data'][:4]}")
        for i, r in enumerate(lst):
            if orig and r["data"] == orig:
                lst[i] = novo
                return "ssa.editar", novo["data"], r, novo
        if orig:
            raise ValueError("semana não encontrada (pode ter sido excluída em outra máquina)")
        lst.append(novo)
        return "ssa.incluir", novo["data"], None, novo
    return _alterar(revisao, usuario, aplicar)


@_seguro
def salvar_interior(reg, data_original, revisao, usuario):
    reg = reg or {}
    novo = {"data": _data_iso(reg.get("data"))}
    rotulos = {"del_total": "delegacias - total de custodiados", "del_aud": "delegacias - com audiência realizada",
               "up_internos": "unidades prisionais - total de internos",
               "up_capacidade": "unidades prisionais - capacidade", "up_conjuntos": "total de conjuntos prisionais"}
    for k in CAMPOS_INTERIOR:
        novo[k] = _inteiro(reg.get(k), rotulos[k])
    if all(novo[k] is None for k in CAMPOS_INTERIOR):
        raise ValueError("preencha pelo menos um dos números da semana")
    if novo["del_total"] is not None and novo["del_aud"] is not None and novo["del_aud"] > novo["del_total"]:
        raise ValueError("com audiência realizada não pode ser maior que o total de custodiados nas delegacias")
    novo["fonte"], novo["obs"] = _texto(reg.get("fonte")), _texto(reg.get("obs"))

    def aplicar(dados):
        lst = dados.setdefault("interior", [])
        orig = str(data_original or "")[:10]
        if any(r["data"] == novo["data"] and r["data"] != orig for r in lst):
            raise ValueError(f"já existe lançamento do Interior na data {novo['data'][8:10]}/{novo['data'][5:7]}/{novo['data'][:4]}")
        for i, r in enumerate(lst):
            if orig and r["data"] == orig:
                lst[i] = novo
                return "interior.editar", novo["data"], r, novo
        if orig:
            raise ValueError("semana não encontrada (pode ter sido excluída em outra máquina)")
        lst.append(novo)
        return "interior.incluir", novo["data"], None, novo
    return _alterar(revisao, usuario, aplicar)


@_seguro
def excluir_semana(tipo, data, revisao, usuario):
    if tipo not in ("ssa", "interior"):
        return {"ok": False, "erro": "tipo inválido"}

    def aplicar(dados):
        lst = dados.setdefault(tipo, [])
        for i, r in enumerate(lst):
            if r["data"] == str(data)[:10]:
                lst.pop(i)
                return tipo + ".excluir", r["data"], r, None
        raise ValueError("semana não encontrada (pode ter sido excluída em outra máquina)")
    return _alterar(revisao, usuario, aplicar)


@_seguro
def salvar_unidade(reg, nome_original, revisao, usuario):
    reg = reg or {}
    nome = " ".join(str(reg.get("nome") or "").split()).upper()
    if not nome:
        return {"ok": False, "erro": "preencha o nome da unidade"}
    cor = str(reg.get("cor") or "").strip()
    if cor and not (len(cor.lstrip("#")) == 6 and all(c in "0123456789abcdefABCDEF" for c in cor.lstrip("#"))):
        return {"ok": False, "erro": "cor inválida (use o seletor de cor)"}
    try:
        novo = {"nome": nome[:200], "exibicao": _texto(reg.get("exibicao"))[:200],
                "regiao": _texto(reg.get("regiao")) or "SALVADOR E RMS",
                "capacidade": _inteiro(reg.get("capacidade"), "capacidade"),
                "exibir": bool(reg.get("exibir")), "cor": ("#" + cor.lstrip("#").upper()) if cor else "",
                "obs": _texto(reg.get("obs"))}
    except ValueError as e:
        return {"ok": False, "erro": str(e)}

    def aplicar(dados):
        lst = dados.setdefault("unidades", [])
        orig = str(nome_original or "")
        if any(_chave_un(u["nome"]) == _chave_un(nome) and u["nome"] != orig for u in lst):
            raise ValueError("já existe uma unidade com esse nome")
        for i, u in enumerate(lst):
            if orig and u["nome"] == orig:
                if nome != orig:
                    # renomeou: leva os números das semanas junto
                    for r in dados.get("ssa", []):
                        v = r.get("valores") or {}
                        if orig in v:
                            v[nome] = v.pop(orig)
                lst[i] = novo
                return "unidade.editar", nome, u, novo
        if orig:
            raise ValueError("unidade não encontrada (pode ter sido excluída em outra máquina)")
        lst.append(novo)
        return "unidade.incluir", nome, None, novo
    return _alterar(revisao, usuario, aplicar)


@_seguro
def excluir_unidade(nome, revisao, usuario):
    def aplicar(dados):
        lst = dados.setdefault("unidades", [])
        for i, u in enumerate(lst):
            if u["nome"] == nome:
                lst.pop(i)
                valores = {}
                for r in dados.get("ssa", []):
                    v = (r.get("valores") or {}).pop(nome, None)
                    if v is not None:
                        valores[r["data"]] = v
                return "unidade.excluir", nome, dict(u, valores_removidos=valores), None
        raise ValueError("unidade não encontrada (pode ter sido excluída em outra máquina)")
    return _alterar(revisao, usuario, aplicar)


def excel_bytes():
    """Excel só de consulta, com as mesmas abas/cabeçalhos (linha 4) da planilha antiga."""
    from openpyxl.styles import Alignment, Font, PatternFill
    b = montar_base(_carregar_dados())
    wb = openpyxl.Workbook()
    cab_fill, cab_font = PatternFill("solid", fgColor="1F3A5F"), Font(bold=True, color="FFFFFF")

    def aba(titulo, sub, cab, linhas, larguras, fmt_data=True):
        ws = wb.create_sheet(titulo)
        ws["A1"], ws["A2"] = titulo.replace("_", " "), sub
        ws["A1"].font = Font(bold=True, size=13)
        ws.append([])
        ws.append(cab)
        for c in ws[4]:
            c.fill, c.font = cab_fill, cab_font
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for l in linhas:
            ws.append(l)
            if fmt_data:
                ws.cell(ws.max_row, 1).number_format = "DD/MM/YYYY"
        for i, w in enumerate(larguras, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.freeze_panes = "A5"
    nomes = [u["nome"] for u in b["unidades"]]
    aba("UNIDADES", "Cadastro de unidades (exportado do CGAPE).",
        ["UNIDADE", "NOME NO GRÁFICO", "REGIÃO", "CAPACIDADE DE LOTAÇÃO (VAGAS)", "EXIBIR NO PAINEL", "COR", "OBSERVAÇÃO"],
        [[u["nome"], u.get("exibicao"), u.get("regiao"), u.get("capacidade"), "SIM" if u.get("exibir") else "NÃO",
          u.get("cor"), u.get("obs")] for u in b["unidades"]], (38, 34, 18, 18, 14, 10, 40), fmt_data=False)
    aba("SSA_RMS_SEMANAL", "Total de custodiados por unidade. Uma linha por semana.",
        ["DATA DE REFERÊNCIA"] + nomes + ["FONTE", "OBSERVAÇÃO"],
        [[date.fromisoformat(d)] + [b["ssa"]["series"][n][i] for n in nomes] + [b["ssa"]["fontes"][i], b["ssa"]["obs"][i]]
         for i, d in enumerate(b["ssa"]["datas"])], [16] + [22] * len(nomes) + [40, 40])
    aba("INTERIOR_SEMANAL", "Custodiados em delegacias e lotação das unidades prisionais do Interior.",
        ["DATA DE REFERÊNCIA", "DELEGACIAS - TOTAL DE CUSTODIADOS", "DELEGACIAS - COM AUDIÊNCIA REALIZADA",
         "UNIDADES PRISIONAIS - TOTAL DE INTERNOS", "UNIDADES PRISIONAIS - CAPACIDADE DE LOTAÇÃO",
         "TOTAL DE CONJUNTOS PRISIONAIS", "REFERÊNCIA (FONTE)", "OBSERVAÇÃO"],
        [[date.fromisoformat(r["data"]), r.get("del_total"), r.get("del_aud"), r.get("up_internos"),
          r.get("bruto_capacidade"), r.get("bruto_conjuntos"), r.get("fonte"), r.get("obs")] for r in b["interior"]],
        (16, 18, 18, 18, 18, 14, 40, 40))
    del wb[wb.sheetnames[0]]
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# EXPORTAÇÃO
# ---------------------------------------------------------------------------
def _png_de_dataurl(dataurl):
    return base64.b64decode(dataurl.split(",", 1)[1])


def gerar_pptx_bytes(imagens):
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for img in imagens:
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_picture(io.BytesIO(_png_de_dataurl(img["png"])), 0, 0,
                             width=prs.slide_width, height=prs.slide_height)
        if img.get("nota"):
            s.notes_slide.notes_text_frame.text = img["nota"]
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def copiar_png_clipboard(png_bytes):
    """Coloca a imagem na área de transferência do Windows (formato DIB)."""
    import ctypes
    from ctypes import wintypes
    from PIL import Image

    im = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    buf = io.BytesIO()
    im.save(buf, "BMP")
    dib = buf.getvalue()[14:]  # sem o cabeçalho BITMAPFILEHEADER

    k32, u32 = ctypes.windll.kernel32, ctypes.windll.user32
    k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    k32.GlobalAlloc.restype = wintypes.HGLOBAL
    k32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    k32.GlobalLock.restype = wintypes.LPVOID
    k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    u32.OpenClipboard.argtypes = [wintypes.HWND]
    u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    u32.SetClipboardData.restype = wintypes.HANDLE

    h = k32.GlobalAlloc(0x0002, len(dib))
    p = k32.GlobalLock(h)
    ctypes.memmove(p, dib, len(dib))
    k32.GlobalUnlock(h)
    for _ in range(10):
        if u32.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        raise RuntimeError("A área de transferência está ocupada por outro programa.")
    try:
        u32.EmptyClipboard()
        if not u32.SetClipboardData(8, h):  # CF_DIB
            raise RuntimeError("Falha ao copiar a imagem.")
    finally:
        u32.CloseClipboard()


# ---------------------------------------------------------------------------
# PONTE JS <-> PYTHON (janela desktop)
# ---------------------------------------------------------------------------
janela = None


def _dialogo_salvar(nome, tipo):
    pasta = PASTA_DOWNLOADS if os.path.isdir(PASTA_DOWNLOADS) else PASTA_BASE
    modo = getattr(getattr(webview, "FileDialog", None), "SAVE", None) or webview.SAVE_DIALOG
    res = janela.create_file_dialog(modo, directory=pasta, save_filename=nome, file_types=(tipo,))
    if not res:
        return None
    caminho = res[0] if isinstance(res, (list, tuple)) else res
    ext = os.path.splitext(nome)[1]
    if not caminho.lower().endswith(ext):
        caminho += ext
    return caminho


def _abrir(caminho):
    # os.startfile chamado da thread da API do pywebview derruba o app (falha no SHELL32);
    # o explorer.exe abre o arquivo/pasta no programa padrão, em processo separado.
    subprocess.Popen(["explorer.exe", os.path.normpath(caminho)])


class API:
    """API da janela. `permitido()` (do CGAPE) é conferido em toda chamada:
    se a sessão mudar (saiu/trocou de usuário), a janela para de responder."""
    def __init__(self, permitido=None, usuario=None):
        self._permitido = permitido or (lambda: True)
        # quem está gravando (para o histórico): função do CGAPE -> "login (Nome)"
        self._usuario = usuario or (lambda: "")

    def __getattribute__(self, nome):
        attr = object.__getattribute__(self, nome)
        if nome.startswith("_") or not callable(attr):
            return attr
        permitido = object.__getattribute__(self, "_permitido")

        def protegido(*a, **k):
            if not permitido():
                return {"ok": False, "erro": "Acesso restrito: o Sistema Prisional é só para Silvana e Diego."}
            return attr(*a, **k)
        return protegido

    def carregar(self):
        try:
            return carregar_base()
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"ok": False, "erro": str(e)}

    def versao_planilha(self):
        # data de modificação do CADASTRO (desde 27/11/2026; antes, do xlsx):
        # o painel consulta isto para se atualizar sozinho quando outra
        # máquina grava um lançamento
        try:
            return os.path.getmtime(CAMINHO_DADOS)
        except OSError:
            return None

    # --- lançamentos (27/11/2026)
    def salvar_ssa(self, reg=None, data_original=None, revisao=None):
        return salvar_ssa(reg, data_original, revisao, self._usuario())

    def salvar_interior(self, reg=None, data_original=None, revisao=None):
        return salvar_interior(reg, data_original, revisao, self._usuario())

    def excluir_semana(self, tipo=None, data=None, revisao=None):
        return excluir_semana(tipo, data, revisao, self._usuario())

    def salvar_unidade(self, reg=None, nome_original=None, revisao=None):
        return salvar_unidade(reg, nome_original, revisao, self._usuario())

    def excluir_unidade(self, nome=None, revisao=None):
        return excluir_unidade(nome, revisao, self._usuario())

    def exportar_excel(self):
        try:
            nome = f"Sistema Prisional - base - {datetime.now():%d-%m-%Y %Hh%M}.xlsx"
            pasta = PASTA_DOWNLOADS if os.path.isdir(PASTA_DOWNLOADS) else tempfile.gettempdir()
            caminho = os.path.join(pasta, nome)
            with open(caminho, "wb") as f:
                f.write(excel_bytes())
            _abrir(caminho)
            return {"ok": True, "arquivo": caminho}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def salvar_png(self, nome, dataurl):
        try:
            caminho = _dialogo_salvar(nome, "Imagem PNG (*.png)")
            if not caminho:
                return {"ok": False, "cancelado": True}
            with open(caminho, "wb") as f:
                f.write(_png_de_dataurl(dataurl))
            return {"ok": True, "arquivo": caminho}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def copiar_png(self, dataurl):
        try:
            copiar_png_clipboard(_png_de_dataurl(dataurl))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def gerar_pptx(self, nome, imagens):
        try:
            caminho = _dialogo_salvar(nome, "Apresentação PowerPoint (*.pptx)")
            if not caminho:
                return {"ok": False, "cancelado": True}
            with open(caminho, "wb") as f:
                f.write(gerar_pptx_bytes(imagens))
            if os.name == "nt":
                try:
                    _abrir(caminho)
                except Exception:
                    pass
            return {"ok": True, "arquivo": caminho}
        except PermissionError:
            return {"ok": False, "erro": "O arquivo está aberto no PowerPoint. Feche e tente de novo."}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def abrir_planilha(self):
        try:
            _abrir(CAMINHO_PLANILHA)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def abrir_link(self, url):
        # só o link da SEAP do cabeçalho; abre no navegador padrão, em processo separado
        if url != LINK_SEAP:
            return {"ok": False, "erro": "Link não permitido."}
        try:
            subprocess.Popen(["explorer.exe", url])
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}

    def abrir_pasta(self):
        try:
            _abrir(os.path.dirname(CAMINHO_PLANILHA))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "erro": str(e)}


# ---------------------------------------------------------------------------
# INTERFACE (HTML/CSS/JS)
# ---------------------------------------------------------------------------
HTML = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Sistema Prisional - Gráficos semanais</title>
<style>
:root{
  --bg:#EEF1F5; --painel:#FFFFFF; --linha:#D8DEE6; --texto:#1F2A3D; --suave:#5B6676;
  --azul:#2B4C8C; --azul-esc:#1F3A5F; --ouro:#B8902A; --vermelho:#B3261E; --ok:#2E7D4F;
}
*{box-sizing:border-box}
html,body{margin:0;height:100%;background:var(--bg);color:var(--texto);
  font-family:"Segoe UI",Calibri,Arial,sans-serif;font-size:14px}
header{display:flex;align-items:center;gap:16px;padding:12px 20px;background:var(--azul-esc);color:#fff}
header .marca{font-weight:700;font-size:16px;letter-spacing:.2px}
header .marca small{display:block;font-weight:400;font-size:12px;opacity:.75;margin-top:2px}
header .acoes{margin-left:auto;display:flex;gap:8px}
header .link-seap{margin-left:auto;color:#E3C36A;font-size:13px;text-decoration:none;white-space:nowrap;
  border-bottom:1px dotted rgba(227,195,106,.6)}
header .link-seap:hover{color:#fff;border-bottom-color:#fff}
header .link-seap + .acoes{margin-left:8px}
nav{display:flex;gap:4px;padding:0 20px;background:var(--azul-esc)}
nav button{background:transparent;border:0;color:#C9D3E3;padding:10px 16px;font:inherit;cursor:pointer;
  border-bottom:3px solid transparent}
nav button.ativo{color:#fff;border-bottom-color:var(--ouro);font-weight:600}
nav button:hover{color:#fff}
.btn{border:1px solid var(--linha);background:#fff;color:var(--texto);padding:7px 14px;border-radius:6px;
  font:inherit;cursor:pointer;white-space:nowrap}
.btn:hover{border-color:var(--azul)}
.btn.prim{background:var(--azul);border-color:var(--azul);color:#fff}
.btn.prim:hover{background:#23407A}
header .btn{background:rgba(255,255,255,.08);border-color:rgba(255,255,255,.25);color:#fff}
header .btn:hover{background:rgba(255,255,255,.18)}
header .btn.prim{background:var(--ouro);border-color:var(--ouro)}
main{display:grid;grid-template-columns:300px 1fr;gap:16px;padding:16px 20px;height:calc(100% - 100px)}
aside{background:var(--painel);border:1px solid var(--linha);border-radius:8px;padding:16px;overflow:auto}
aside h3{margin:18px 0 8px;font-size:12px;letter-spacing:1px;color:var(--suave);text-transform:uppercase}
aside h3:first-child{margin-top:0}
aside label{display:flex;align-items:center;gap:8px;margin:6px 0;cursor:pointer}
aside select,aside input[type=text]{width:100%;padding:6px 8px;border:1px solid var(--linha);border-radius:5px;font:inherit}
.seg{display:flex;border:1px solid var(--linha);border-radius:6px;overflow:hidden}
.seg button{flex:1;border:0;background:#fff;padding:7px;font:inherit;cursor:pointer;color:var(--suave)}
.seg button.ativo{background:var(--azul);color:#fff}
.linha2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.linha2 span{font-size:12px;color:var(--suave)}
.palco{display:flex;flex-direction:column;gap:10px;min-width:0;overflow:auto}
.barra{display:flex;gap:8px;align-items:center}
.barra .info{color:var(--suave);font-size:12px;margin-right:auto}
.slide{background:#fff;border:1px solid var(--linha);border-radius:4px;box-shadow:0 6px 24px rgba(31,42,61,.10)}
.slide svg{display:block;width:100%;height:auto}
.dados{background:var(--painel);border:1px solid var(--linha);border-radius:8px;padding:16px;overflow:auto}
table{border-collapse:collapse;width:100%;margin-bottom:22px;font-size:13px}
th{background:#F3F5F8;color:var(--suave);text-align:left;font-weight:600;padding:7px 10px;border-bottom:1px solid var(--linha)}
td{padding:6px 10px;border-bottom:1px solid #EEF1F5}
td.n{text-align:right;font-variant-numeric:tabular-nums}
.aviso{background:#FFF6E0;border:1px solid #F0D48A;border-radius:6px;padding:10px 14px;margin-bottom:14px}
.aviso li{margin:3px 0}
.vazio{padding:40px;text-align:center;color:var(--suave)}
#toast{position:fixed;right:20px;bottom:20px;background:var(--texto);color:#fff;padding:10px 16px;border-radius:6px;
  opacity:0;transform:translateY(10px);transition:.2s;pointer-events:none;max-width:480px}
#toast.on{opacity:1;transform:none}
#toast.erro{background:var(--vermelho)}
.oculto{display:none!important}
.dica{font-size:12px;color:var(--suave);line-height:1.4}
/* lançamentos dentro do gerador (27/11/2026) */
.sec-cab{display:flex;align-items:center;gap:12px;margin:22px 0 8px}
.sec-cab:first-of-type{margin-top:6px}
.sec-cab h3{margin:0;font-size:13px;letter-spacing:1px;color:var(--suave);text-transform:uppercase}
.sec-cab .btn{margin-left:auto}
.btn.mini{padding:3px 10px;font-size:12px}
.btn.perigo{color:var(--vermelho)}
.btn.perigo:hover{border-color:var(--vermelho)}
td.acoes-linha{white-space:nowrap;text-align:right;width:1%}
tr:hover td{background:#F7F9FC}
.cor-amostra{display:inline-block;width:14px;height:14px;border-radius:3px;vertical-align:-2px;border:1px solid #0002}
.modal{position:fixed;inset:0;background:rgba(15,25,40,.55);display:flex;align-items:center;justify-content:center;z-index:50;padding:20px}
.modal-caixa{background:#fff;border-radius:10px;width:min(720px,100%);max-height:100%;display:flex;flex-direction:column;
  box-shadow:0 20px 60px rgba(0,0,0,.35);border-top:4px solid var(--azul)}
.modal-caixa h2{margin:0;padding:16px 20px 6px;font-size:18px}
.modal-caixa .sub{padding:0 20px;color:var(--suave);font-size:12.5px}
.form-grade{display:grid;grid-template-columns:1fr 1fr;gap:12px 16px;padding:14px 20px;overflow:auto}
.form-grade label{display:flex;flex-direction:column;gap:4px;font-size:11.5px;font-weight:600;letter-spacing:.3px;
  text-transform:uppercase;color:var(--suave)}
.form-grade label.largo{grid-column:1/-1}
.form-grade label.check{flex-direction:row;align-items:center;gap:8px;text-transform:none;font-size:13px;color:var(--texto)}
.form-grade input,.form-grade textarea{font:inherit;font-size:14px;text-transform:none;letter-spacing:0;font-weight:400;
  color:var(--texto);border:1px solid var(--linha);border-radius:6px;padding:7px 9px;outline:none}
.form-grade input:focus,.form-grade textarea:focus{border-color:var(--azul);box-shadow:0 0 0 3px rgba(43,76,140,.15)}
.form-grade input[readonly]{background:#F3F5F8;color:var(--suave)}
.form-grade input.num{text-align:right;font-variant-numeric:tabular-nums}
.form-grade input[type=color]{padding:2px;height:36px;width:80px}
.form-grade .grupo-tit{grid-column:1/-1;margin:6px 0 -4px;font-size:12px;font-weight:700;color:var(--azul);letter-spacing:.5px}
.form-erro{color:var(--vermelho);font-weight:600;padding:0 20px;min-height:4px}
.form-rod{display:flex;gap:8px;justify-content:flex-end;padding:12px 20px 16px;border-top:1px solid var(--linha)}
</style>
</head>
<body>
<header>
  <div class="marca">Sistema Prisional · Gráficos semanais<small id="infoBase">Carregando base...</small></div>
  <a class="link-seap" href="https://www.ba.gov.br/seap/publicacoes/dadosestatisticos" target="_blank" rel="noopener"
     onclick="return abrirLinkSeap(this.href)" title="Abre no navegador a página de dados estatísticos da SEAP">Dados estatísticos SEAP ↗</a>
  <div class="acoes">
    <button class="btn" onclick="recarregar()" title="Relê a planilha (use depois de salvar a semana nova)">Recarregar base</button>
    <button class="btn so-desktop" onclick="exportarExcel()" title="Baixa um Excel com todos os lançamentos (só consulta: a base é o próprio sistema)">Exportar Excel</button>
    <button class="btn prim so-desktop" onclick="exportarPptx()">Gerar PPTX completo</button>
  </div>
</header>
<nav>
  <button data-aba="fluxo" class="ativo">Fluxo Salvador e RMS</button>
  <button data-aba="evolucao">Interior · evolução</button>
  <button data-aba="dados">Lançamentos</button>
</nav>

<main id="main">
  <aside id="controles"></aside>
  <section class="palco">
    <div class="barra">
      <span class="info" id="infoSlide"></span>
      <button class="btn so-desktop" onclick="copiarSlide()">Copiar imagem</button>
      <button class="btn prim" onclick="salvarSlide()">Salvar PNG</button>
    </div>
    <div class="slide" id="slide"></div>
  </section>
</main>
<section class="dados oculto" id="dados" style="margin:16px 20px"></section>
<div id="toast"></div>
<div class="modal oculto" id="modal">
  <form class="modal-caixa" id="formReg" onsubmit="return salvarForm(event)" novalidate>
    <h2 id="formTit"></h2>
    <div class="sub" id="formSub"></div>
    <div class="form-grade" id="formCampos"></div>
    <div class="form-erro" id="formErro" role="alert"></div>
    <div class="form-rod">
      <button type="button" class="btn" onclick="fecharForm()">Cancelar</button>
      <button type="submit" class="btn prim" id="formSalvar">Salvar</button>
    </div>
  </form>
</div>

<script>
window.DADOS_INICIAIS = /*__DADOS__*/null;
const DESKTOP = () => !!(window.pywebview && window.pywebview.api);
let BASE = null;
let ABA = "fluxo";
const W = 1920, H = 1080;
const FONTE = "'Segoe UI', Calibri, Arial, sans-serif";
const CORES_PADRAO = ["#2B4C8C", "#B8902A", "#7A4E9A", "#2E7D6B", "#B3561E"];

// estado dos controles (preferências guardadas só neste computador)
let EST = {
  fluxo: {de: null, ate: null, metrica: "abs", unidades: null,
          kicker: "SALVADOR E RMS · GRÁFICO DE ACOMPANHAMENTO DE FLUXO",
          titulo: "Fluxo Semanal de Transferência de Pessoas Privadas de Liberdade – Audiência de Custódia"},
  interior: {data: null,
          kicker: "INTERIOR DO ESTADO · QUANTITATIVOS",
          titulo: "Custodiados em Delegacias - Transferência para unidades prisionais",
          obs: "OBS: A recepção dos custodiados não segue o mesmo fluxo diário em todas as unidades prisionais, variando conforme gestão de operacionalização e capacidade de efetivo."},
  evolucao: {de: null, ate: null, verTotal: true, verAud: false,
          kicker: "INTERIOR DO ESTADO · EVOLUÇÃO SEMANAL",
          titulo: "Custodiados em Delegacias e lotação das Unidades Prisionais"}
};
try {
  const salvo = JSON.parse(localStorage.getItem("sp_textos") || "null");
  if (salvo) for (const k of ["fluxo","interior","evolucao"])
    for (const c of ["kicker","titulo","obs"]) if (salvo[k] && salvo[k][c]) EST[k][c] = salvo[k][c];
} catch(e) {}
function guardarTextos(){
  try {
    const t = {};
    for (const k of ["fluxo","interior","evolucao"]) t[k] = {kicker:EST[k].kicker, titulo:EST[k].titulo, obs:EST[k].obs};
    localStorage.setItem("sp_textos", JSON.stringify(t));
  } catch(e) {}
}

// ---------------- formatação
const nf0 = new Intl.NumberFormat("pt-BR", {maximumFractionDigits: 0});
const fInt = v => v == null ? "-" : nf0.format(Math.round(v));
const fPct = (v, d = 1) => v == null || !isFinite(v) ? "-" :
  (v * 100).toLocaleString("pt-BR", {minimumFractionDigits: d, maximumFractionDigits: d}) + "%";
const fDM = iso => iso.slice(8,10) + "/" + iso.slice(5,7);
const fDMA = iso => fDM(iso) + "/" + iso.slice(0,4);
const esc = s => String(s ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
// mede o texto com a fonte real da máquina (svg oculto)
let _med = null;
function larguraTxt(txt, fs, peso = 400){
  try {
    if (!_med) {
      const ns = "http://www.w3.org/2000/svg", sv = document.createElementNS(ns, "svg");
      sv.setAttribute("style", "position:absolute;left:-9999px;top:-9999px;width:10px;height:10px");
      _med = document.createElementNS(ns, "text"); _med.setAttribute("font-family", FONTE);
      sv.appendChild(_med); document.body.appendChild(sv);
    }
    _med.setAttribute("font-size", fs); _med.setAttribute("font-weight", peso);
    _med.textContent = String(txt);
    const l = _med.getComputedTextLength();
    if (l > 0) return l;
  } catch(e) {}
  return String(txt).length * fs * (peso >= 600 ? 0.58 : 0.53);
}
// maior tamanho de fonte (até fsMax) que faz o texto caber em maxW
const caber = (txt, fsMax, maxW, peso = 400, fsMin = 14) => {
  const l = larguraTxt(txt, fsMax, peso);
  return Math.max(fsMin, l <= maxW ? fsMax : Math.floor(fsMax * maxW / l));
};
const semTags = h => h.replace(/<[^>]+>/g, "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">");

function toast(msg, erro = false){
  const t = document.getElementById("toast");
  t.textContent = msg; t.className = "on" + (erro ? " erro" : "");
  clearTimeout(t._h); t._h = setTimeout(() => t.className = "", erro ? 6000 : 3000);
}

// link da SEAP no cabeçalho: no app abre no navegador padrão (via Python); no navegador, segue o href
function abrirLinkSeap(url){
  if (!DESKTOP()) return true;
  chamar("abrir_link", url);
  return false;
}

async function chamar(nome, ...args){
  if (!DESKTOP()) { toast("Disponível apenas no aplicativo (.exe).", true); return null; }
  const r = await window.pywebview.api[nome](...args);
  if (r && r.ok === false && !r.cancelado) toast(r.erro || "Falha.", true);
  return r;
}

// ---------------- escala do eixo
function escalaBonita(min, max, nTicks = 4){
  if (min === max){ min -= 1; max += 1; }
  const bruto = (max - min) / nTicks;
  const mag = Math.pow(10, Math.floor(Math.log10(bruto)));
  const passo = [1, 2, 2.5, 5, 10].map(m => m * mag).find(p => p >= bruto);
  const ini = Math.floor(min / passo) * passo, fim = Math.ceil(max / passo) * passo;
  const ticks = [];
  for (let v = ini; v <= fim + passo / 1e6; v += passo) ticks.push(+v.toFixed(10));
  return {ini, fim, ticks, passo};
}

let _gid = 0;
// Gráfico de linhas em SVG.
// series: [{nome, cor, valores:[n|null], rotuloAbaixo:bool}]
function graficoLinhas({x, y, w, h, cats, series, fmt, fmtEixo, ref, tituloY = "", tituloX = "Semana", area = true, legenda = false}){
  const id = "g" + (++_gid);
  const pL = 26, pR = 96, pT = 58;
  const px = x + pL, py = y + pT, pw = w - pL - pR;
  const todos = series.flatMap(s => s.valores).filter(v => v != null);
  if (!todos.length) return `<text x="${x + w/2}" y="${y + h/2}" text-anchor="middle" font-size="26" fill="#8A94A3">Sem dados no período</text>`;
  const n = cats.length;
  const inset = n > 1 ? Math.min(46, pw / (n * 2)) : pw / 2;
  const espX = n > 1 ? (pw - 2 * inset) / (n - 1) : pw;
  // eixo X: se as datas não cabem lado a lado, inclina; se ainda faltar espaço, pula algumas
  const fsX = 23, largX = Math.max(...cats.map(c => larguraTxt(c, fsX)));
  const inclinar = largX + 12 > espX;
  const ANG = 40, sen = Math.sin(ANG * Math.PI / 180), cos = Math.cos(ANG * Math.PI / 180);
  const passoX = inclinar ? Math.max(1, Math.ceil(fsX * 1.5 / espX)) : 1;
  const altX = inclinar ? largX * sen + fsX * cos : fsX;
  const pB = inclinar ? Math.round(84 + altX) : 92;
  const ph = h - pT - pB;
  let mn = Math.min(...todos), mx = Math.max(...todos);
  if (ref && ref.valor != null) { mn = Math.min(mn, ref.valor); mx = Math.max(mx, ref.valor); }
  const folga = (mx - mn || Math.abs(mx) || 1) * 0.14;
  const esc_ = escalaBonita(mn - folga, mx + folga, 4);
  let lo = Math.max(esc_.ini, mn - folga), hi = Math.min(esc_.fim, mx + folga * 1.6);
  if (mn >= 0 && lo < 0) lo = 0;
  const Y = v => py + ph - (v - lo) / (hi - lo) * ph;
  const X = i => n > 1 ? px + inset + i * espX : px + pw / 2;
  let s = "";
  // moldura tracejada
  s += `<rect x="${px}" y="${py}" width="${pw}" height="${ph}" rx="12" fill="none" stroke="#E1E5EB" stroke-width="2" stroke-dasharray="6 6"/>`;
  // grade e eixo Y (à direita)
  for (const t of esc_.ticks) {
    if (t < lo - 1e-9 || t > hi + 1e-9) continue;
    s += `<line x1="${px}" x2="${px + pw}" y1="${Y(t)}" y2="${Y(t)}" stroke="#E6E9EE" stroke-width="2" stroke-dasharray="6 6"/>`;
    s += `<text x="${px + pw + 16}" y="${Y(t) + 8}" font-size="23" fill="#4A5568">${esc((fmtEixo || fmt)(t))}</text>`;
  }
  if (tituloY) s += `<text x="${x + w}" y="${y + 24}" text-anchor="end" font-size="22" fill="#4A5568">${esc(tituloY)}</text>`;
  // grade vertical e eixo X
  cats.forEach((c, i) => {
    s += `<line x1="${X(i)}" x2="${X(i)}" y1="${py}" y2="${py + ph}" stroke="#EEF0F3" stroke-width="2" stroke-dasharray="6 6"/>`;
    // a última data sempre aparece; a anterior a ela some se ficar colada
    if (!(i % passoX === 0 || i === n - 1) || (i !== n - 1 && (n - 1 - i) < passoX && passoX > 1)) return;
    if (inclinar) {
      const ty = py + ph + 22 + fsX * 0.7;
      s += `<text x="${X(i) + 6}" y="${ty}" transform="rotate(-${ANG} ${X(i) + 6} ${ty})" text-anchor="end" font-size="${fsX}" fill="#2D3748">${esc(c)}</text>`;
    } else
      s += `<text x="${X(i)}" y="${py + ph + 40}" text-anchor="middle" font-size="${fsX}" fill="#2D3748">${esc(c)}</text>`;
  });
  if (tituloX) s += `<text x="${px + pw / 2}" y="${py + ph + (inclinar ? 58 + altX : 82)}" text-anchor="middle" font-size="22" fill="#4A5568">${esc(tituloX)}</text>`;
  const ocupados = [];  // caixas de rótulo já desenhadas, para evitar sobreposição
  const colide = (a) => ocupados.some(b => a.x < b.x + b.w + 4 && b.x < a.x + a.w + 4 && a.y < b.y + b.h + 2 && b.y < a.y + a.h + 2);
  // linha de referência (capacidade); o texto vai para o canto com menos pontos/linhas por perto
  let textoRef = "";
  if (ref && ref.valor != null && ref.valor >= lo && ref.valor <= hi) {
    const yr = Y(ref.valor);
    s += `<line x1="${px}" x2="${px + pw}" y1="${yr}" y2="${yr}" stroke="#B3261E" stroke-width="2.5" stroke-dasharray="12 8" opacity=".75"/>`;
    const fsR = 20, twR = larguraTxt(ref.rotulo, fsR, 600) + 8, thR = fsR + 6;
    const amostras = [];  // pontos das séries e trechos entre eles, mais a caixa do rótulo de valor acima
    series.forEach(se => se.valores.forEach((v, i) => {
      if (v == null) return;
      amostras.push([X(i), Y(v)], [X(i), Y(v) - 40]);
      const prox = se.valores[i + 1];
      if (prox != null) for (let t = 1; t < 10; t++) amostras.push([X(i) + espX * t / 10, Y(v) + (Y(prox) - Y(v)) * t / 10]);
    }));
    const cands = [[px + 12, yr - 8 - thR], [px + 12, yr + 8], [px + pw - 12 - twR, yr - 8 - thR], [px + pw - 12 - twR, yr + 8]]
      .filter(([, cy]) => cy >= py + 2 && cy + thR <= py + ph - 2);
    const custo = ([cx, cy]) => amostras.filter(([ax, ay]) => ax > cx - 14 && ax < cx + twR + 14 && ay > cy - 14 && ay < cy + thR + 14).length;
    const [rx, ry] = cands.length ? cands.reduce((a, b) => custo(b) < custo(a) ? b : a) : [px + 12, yr - 8 - thR];
    ocupados.push({x: rx, y: ry, w: twR, h: thR});
    textoRef = `<text x="${rx + 4}" y="${ry + thR - 7}" font-size="${fsR}" fill="#B3261E" font-weight="600" stroke="#fff" stroke-width="5" stroke-linejoin="round" paint-order="stroke">${esc(ref.rotulo)}</text>`;
  }
  // séries
  const passoRot = n > 16 ? Math.ceil(n / 12) : 1;
  series.forEach((se, k) => {
    const pts = se.valores.map((v, i) => v == null ? null : [X(i), Y(v), v, i]).filter(Boolean);
    if (!pts.length) return;
    if (area && series.length === 1 && pts.length > 1) {
      s += `<defs><linearGradient id="${id}a${k}" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0" stop-color="${se.cor}" stop-opacity=".16"/>
              <stop offset="1" stop-color="${se.cor}" stop-opacity="0"/></linearGradient></defs>`;
      const d = pts.map((p, j) => (j ? "L" : "M") + p[0].toFixed(1) + " " + p[1].toFixed(1)).join(" ") +
                ` L${pts[pts.length-1][0].toFixed(1)} ${py + ph} L${pts[0][0].toFixed(1)} ${py + ph} Z`;
      s += `<path d="${d}" fill="url(#${id}a${k})"/>`;
    }
    // segmentos (quebra onde há semana sem dado)
    let seg = [];
    const flush = () => { if (seg.length > 1) s += `<polyline points="${seg.map(p => p[0].toFixed(1) + "," + p[1].toFixed(1)).join(" ")}" fill="none" stroke="${se.cor}" stroke-width="4.5" stroke-linejoin="round" stroke-linecap="round"/>`; seg = []; };
    se.valores.forEach((v, i) => { if (v == null) flush(); else seg.push([X(i), Y(v)]); });
    flush();
    for (const [cx, cy, v, i] of pts) {
      s += `<circle cx="${cx}" cy="${cy}" r="7" fill="${se.cor}" stroke="#fff" stroke-width="2.5"/>`;
      if (i % passoRot !== 0 && i !== n - 1) continue;
      const txt = fmt(v);
      // fonte cabe no espaço entre pontos (com folga para a caixa)
      const fs = Math.max(16, Math.min(n > 10 ? 19 : 22, Math.floor(22 * (espX * passoRot - 14) / (larguraTxt(txt, 22, 600) + 22))));
      const tw = larguraTxt(txt, fs, 600) + (fs >= 20 ? 22 : 14), th = fs + 14;
      const abaixo = !!se.rotuloAbaixo;
      const cima = cy - 18 - th, baixo = cy + 18;
      let ry = abaixo ? baixo : cima;
      // encostou em outro rótulo: tenta o lado oposto do ponto
      if (colide({x: cx - tw/2, y: ry, w: tw, h: th})) {
        const alt = ry === cima ? baixo : cima;
        if (!colide({x: cx - tw/2, y: alt, w: tw, h: th})) ry = alt;
      }
      ocupados.push({x: cx - tw/2, y: ry, w: tw, h: th});
      s += `<rect x="${cx - tw/2}" y="${ry}" width="${tw}" height="${th}" rx="7" fill="#E8EAEE"/>`;
      s += `<text x="${cx}" y="${ry + th/2 + fs*0.36}" text-anchor="middle" font-size="${fs}" fill="#1F2A3D" font-weight="600">${esc(txt)}</text>`;
    }
  });
  s += textoRef;
  if (legenda) {
    let lx = px;
    series.forEach(se => {
      s += `<line x1="${lx}" x2="${lx + 34}" y1="${y + 18}" y2="${y + 18}" stroke="${se.cor}" stroke-width="5" stroke-linecap="round"/>`;
      s += `<text x="${lx + 44}" y="${y + 26}" font-size="22" fill="#2D3748">${esc(se.nome)}</text>`;
      lx += 44 + larguraTxt(se.nome, 22) + 40;
    });
  }
  return s;
}

function iconeTendencia(x, y, sobe){
  const pts = sobe ? "8,30 18,20 25,25 36,10" : "8,12 18,22 25,17 36,30";
  return `<g transform="translate(${x},${y})"><rect x="0" y="0" width="44" height="40" rx="5" fill="#5B6676"/>
    <polyline points="${pts}" fill="none" stroke="#fff" stroke-width="3.5" stroke-linejoin="round" stroke-linecap="round"/></g>`;
}

function cabecalho(kicker, titulo, sub){
  let s = `<rect width="${W}" height="${H}" fill="#FFFFFF"/>`;
  s += `<text x="80" y="78" font-size="21" font-weight="700" letter-spacing="3.5" fill="#5B6676">${esc(kicker)}</text>`;
  s += `<text x="80" y="140" font-size="${caber(titulo, 50, W - 160, 700)}" font-weight="700" fill="#1F2A3D">${esc(titulo)}</text>`;
  if (sub) s += `<text x="80" y="192" font-size="28" fill="#4A5568">${esc(sub)}</text>`;
  return s;
}
const svgDoc = corpo => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" font-family="${FONTE}">${corpo}</svg>`;

// ---------------- recortes de período
function indicesPeriodo(datas, de, ate){
  const out = [];
  datas.forEach((d, i) => { if ((!de || d >= de) && (!ate || d <= ate)) out.push(i); });
  return out;
}
function unidadeInfo(nome){
  return BASE.unidades.find(u => u.nome.toUpperCase() === nome.toUpperCase()) || {nome, capacidade: null, cor: ""};
}
function variacaoTexto(v){
  if (v == null || !isFinite(v)) return null;
  if (Math.abs(v) < 0.0005) return {rot: "Estável", resto: " no período."};
  return {rot: (v < 0 ? "Queda de " : "Alta de ") + fPct(Math.abs(v)), resto: " no período."};
}

// ---------------- SLIDE 1: fluxo Salvador e RMS
function slideFluxo(){
  const e = EST.fluxo, ssa = BASE.ssa;
  const idx = indicesPeriodo(ssa.datas, e.de, e.ate);
  const nomes = (e.unidades || []).filter(n => ssa.series[n]);
  const datas = idx.map(i => ssa.datas[i]);
  const sub = datas.length ? `Acompanhamento de ${fDMA(datas[0])} a ${fDMA(datas[datas.length-1])}.` : "Sem semanas no período escolhido.";
  let s = cabecalho(e.kicker, e.titulo, sub);
  if (!nomes.length) { s += `<text x="${W/2}" y="560" text-anchor="middle" font-size="30" fill="#8A94A3">Marque ao menos uma unidade no painel à esquerda.</text>`; return svgDoc(s); }
  const n = nomes.length, gap = 70, mL = 80, larg = (W - 2 * mL - gap * (n - 1)) / n;
  nomes.forEach((nome, k) => {
    const u = unidadeInfo(nome), cap = u.capacidade;
    const cor = u.cor && /^#?[0-9a-f]{6}$/i.test(u.cor) ? (u.cor[0] === "#" ? u.cor : "#" + u.cor) : CORES_PADRAO[k % CORES_PADRAO.length];
    const brutos = idx.map(i => ssa.series[nome][i]);
    const validos = brutos.filter(v => v != null);
    const vari = validos.length > 1 ? validos[validos.length-1] / validos[0] - 1 : null;
    const pct = e.metrica === "pct" && cap;
    const valores = pct ? brutos.map(v => v == null ? null : v / cap) : brutos;
    const x0 = mL + k * (larg + gap);
    const sobe = vari == null ? true : vari >= 0;
    s += iconeTendencia(x0, 262, sobe);
    const nomeU = u.exibicao || u.nome;
    s += `<text x="${x0 + 58}" y="296" font-size="${caber(nomeU, n > 2 ? 30 : 34, larg - 60, 700)}" font-weight="700" fill="#4F5B6E">${esc(nomeU)}</text>`;
    const vt = variacaoTexto(vari);
    const pCap = `<tspan font-weight="700">Capacidade:</tspan> ${cap ? fInt(cap) + " vagas" : "não cadastrada"}`;
    const pVar = vt ? `<tspan font-weight="700">${esc(vt.rot)}</tspan>${vt.resto}` : "";
    const pOcu = validos.length && cap ? `Ocupação atual: <tspan font-weight="700">${fPct(validos[validos.length-1] / cap, 0)}</tspan>` : "";
    const fsInfo = n > 2 ? 22 : 24, umaLinha = [pCap, pVar, pOcu].filter(Boolean).join(" | ");
    if (larguraTxt(semTags(umaLinha), fsInfo, 600) <= larg)
      s += `<text x="${x0}" y="346" font-size="${fsInfo}" fill="#1F2A3D">${umaLinha}</text>`;
    else {
      // não cabe numa linha sem encolher demais: capacidade/ocupação em cima, variação embaixo
      const l1 = [pCap, pOcu].filter(Boolean).join(" | ");
      s += `<text x="${x0}" y="336" font-size="${caber(semTags(l1), fsInfo, larg, 600)}" fill="#1F2A3D">${l1}</text>`;
      if (pVar) s += `<text x="${x0}" y="366" font-size="${caber(semTags(pVar), fsInfo, larg, 600)}" fill="#1F2A3D">${pVar}</text>`;
    }
    const fmt = pct ? (v => fPct(v, 0)) : fInt;
    const fmtEixo = pct ? (v => fPct(v, Math.abs(v * 100 - Math.round(v * 100)) < 1e-6 ? 0 : 1)) : fInt;
    const ref = cap ? (pct ? {valor: 1, rotulo: "100% da capacidade"} : {valor: cap, rotulo: "Capacidade: " + fInt(cap)}) : null;
    // em "% da capacidade" a linha de 100% sempre aparece; em custodiados, só se não esmagar a série
    let refUsar = pct ? ref : null;
    if (!pct && ref && validos.length) {
      const vv = valores.filter(v => v != null), mn = Math.min(...vv), mx = Math.max(...vv), amp = (mx - mn) || mx;
      if (ref.valor >= mn - amp * 0.8 && ref.valor <= mx + amp * 0.8) refUsar = ref;
    }
    s += graficoLinhas({x: x0, y: 380, w: larg, h: 640, cats: datas.map(fDM), series: [{nome, cor, valores}],
                        fmt, fmtEixo, ref: refUsar, tituloY: pct ? "% da capacidade" : "Custodiados"});
  });
  const fontes = [...new Set(idx.map(i => ssa.fontes[i]).filter(Boolean))];
  s += `<text x="80" y="1058" font-size="17" fill="#8A94A3">${esc(fontes.length ? "Fonte: " + fontes.join("; ") + "." : "")}</text>`;
  return svgDoc(s);
}

// ---------------- SLIDE 2: Interior, semana
// FORA DO APP desde 18/11/2026 (pedido do usuário): sem aba, sem controles e fora do PPTX.
// Mantido só para reativar fácil (aba "interior" no <nav>, montarControles, svgAtual e exportarPptx).
function slideInterior(){
  const e = EST.interior, lista = BASE.interior;
  const i = Math.max(0, lista.findIndex(r => r.data === e.data));
  const r = lista[i], ant = i > 0 ? lista[i-1] : null;
  const ref = r ? (r.fonte || ("Semana de " + fDMA(r.data))) : "";
  let s = cabecalho(e.kicker, e.titulo, "Pessoas privadas de liberdade com audiência de custódia já realizada.");
  s += `<text x="80" y="232" font-size="22" fill="#4A5568">Referências: <tspan font-weight="700">${esc(ref)}</tspan>.</text>`;
  if (!r) { s += `<text x="${W/2}" y="560" text-anchor="middle" font-size="30" fill="#8A94A3">Sem lançamentos na aba INTERIOR_SEMANAL.</text>`; return svgDoc(s); }
  const aptos = r.del_total ? r.del_aud / r.del_total : null;
  const acima = r.up_capacidade ? r.up_internos / r.up_capacidade - 1 : null;
  const delta = (atual, anterior, pct) => {
    if (!ant || atual == null || anterior == null) return "";
    const d = atual - anterior;
    if (Math.abs(d) < (pct ? 0.0005 : 0.5)) return `= igual a ${fDM(ant.data)}`;
    const seta = d > 0 ? "▲" : "▼";
    const val = pct ? (Math.abs(d) * 100).toLocaleString("pt-BR", {maximumFractionDigits: 1}) + " p.p." : fInt(Math.abs(d));
    return `${seta} ${val} desde ${fDM(ant.data)}`;
  };
  const cartao = (x, cor, titulo, w, h) =>
    `<rect x="${x+6}" y="${306}" width="${w}" height="${h}" rx="18" fill="#1F2A3D" opacity=".12"/>
     <rect x="${x}" y="300" width="${w}" height="${h}" rx="18" fill="${cor}"/>
     <text x="${x + w/2}" y="356" text-anchor="middle" font-size="28" font-weight="700" fill="#fff" letter-spacing="1">${titulo}</text>
     <rect x="${x+22}" y="388" width="${w-44}" height="${h-110}" rx="10" fill="#fff"/>`;
  const numero = (x, y, v, rot, dl) =>
    `<text x="${x}" y="${y}" font-size="66" font-weight="700" fill="#1F2A3D">${esc(v)}</text>
     <text x="${x}" y="${y+36}" font-size="22" fill="#4A5568">${esc(rot)}</text>` +
    (dl ? `<text x="${x}" y="${y+66}" font-size="19" fill="#6B7585">${esc(dl)}</text>` : "");
  const destaque = (x, y, v, rot1, rot2, dl) =>
    `<rect x="${x+4}" y="${y+5}" width="250" height="200" rx="10" fill="#1F2A3D" opacity=".10"/>
     <rect x="${x}" y="${y}" width="250" height="200" rx="10" fill="#fff" stroke="#E1E5EB" stroke-width="2"/>
     <text x="${x+125}" y="${y+88}" text-anchor="middle" font-size="64" font-weight="700" fill="#B3261E">${esc(v)}</text>
     <text x="${x+125}" y="${y+128}" text-anchor="middle" font-size="21" fill="#4A5568">${esc(rot1)}</text>
     <text x="${x+125}" y="${y+154}" text-anchor="middle" font-size="21" fill="#4A5568">${esc(rot2)}</text>` +
    (dl ? `<text x="${x+125}" y="${y+186}" text-anchor="middle" font-size="17" fill="#6B7585">${esc(dl)}</text>` : "");
  const cw = 640, ch = 620, xa = 230, xb = W - 230 - cw;
  s += cartao(xa, "#2B4C8C", "DELEGACIAS INTERIOR", cw, ch);
  s += numero(xa + 56, 490, fInt(r.del_total), "Total de custodiados", delta(r.del_total, ant && ant.del_total));
  s += numero(xa + 56, 720, fInt(r.del_aud), "Custodiados com audiência realizada", delta(r.del_aud, ant && ant.del_aud));
  s += destaque(xa + cw - 300, 420, fPct(aptos, 0), "Aptos para", "transferência", ant && ant.del_total ? delta(aptos, ant.del_aud / ant.del_total, true) : "");
  // seta
  s += `<path d="M ${W/2-50} 560 L ${W/2+20} 560 L ${W/2+20} 530 L ${W/2+70} 610 L ${W/2+20} 690 L ${W/2+20} 660 L ${W/2-50} 660 Z" fill="#C3CCD8"/>`;
  s += cartao(xb, "#3A4556", "UNIDADES PRISIONAIS", cw, ch);
  s += numero(xb + 56, 480, fInt(r.up_internos), "Total de internos", delta(r.up_internos, ant && ant.up_internos));
  s += numero(xb + 56, 650, fInt(r.up_capacidade), "Capacidade de lotação", "");
  s += numero(xb + 56, 800, fInt(r.up_conjuntos), "Total de conjuntos prisionais", "");
  s += destaque(xb + cw - 300, 420, fPct(acima, 0), "Acima da", "capacidade", ant && ant.up_capacidade ? delta(acima, ant.up_internos / ant.up_capacidade - 1, true) : "");
  // observação com quebra simples
  const obs = (r.obs ? r.obs + " " : "") + (e.obs || "");
  const linhas = quebrar(obs, 150);
  linhas.forEach((l, j) => s += `<text x="80" y="${990 + j*30}" font-size="21" fill="#4A5568">${esc(l)}</text>`);
  return svgDoc(s);
}
function quebrar(txt, max){
  const out = []; let at = "";
  for (const p of String(txt).split(/\s+/)) {
    if ((at + " " + p).trim().length > max) { out.push(at.trim()); at = p; } else at += " " + p;
  }
  if (at.trim()) out.push(at.trim());
  return out.slice(0, 3);
}

// ---------------- SLIDE 3: Interior, evolução
function slideEvolucao(){
  const e = EST.evolucao, lista = BASE.interior;
  const datasT = lista.map(r => r.data);
  const idx = indicesPeriodo(datasT, e.de, e.ate);
  const rs = idx.map(i => lista[i]);
  const sub = rs.length ? `Acompanhamento de ${fDMA(rs[0].data)} a ${fDMA(rs[rs.length-1].data)}.` : "Sem semanas no período escolhido.";
  let s = cabecalho(e.kicker, e.titulo, sub);
  if (!rs.length) return svgDoc(s);
  const cats = rs.map(r => fDM(r.data));
  const larg = (W - 160 - 70) / 2, x1 = 80, x2 = 80 + larg + 70;
  const pri = rs[0], ult = rs[rs.length-1];
  // esquerda: delegacias
  s += iconeTendencia(x1, 262, (ult.del_total ?? 0) >= (pri.del_total ?? 0));
  s += `<text x="${x1 + 58}" y="296" font-size="34" font-weight="700" fill="#4F5B6E">Delegacias do Interior</text>`;
  const vD = pri.del_total ? ult.del_total / pri.del_total - 1 : null, vtD = variacaoTexto(vD);
  const infoD = `<tspan font-weight="700">Custodiados:</tspan> ${fInt(ult.del_total)} na última semana${vtD ? ` | <tspan font-weight="700">${esc(vtD.rot)}</tspan>${vtD.resto}` : ""}`;
  s += `<text x="${x1}" y="346" font-size="${caber(semTags(infoD), 24, larg, 600)}" fill="#1F2A3D">${infoD}</text>`;
  // séries escolhidas no painel lateral (caixas "Linhas do gráfico")
  const seriesD = [];
  if (e.verTotal) seriesD.push({nome: "Total de custodiados", cor: "#2B4C8C", valores: rs.map(r => r.del_total)});
  if (e.verAud) seriesD.push({nome: "Com audiência realizada", cor: "#5FA8D3", valores: rs.map(r => r.del_aud), rotuloAbaixo: e.verTotal});
  s += graficoLinhas({x: x1, y: 380, w: larg, h: 640, cats, legenda: true, area: false, fmt: fInt, tituloY: "Custodiados",
    series: seriesD});
  // direita: unidades prisionais
  const acima = ult.up_capacidade ? ult.up_internos / ult.up_capacidade - 1 : null;
  s += iconeTendencia(x2, 262, (ult.up_internos ?? 0) >= (pri.up_internos ?? 0));
  s += `<text x="${x2 + 58}" y="296" font-size="34" font-weight="700" fill="#4F5B6E">Unidades Prisionais do Interior</text>`;
  const vU = pri.up_internos ? ult.up_internos / pri.up_internos - 1 : null, vtU = variacaoTexto(vU);
  const infoU = `<tspan font-weight="700">Capacidade:</tspan> ${fInt(ult.up_capacidade)} vagas | <tspan font-weight="700">${fPct(acima, 0)}</tspan> acima da capacidade${vtU ? ` | <tspan font-weight="700">${esc(vtU.rot)}</tspan>${vtU.resto}` : ""}`;
  s += `<text x="${x2}" y="346" font-size="${caber(semTags(infoU), 24, larg, 600)}" fill="#1F2A3D">${infoU}</text>`;
  s += graficoLinhas({x: x2, y: 380, w: larg, h: 640, cats, fmt: fInt, tituloY: "Internos",
    series: [{nome: "Total de internos", cor: "#B8902A", valores: rs.map(r => r.up_internos)}]});
  const fontes = [...new Set(rs.map(r => r.fonte).filter(Boolean))];
  s += `<text x="80" y="1058" font-size="17" fill="#8A94A3">Fonte: Relatório site SEAP e Card PCBA (${rs.length} semana${rs.length > 1 ? "s" : ""}).</text>`;
  return svgDoc(s);
}

// ---------------- controles
function opcoesDatas(datas, sel){
  return datas.map(d => `<option value="${d}" ${d === sel ? "selected" : ""}>${fDMA(d)}</option>`).join("");
}
function campoTexto(aba, campo, rotulo){
  return `<h3>${rotulo}</h3><input type="text" value="${esc(EST[aba][campo])}" oninput="EST['${aba}']['${campo}']=this.value;guardarTextos();desenhar()">`;
}
function montarControles(){
  const c = document.getElementById("controles");
  if (!BASE) { c.innerHTML = ""; return; }
  if (ABA === "fluxo") {
    const e = EST.fluxo, ds = BASE.ssa.datas, nomes = Object.keys(BASE.ssa.series);
    c.innerHTML = `
      <h3>Período</h3>
      <div class="linha2"><div><span>De</span><select onchange="EST.fluxo.de=this.value;desenhar()">${opcoesDatas(ds, e.de)}</select></div>
      <div><span>Até</span><select onchange="EST.fluxo.ate=this.value;desenhar()">${opcoesDatas(ds, e.ate)}</select></div></div>
      <h3>Exibir valores como</h3>
      <div class="seg"><button class="${e.metrica === "abs" ? "ativo" : ""}" onclick="EST.fluxo.metrica='abs';montarControles();desenhar()">Custodiados</button>
      <button class="${e.metrica === "pct" ? "ativo" : ""}" onclick="EST.fluxo.metrica='pct';montarControles();desenhar()">% da capacidade</button></div>
      <h3>Unidades no slide</h3>
      ${nomes.map(n => `<label><input type="checkbox" ${e.unidades.includes(n) ? "checked" : ""} onchange="alternarUnidade(${esc(JSON.stringify(n))},this.checked)">${esc(unidadeInfo(n).exibicao || n)}</label>`).join("")}
      <p class="dica">O padrão vem da coluna EXIBIR NO PAINEL da aba UNIDADES. Até 3 unidades cabem bem lado a lado.</p>
      ${campoTexto("fluxo", "kicker", "Sobretítulo")}
      ${campoTexto("fluxo", "titulo", "Título")}`;
  } else if (ABA === "evolucao") {
    const e = EST.evolucao, ds = BASE.interior.map(r => r.data);
    c.innerHTML = `
      <h3>Período</h3>
      <div class="linha2"><div><span>De</span><select onchange="EST.evolucao.de=this.value;desenhar()">${opcoesDatas(ds, e.de)}</select></div>
      <div><span>Até</span><select onchange="EST.evolucao.ate=this.value;desenhar()">${opcoesDatas(ds, e.ate)}</select></div></div>
      <h3>Linhas do gráfico (Delegacias)</h3>
      <label><input type="checkbox" ${e.verTotal ? "checked" : ""} onchange="alternarSerieEvolucao('verTotal',this)">Total de custodiados</label>
      <label><input type="checkbox" ${e.verAud ? "checked" : ""} onchange="alternarSerieEvolucao('verAud',this)">Com audiência realizada</label>
      <p class="dica">Pelo menos uma linha fica sempre visível.</p>
      ${campoTexto("evolucao", "kicker", "Sobretítulo")}
      ${campoTexto("evolucao", "titulo", "Título")}`;
  }
}
function alternarSerieEvolucao(chave, caixa){
  const e = EST.evolucao, outra = chave === "verTotal" ? "verAud" : "verTotal";
  // não deixa desmarcar as duas (o gráfico ficaria vazio)
  if (!caixa.checked && !e[outra]) { caixa.checked = true; return; }
  e[chave] = caixa.checked;
  desenhar();
}
function alternarUnidade(n, on){
  const u = EST.fluxo.unidades;
  if (on && !u.includes(n)) u.push(n);
  if (!on) EST.fluxo.unidades = u.filter(x => x !== n);
  const ordem = Object.keys(BASE.ssa.series);
  EST.fluxo.unidades.sort((a, b) => ordem.indexOf(a) - ordem.indexOf(b));
  desenhar();
}

function svgAtual(){
  return ABA === "fluxo" ? slideFluxo() : slideEvolucao();
}
function desenhar(){
  const main = document.getElementById("main"), dados = document.getElementById("dados");
  if (ABA === "dados") { main.classList.add("oculto"); dados.classList.remove("oculto"); montarDados(); return; }
  main.classList.remove("oculto"); dados.classList.add("oculto");
  if (!BASE) return;
  document.getElementById("slide").innerHTML = svgAtual();
  document.getElementById("infoSlide").textContent = {
    fluxo: "Slide: fluxo semanal Salvador e RMS",
    evolucao: "Slide: evolução semanal do Interior"}[ABA] + " · 1920 × 1080";
}

function montarDados(){
  const d = document.getElementById("dados");
  if (!BASE) { d.innerHTML = `<div class="vazio">Base não carregada.</div>`; return; }
  // LANÇAMENTOS (27/11/2026): as tabelas da base, agora editáveis aqui
  // mesmo (incluir / editar / excluir), mais recentes em cima.
  const nomes = Object.keys(BASE.ssa.series);
  const acoes = (tipo, chave) => `<td class="acoes-linha"><button type="button" class="btn mini" onclick="abrirForm('${tipo}', ${esc(JSON.stringify(chave))})">Editar</button> <button type="button" class="btn mini perigo" onclick="excluirReg('${tipo}', ${esc(JSON.stringify(chave))})">Excluir</button></td>`;
  let h = "";
  if (BASE.avisos.length) h += `<div class="aviso"><b>Pontos de atenção</b><ul>${BASE.avisos.map(a => `<li>${esc(a)}</li>`).join("")}</ul></div>`;
  h += `<p class="dica">Os lançamentos ficam gravados no próprio sistema (${esc(BASE.arquivo)}) · última alteração em ${esc(BASE.modificado_em)}${BASE.alterado_por ? " por " + esc(BASE.alterado_por) : ""}. Os gráficos se atualizam assim que você salva.</p>`;
  const idxS = BASE.ssa.datas.map((_, i) => i).reverse();
  h += `<div class="sec-cab"><h3>Salvador e RMS · custodiados por semana</h3><button type="button" class="btn prim" onclick="abrirForm('ssa', null)">+ Nova semana</button></div>`;
  h += `<table><tr><th>Semana</th>${nomes.map(n => `<th>${esc(unidadeInfo(n).exibicao || n)}</th>`).join("")}<th>Fonte</th><th>Observação</th><th></th></tr>` +
       (idxS.length ? idxS.map(i => `<tr><td>${fDMA(BASE.ssa.datas[i])}</td>${nomes.map(n => `<td class="n">${fInt(BASE.ssa.series[n][i])}</td>`).join("")}<td>${esc(BASE.ssa.fontes[i])}</td><td>${esc(BASE.ssa.obs[i])}</td>${acoes("ssa", BASE.ssa.datas[i])}</tr>`).join("")
        : `<tr><td colspan="${nomes.length + 4}" class="dica">Nenhuma semana lançada.</td></tr>`) + `</table>`;
  const intRev = BASE.interior.slice().reverse();
  h += `<div class="sec-cab"><h3>Interior · quantitativos por semana</h3><button type="button" class="btn prim" onclick="abrirForm('interior', null)">+ Nova semana</button></div>`;
  h += `<table><tr><th>Semana</th><th>Deleg. total</th><th>Com audiência</th><th>% aptos</th><th>Internos</th><th>Capacidade</th><th>% acima</th><th>Conjuntos</th><th>Referência</th><th></th></tr>` +
       (intRev.length ? intRev.map(r => `<tr><td>${fDMA(r.data)}</td><td class="n">${fInt(r.del_total)}</td><td class="n">${fInt(r.del_aud)}</td><td class="n">${fPct(r.del_total ? r.del_aud / r.del_total : null)}</td><td class="n">${fInt(r.up_internos)}</td><td class="n">${fInt(r.up_capacidade)}${r.bruto_capacidade == null && r.up_capacidade != null ? " *" : ""}</td><td class="n">${fPct(r.up_capacidade ? r.up_internos / r.up_capacidade - 1 : null)}</td><td class="n">${fInt(r.up_conjuntos)}${r.bruto_conjuntos == null && r.up_conjuntos != null ? " *" : ""}</td><td>${esc(r.fonte)}</td>${acoes("interior", r.data)}</tr>`).join("")
        : `<tr><td colspan="10" class="dica">Nenhuma semana lançada.</td></tr>`) + `</table>`;
  h += `<p class="dica" style="margin-top:-14px">* capacidade / conjuntos em branco na semana: repete o valor da semana anterior.</p>`;
  h += `<div class="sec-cab"><h3>Unidades · Salvador e RMS</h3><button type="button" class="btn prim" onclick="abrirForm('unidade', null)">+ Nova unidade</button></div>`;
  h += `<table><tr><th>Unidade</th><th>Nome no gráfico</th><th>Capacidade</th><th>No painel</th><th>Cor</th><th>Observação</th><th></th></tr>` +
       BASE.unidades.map(u => `<tr><td>${esc(u.nome)}</td><td>${esc(u.exibicao)}</td><td class="n">${fInt(u.capacidade)}</td><td>${u.exibir ? "Sim" : "Não"}</td><td>${u.cor ? `<span class="cor-amostra" style="background:${esc(u.cor)}"></span> ${esc(u.cor)}` : "padrão"}</td><td>${esc(u.obs || "")}</td>${acoes("unidade", u.nome)}</tr>`).join("") + `</table>`;
  d.innerHTML = h;
}

// ---------------- formulário dos lançamentos (27/11/2026)
let FORM = null;
const numIn = v => v == null ? "" : nf0.format(v);
function campoForm(nome, rotulo, valor, extra = {}){
  const cls = extra.largo ? "largo" : "";
  if (extra.tipo === "check")
    return `<label class="check ${cls}"><input type="checkbox" name="${nome}" ${valor ? "checked" : ""}> ${esc(rotulo)}</label>`;
  if (extra.tipo === "area")
    return `<label class="${cls}">${esc(rotulo)}<textarea name="${nome}" rows="2" maxlength="500">${esc(valor ?? "")}</textarea></label>`;
  const tipo = extra.tipo || "text";
  const num = extra.num ? ` class="num" inputmode="numeric" placeholder="${extra.ph ?? ""}"` : (extra.ph ? ` placeholder="${esc(extra.ph)}"` : "");
  return `<label class="${cls}">${esc(rotulo)}<input type="${tipo}" name="${nome}" value="${esc(valor ?? "")}"${num}${extra.ro ? " readonly" : ""}${extra.max ? ` maxlength="${extra.max}"` : ""}></label>`;
}
function proximaSemana(datas){
  if (!datas.length) return new Date().toISOString().slice(0, 10);
  const d = new Date(datas[datas.length - 1] + "T12:00:00"); d.setDate(d.getDate() + 7);
  return d.toISOString().slice(0, 10);
}
function abrirForm(tipo, chave){
  if (!BASE) return;
  FORM = {tipo, chave};
  let tit = "", sub = "", campos = "";
  if (tipo === "ssa") {
    const i = chave ? BASE.ssa.datas.indexOf(chave) : -1;
    tit = chave ? "Editar semana · Salvador e RMS" : "Nova semana · Salvador e RMS";
    sub = "Total de custodiados de cada unidade na data de referência. Deixe em branco a unidade sem informação na semana.";
    campos += campoForm("data", "Data de referência", i >= 0 ? BASE.ssa.datas[i] : proximaSemana(BASE.ssa.datas), {tipo: "date"});
    campos += `<span></span><div class="grupo-tit">CUSTODIADOS POR UNIDADE</div>`;
    Object.keys(BASE.ssa.series).forEach(n => {
      const u = unidadeInfo(n);
      campos += campoForm("u:" + n, (u.exibicao || n) + (u.capacidade ? ` (capacidade ${fInt(u.capacidade)})` : ""),
        i >= 0 ? numIn(BASE.ssa.series[n][i]) : "", {num: true});
    });
    campos += campoForm("fonte", "Fonte", i >= 0 ? BASE.ssa.fontes[i] : "", {largo: true, max: 500});
    campos += campoForm("obs", "Observação", i >= 0 ? BASE.ssa.obs[i] : "", {largo: true, tipo: "area"});
  } else if (tipo === "interior") {
    const r = chave ? BASE.interior.find(x => x.data === chave) : null;
    tit = chave ? "Editar semana · Interior" : "Nova semana · Interior";
    sub = "Capacidade e nº de conjuntos em branco repetem o valor da semana anterior.";
    campos += campoForm("data", "Data de referência", r ? r.data : proximaSemana(BASE.interior.map(x => x.data)), {tipo: "date"});
    campos += `<span></span><div class="grupo-tit">DELEGACIAS DO INTERIOR</div>`;
    campos += campoForm("del_total", "Total de custodiados", r ? numIn(r.del_total) : "", {num: true});
    campos += campoForm("del_aud", "Com audiência realizada", r ? numIn(r.del_aud) : "", {num: true});
    campos += `<div class="grupo-tit">UNIDADES PRISIONAIS DO INTERIOR</div>`;
    campos += campoForm("up_internos", "Total de internos", r ? numIn(r.up_internos) : "", {num: true});
    campos += campoForm("up_capacidade", "Capacidade de lotação", r ? numIn(r.bruto_capacidade) : "", {num: true, ph: "repete a anterior"});
    campos += campoForm("up_conjuntos", "Total de conjuntos prisionais", r ? numIn(r.bruto_conjuntos) : "", {num: true, ph: "repete a anterior"});
    campos += `<span></span>`;
    campos += campoForm("fonte", "Referência (fonte)", r ? r.fonte : "", {largo: true, max: 500});
    campos += campoForm("obs", "Observação", r ? r.obs : "", {largo: true, tipo: "area"});
  } else {
    const u = chave ? BASE.unidades.find(x => x.nome === chave) : null;
    tit = chave ? "Editar unidade" : "Nova unidade";
    sub = chave ? "Mudar o nome da unidade leva junto os números já lançados." : "A unidade nova ganha uma coluna nas semanas de Salvador e RMS.";
    campos += campoForm("nome", "Unidade", u ? u.nome : "", {largo: true, max: 200});
    campos += campoForm("exibicao", "Nome no gráfico", u ? u.exibicao : "", {max: 200});
    campos += campoForm("regiao", "Região", u ? u.regiao : "SALVADOR E RMS", {max: 100});
    campos += campoForm("capacidade", "Capacidade de lotação (vagas)", u ? numIn(u.capacidade) : "", {num: true});
    campos += `<label>Cor da linha<input type="color" name="cor" value="${esc(u && u.cor ? u.cor : "#2B4C8C")}"></label>`;
    campos += campoForm("exibir", "Exibir no painel (marcada por padrão no slide)", u ? u.exibir : true, {tipo: "check", largo: true});
    campos += campoForm("obs", "Observação", u ? (u.obs || "") : "", {largo: true, tipo: "area"});
  }
  document.getElementById("formTit").textContent = tit;
  document.getElementById("formSub").textContent = sub;
  document.getElementById("formCampos").innerHTML = campos;
  document.getElementById("formErro").textContent = "";
  document.getElementById("modal").classList.remove("oculto");
  const primeiro = document.querySelector("#formCampos input:not([readonly])");
  if (primeiro) primeiro.focus();
}
function fecharForm(){ document.getElementById("modal").classList.add("oculto"); FORM = null; }
document.addEventListener("keydown", e => { if (e.key === "Escape" && FORM) fecharForm(); });
// números: só dígitos, com ponto de milhar enquanto digita (5657 -> 5.657)
document.addEventListener("input", e => {
  const c = e.target;
  if (!c.matches || !c.matches("#formCampos input.num")) return;
  const dig = c.value.replace(/[^0-9]/g, "").replace(/^0+(?=\d)/, "").slice(0, 9);
  c.value = dig ? nf0.format(+dig) : "";
});
async function depoisDeGravar(r, msg){
  if (r && r.base) {
    aplicarBase(r.base);
    try { _mtimeBase = await window.pywebview.api.versao_planilha(); _mtimePend = null; } catch(e) {}
  }
  if (r && r.ok) toast(msg);
}
async function salvarForm(ev){
  ev.preventDefault();
  if (!FORM) return false;
  const f = document.getElementById("formReg"), erro = document.getElementById("formErro");
  const val = n => { const c = f.elements[n]; return c ? c.value : ""; };
  let r, msg;
  const btn = document.getElementById("formSalvar");
  btn.disabled = true; erro.textContent = "";
  try {
    if (FORM.tipo === "ssa") {
      const valores = {};
      Object.keys(BASE.ssa.series).forEach(n => valores[n] = val("u:" + n));
      r = await window.pywebview.api.salvar_ssa({data: val("data"), valores, fonte: val("fonte"), obs: val("obs")}, FORM.chave, BASE.revisao);
      msg = "Semana de Salvador e RMS salva.";
    } else if (FORM.tipo === "interior") {
      const reg = {data: val("data"), fonte: val("fonte"), obs: val("obs")};
      ["del_total", "del_aud", "up_internos", "up_capacidade", "up_conjuntos"].forEach(k => reg[k] = val(k));
      r = await window.pywebview.api.salvar_interior(reg, FORM.chave, BASE.revisao);
      msg = "Semana do Interior salva.";
    } else {
      r = await window.pywebview.api.salvar_unidade({nome: val("nome"), exibicao: val("exibicao"), regiao: val("regiao"),
        capacidade: val("capacidade"), cor: val("cor"), exibir: f.elements["exibir"].checked, obs: val("obs")}, FORM.chave, BASE.revisao);
      msg = "Unidade salva.";
    }
  } catch(e) { r = {ok: false, erro: String(e)}; }
  btn.disabled = false;
  if (r && r.ok) { fecharForm(); await depoisDeGravar(r, msg); }
  else {
    if (r && r.conflito) await depoisDeGravar(r, "");
    erro.textContent = (r && r.erro) || "Não foi possível salvar.";
  }
  return false;
}
async function excluirReg(tipo, chave){
  if (!BASE) return;
  const desc = tipo === "unidade" ? `a unidade "${chave}" e todos os números dela já lançados` :
    `a semana de ${fDMA(chave)} (${tipo === "ssa" ? "Salvador e RMS" : "Interior"})`;
  if (!confirm(`Excluir ${desc}?\n\nFica registrado no histórico e há cópia de segurança.`)) return;
  const r = tipo === "unidade" ? await window.pywebview.api.excluir_unidade(chave, BASE.revisao)
                               : await window.pywebview.api.excluir_semana(tipo, chave, BASE.revisao);
  if (r && r.ok) await depoisDeGravar(r, "Excluído.");
  else { if (r && r.conflito) await depoisDeGravar(r, ""); toast((r && r.erro) || "Não foi possível excluir.", true); }
}
async function exportarExcel(){
  const r = await chamar("exportar_excel");
  if (r && r.ok) toast("Excel salvo em " + r.arquivo);
}

// ---------------- base
function aplicarBase(b){
  if (!b || !b.ok) {
    BASE = null;
    document.getElementById("infoBase").textContent = "Base não carregada";
    document.getElementById("slide").innerHTML = `<div class="vazio" style="white-space:pre-line">${esc(b ? b.erro : "Sem dados.")}</div>`;
    return;
  }
  BASE = b;
  const ds = b.ssa.datas, di = b.interior.map(r => r.data);
  const f = EST.fluxo;
  if (!f.de || !ds.includes(f.de)) f.de = ds[0];
  // "até" sempre acompanha a última semana ao recarregar
  f.ate = ds[ds.length - 1];
  if (!f.unidades) f.unidades = b.unidades.filter(u => u.exibir && b.ssa.series[u.nome]).map(u => u.nome);
  EST.interior.data = di[di.length - 1];
  if (!EST.evolucao.de || !di.includes(EST.evolucao.de)) EST.evolucao.de = di[0];
  EST.evolucao.ate = di[di.length - 1];
  const ult = [ds[ds.length-1], di[di.length-1]].filter(Boolean).sort().pop();
  document.getElementById("infoBase").textContent =
    `Base: ${b.arquivo.split(/[\\/]/).pop()} · última semana lançada ${ult ? fDMA(ult) : "-"}` + (b.avisos.length ? ` · ${b.avisos.length} aviso(s) na aba Dados` : "");
  montarControles(); desenhar();
}
async function recarregar(){
  if (!DESKTOP()) { toast("No modo navegador, feche e abra o aplicativo para reler a planilha."); return; }
  const r = await window.pywebview.api.carregar();
  aplicarBase(r);
  if (r && r.ok) toast("Base recarregada.");
}

// recarga automática: quando a planilha é salva, a data de modificação muda.
// Espera ela ficar igual em duas consultas seguidas (Excel terminou de gravar) antes de ler.
let _mtimeBase = null, _mtimePend = null, _vigiando = false;
async function vigiarPlanilha(){
  if (_vigiando) return;
  _vigiando = true;
  try {
    const m = await window.pywebview.api.versao_planilha();
    if (m == null) return;
    if (_mtimeBase == null) { _mtimeBase = m; return; }
    if (m === _mtimeBase) { _mtimePend = null; return; }
    if (m !== _mtimePend) { _mtimePend = m; return; }
    const r = await window.pywebview.api.carregar();
    // se falhar (arquivo ainda em gravação), mantém o que está na tela e tenta no próximo ciclo
    if (r && r.ok) {
      _mtimeBase = m; _mtimePend = null;
      aplicarBase(r);
      toast("Lançamentos atualizados (gravados em outra máquina).");
    }
  } catch(e) {
  } finally {
    _vigiando = false;
  }
}

// ---------------- exportação
function svgParaPng(svg, escala = 1.5){
  return new Promise((ok, falha) => {
    const img = new Image();
    img.onload = () => {
      const cv = document.createElement("canvas");
      cv.width = W * escala; cv.height = H * escala;
      const cx = cv.getContext("2d");
      cx.fillStyle = "#fff"; cx.fillRect(0, 0, cv.width, cv.height);
      cx.drawImage(img, 0, 0, cv.width, cv.height);
      ok(cv.toDataURL("image/png"));
    };
    img.onerror = falha;
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  });
}
const hoje = () => { const d = new Date(); return [d.getDate(), d.getMonth()+1, d.getFullYear()].map(n => String(n).padStart(2, "0")).join("-"); };
const nomeSlide = () => ({fluxo: "Fluxo SSA-RMS", interior: "Interior semana", evolucao: "Interior evolucao"})[ABA];

async function salvarSlide(){
  if (!BASE) return;
  const png = await svgParaPng(svgAtual());
  const nome = `Sistema Prisional - ${nomeSlide()} - ${hoje()}.png`;
  if (DESKTOP()) {
    const r = await chamar("salvar_png", nome, png);
    if (r && r.ok) toast("Imagem salva em " + r.arquivo);
  } else {
    const a = document.createElement("a"); a.href = png; a.download = nome; a.click();
  }
}
async function copiarSlide(){
  if (!BASE) return;
  const r = await chamar("copiar_png", await svgParaPng(svgAtual()));
  if (r && r.ok) toast("Imagem copiada. Cole no PowerPoint com Ctrl+V.");
}
async function exportarPptx(){
  if (!BASE) return;
  toast("Gerando apresentação...");
  const salva = ABA, imagens = [];
  for (const [aba, nota] of [["fluxo", "Fluxo semanal Salvador e RMS"], ["evolucao", "Interior: evolução semanal"]]) {
    ABA = aba;
    imagens.push({png: await svgParaPng(svgAtual(), 2), nota: nota + ". Gerado pelo app Sistema Prisional em " + hoje().replace(/-/g, "/") + "."});
  }
  ABA = salva;
  const r = await chamar("gerar_pptx", `Sistema Prisional - Graficos semanais - ${hoje()}.pptx`, imagens);
  if (r && r.ok) toast("Apresentação salva em " + r.arquivo);
}

// ---------------- abas e início
document.querySelectorAll("nav button").forEach(b => b.onclick = () => {
  document.querySelectorAll("nav button").forEach(x => x.classList.toggle("ativo", x === b));
  ABA = b.dataset.aba; montarControles(); desenhar();
});
function iniciar(){
  if (!DESKTOP()) document.querySelectorAll(".so-desktop").forEach(e => e.classList.add("oculto"));
  if (window.DADOS_INICIAIS) aplicarBase(window.DADOS_INICIAIS);
  else if (DESKTOP()) {
    vigiarPlanilha();  // registra a versão atual antes da primeira leitura
    recarregar();
    setInterval(vigiarPlanilha, 3000);
  }
}
if (window.DADOS_INICIAIS) iniciar();
else if (window.pywebview && window.pywebview.api) iniciar();
else window.addEventListener("pywebviewready", iniciar, {once: true});
</script>
</body>
</html>
"""


def html_com_dados(dados):
    js = json.dumps(dados, ensure_ascii=False).replace("</", "<\\/")
    return HTML.replace("/*__DADOS__*/null", js)


_arquivo_html = None


def abrir_janela(permitido, usuario=None):
    """Abre (ou traz para a frente) a janela do Sistema Prisional.
    Chamada pela API do CGAPE depois de conferir a permissão."""
    global janela, _arquivo_html
    if not permitido():
        return {"ok": False, "mensagem": "Acesso restrito: o Sistema Prisional é só para Silvana e Diego."}
    if webview is None:
        return {"ok": False, "mensagem": "Janela indisponível (pywebview não carregou)."}
    if janela is not None:
        try:
            janela.restore()
            janela.show()
            return {"ok": True, "mensagem": "Sistema Prisional já está aberto."}
        except Exception:
            janela = None
    if not os.path.isfile(CAMINHO_DADOS) and not os.path.isfile(CAMINHO_PLANILHA):
        return {"ok": False, "mensagem": f"Cadastro do Sistema Prisional não encontrado: {CAMINHO_DADOS}"}
    _arquivo_html = os.path.join(tempfile.gettempdir(), f"cgape_sistema_prisional_{os.getpid()}.html")
    with open(_arquivo_html, "w", encoding="utf-8") as f:
        f.write(HTML)
    janela = webview.create_window(
        NOME_APP, url=_arquivo_html, js_api=API(permitido, usuario),
        width=1500, height=920, min_size=(1100, 700), maximized=True, background_color="#EEF1F5",
    )

    def ao_fechar():
        global janela
        janela = None
        try:
            os.remove(_arquivo_html)
        except Exception:
            pass
    try:
        janela.events.closed += ao_fechar
    except Exception:
        pass
    return {"ok": True, "mensagem": "Sistema Prisional aberto."}
