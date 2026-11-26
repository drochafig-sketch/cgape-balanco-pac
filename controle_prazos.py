"""CONTROLE DE PRAZOS — terceira seção do CGAPE (irmã de PAC e MONITORAMENTO).

Incorporado em 25/11/2026 a pedido do usuário a partir do projeto avulso
"X:\\COAPI 2\\2026\\19. APRESENTAÇÃO GABINETE - PAUTAS REUNIÕES\\Controle de
prazos" (app.py / dados.py / mapa_html.py), com o layout e o design system
do PAC e do MONITORAMENTO e a cor VERMELHA (identidade_visual.VERMELHO) nos
destaques e contornos da seção.

Dados: desde 25/11/2026 as ações são cadastradas no PRÓPRIO SISTEMA (ver
"CADASTRO NO PRÓPRIO SISTEMA" abaixo): arquivo "Controle de Prazos -
dados.json" na pasta de rede ...\\PAUTAS REUNIÕES, editado pela tela (só a
Silvana; os demais visualizam). A planilha antiga
    ...\\PAUTAS REUNIÕES\\Controle de Prazos - Pauta Governador (Portos e Transporte).xlsx
só serve para a importação inicial. Cópia local de reserva em
%LOCALAPPDATA%\\CGAPE, usada (só leitura) quando a rede não está acessível.
O alerta (Vencido / Crítico / Atenção / OK / Sem data / Concluído) é
recalculado no navegador a partir de Data Prazo + Status, com a regra da aba
"Resumo e Legenda" da planilha antiga.

Peças:
  - api_dados() / api_assinatura(): dados da seção e "assinatura" (revisão
    do cadastro), para o dash se atualizar sozinho quando outra máquina grava;
    api_salvar() / api_excluir() / planilha_exportacao_bytes(): cadastro e
    Excel de consulta — registradas em "CGAPE - BALANÇO PAC.py" (_api_prazos_*,
    classe da API do pywebview) e em servidor_web.py;
  - HTML_JANELA: dash (#prz-overlay) + grade de filtros (#grade-prz, montada
    dentro do painel de filtros #app) + CSS + JS, injetado no fim do painel
    logo depois do HTML de resumo_acoes.py.
"""
from __future__ import annotations

import datetime as dt
import os
import shutil
import tempfile
import time
import unicodedata

import openpyxl

ABA_DADOS = "Controle de Prazos"
NOME_PLANILHA = "Controle de Prazos - Pauta Governador (Portos e Transporte).xlsx"
CAMINHO_ORIGEM_PADRAO = os.path.join(
    r"X:\COAPI 2\2026\19. APRESENTAÇÃO GABINETE - PAUTAS REUNIÕES", NOME_PLANILHA)
TENTATIVAS_LEITURA = 3

# colunas preenchidas à mão; uma linha só conta como ação se tiver alguma
# delas (as colunas de fórmula vêm preenchidas até a linha 500)
CAMPOS_CONTEUDO = (
    "ministerio", "eixo", "projeto", "descricao", "prazo_informado",
    "data_prazo", "status", "responsavel", "observacoes",
)
# cabeçalho da planilha (normalizado) -> chave interna
CABECALHOS = {
    "id": "id",
    "ministerio/orgao": "ministerio",
    "eixo": "eixo",
    "projeto/frente": "projeto",
    "descricao": "descricao",
    "prazo informado": "prazo_informado",
    "data prazo (ref.)": "data_prazo",
    "status": "status",
    "responsavel": "responsavel",
    "observacoes": "observacoes",
    "fonte (slide)": "fonte",
    "ultima atualizacao": "ultima_atualizacao",
}
COLUNAS_POSICIONAIS = [
    "id", "ministerio", "eixo", "projeto", "descricao", "prazo_informado",
    "data_prazo", "status", "responsavel", "_dias", "_alerta",
    "observacoes", "fonte", "ultima_atualizacao",
]


def caminho_origem() -> str:
    """Planilha original (variável CGAPE_PRAZOS_PLANILHA troca o caminho)."""
    return os.environ.get("CGAPE_PRAZOS_PLANILHA") or CAMINHO_ORIGEM_PADRAO


def caminho_reserva() -> str:
    pasta = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), "CGAPE")
    return os.path.join(pasta, "controle_prazos_reserva.xlsx")


def _normalizar(txt) -> str:
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return " ".join(txt.lower().split())


def _texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, dt.datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, dt.date):
        return valor.strftime("%d/%m/%Y")
    return str(valor).replace("\xa0", " ").strip()


def _data_iso(valor):
    if isinstance(valor, dt.datetime):
        return valor.date().isoformat()
    if isinstance(valor, dt.date):
        return valor.isoformat()
    return None


def _mapear_colunas(cabecalho) -> dict:
    indices = {}
    for i, nome in enumerate(cabecalho):
        chave = CABECALHOS.get(_normalizar(nome))
        if chave and chave not in indices:
            indices[chave] = i
    if "id" not in indices:
        return {c: i for i, c in enumerate(COLUNAS_POSICIONAIS)}
    return indices


def ler_registros(caminho_xlsx: str, padroes: bool = True) -> list[dict]:
    """Ações da planilha. padroes=False: campos vazios ficam vazios (usado
    na importação para o cadastro do sistema, que aplica os rótulos padrão
    só na hora de mostrar)."""
    regs = _ler_registros_planilha(caminho_xlsx)
    return [_com_padroes(r) for r in regs] if padroes else regs


def _com_padroes(r: dict) -> dict:
    r = dict(r)
    r["ministerio"] = r.get("ministerio") or "Sem órgão definido"
    r["eixo"] = r.get("eixo") or "Sem eixo"
    r["projeto"] = r.get("projeto") or "Sem projeto/frente"
    r["status"] = r.get("status") or "A definir"
    return r


def _ler_registros_planilha(caminho_xlsx: str) -> list[dict]:
    wb = openpyxl.load_workbook(caminho_xlsx, data_only=True, read_only=True)
    try:
        ws = wb[ABA_DADOS] if ABA_DADOS in wb.sheetnames else wb.worksheets[0]
        linhas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not linhas:
        return []
    indices = _mapear_colunas(linhas[0])
    registros, ids_usados = [], set()
    for num_linha, linha in enumerate(linhas[1:], start=2):
        if not linha:
            continue
        reg = {c: (linha[i] if i is not None and i < len(linha) else None)
               for c, i in ((c, indices.get(c)) for c in CABECALHOS.values())}
        if not any(_texto(reg.get(c)) for c in CAMPOS_CONTEUDO):
            continue  # linha só com as fórmulas preparadas
        # ação sem ID (ou com ID repetido): usa o número da linha
        id_reg = _texto(reg.get("id"))
        if not id_reg or id_reg in ids_usados:
            id_reg = f"linha {num_linha}"
        ids_usados.add(id_reg)
        registros.append({
            "id": id_reg,
            "ministerio": _texto(reg.get("ministerio")),
            "eixo": _texto(reg.get("eixo")),
            "projeto": _texto(reg.get("projeto")),
            "descricao": _texto(reg.get("descricao")),
            "prazo_informado": _texto(reg.get("prazo_informado")),
            "data_prazo": _data_iso(reg.get("data_prazo")),
            "status": _texto(reg.get("status")),
            "responsavel": _texto(reg.get("responsavel")),
            "observacoes": _texto(reg.get("observacoes")),
            "fonte": _texto(reg.get("fonte")),
            "ultima_atualizacao": _texto(reg.get("ultima_atualizacao")),
        })
    return registros


def _mtime(caminho):
    try:
        return os.path.getmtime(caminho)
    except OSError:
        return None


# =====================================================================
# CADASTRO NO PRÓPRIO SISTEMA (pedido do usuário em 25/11/2026): a planilha
# deixa de ser a fonte — as ações são criadas/editadas/excluídas na tela
# (só pelo .exe, na rede; só a Silvana edita, os demais visualizam — ver
# acesso.EDITORES_PRAZOS). Os dados ficam num arquivo JSON na MESMA pasta de
# rede da antiga planilha (todas as máquinas veem a mesma versão):
#   ...\PAUTAS REUNIÕES\Controle de Prazos - dados.json      (dados)
#   ...\PAUTAS REUNIÕES\Controle de Prazos - historico.jsonl (quem/quando/o quê)
#   ...\PAUTAS REUNIÕES\Controle de Prazos - backups\        (cópia a cada gravação)
# Na 1ª leitura sem o JSON, as ações vêm da planilha antiga (importação única).
# Cada gravação: trava (arquivo .lock), confere a revisão que a tela tinha
# (outra máquina pode ter gravado antes), grava num temporário e troca de uma
# vez (os.replace) — nunca deixa o arquivo pela metade.
# =====================================================================
import json

NOME_DADOS = "Controle de Prazos - dados.json"
NOME_HISTORICO = "Controle de Prazos - historico.jsonl"
NOME_BACKUPS = "Controle de Prazos - backups"
MAX_BACKUPS = 60
CAMPOS_EDITAVEIS = ("ministerio", "eixo", "projeto", "descricao", "prazo_informado",
                    "data_prazo", "status", "responsavel", "observacoes", "fonte")
LIMITE_TEXTO = 4000


def caminho_dados() -> str:
    """JSON do cadastro (variável CGAPE_PRAZOS_DADOS troca o caminho)."""
    return (os.environ.get("CGAPE_PRAZOS_DADOS")
            or os.path.join(os.path.dirname(caminho_origem()), NOME_DADOS))


def _pasta_dados() -> str:
    return os.path.dirname(caminho_dados())


def caminho_reserva_dados() -> str:
    return os.path.join(os.path.dirname(caminho_reserva()), "controle_prazos_reserva.json")


def _planilhas_para_importar():
    """Planilha antiga, na ordem: original na rede, reserva local, cópia na
    pasta do programa."""
    pasta_prog = os.path.dirname(os.path.abspath(__file__))
    return [p for p in (caminho_origem(), caminho_reserva(), os.path.join(pasta_prog, NOME_PLANILHA))
            if os.path.isfile(p)]


def _ler_json(caminho):
    for _ in range(TENTATIVAS_LEITURA):
        try:
            with open(caminho, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            time.sleep(0.4)
    raise OSError(f"não foi possível ler {caminho}")


def _gravar_json_atomico(caminho, dados):
    tmp = caminho + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, caminho)


def _importar_planilha():
    """Dados iniciais a partir da planilha antiga (ids "linha N" viram números)."""
    for planilha in _planilhas_para_importar():
        try:
            regs = ler_registros(planilha, padroes=False)
        except Exception:
            continue
        usados = {int(r["id"]) for r in regs if str(r["id"]).isdigit()}
        prox = max(usados or {0}) + 1
        for r in regs:
            if not str(r["id"]).isdigit():
                r["id"] = str(prox)
                prox += 1
        return {"revisao": 1, "importado_de": os.path.basename(planilha),
                "importado_em": dt.datetime.now().strftime("%d/%m/%Y %H:%M"), "itens": regs}
    return None


def _espelhar_reserva(dados):
    try:
        os.makedirs(os.path.dirname(caminho_reserva_dados()), exist_ok=True)
        _gravar_json_atomico(caminho_reserva_dados(), dados)
    except OSError:
        pass


def _carregar():
    """(dados, caminho lido, aviso, editavel). editavel = o JSON da rede está
    acessível (só nele se grava; reserva/planilha são só para leitura)."""
    principal = caminho_dados()
    if os.path.isfile(principal):
        dados = _ler_json(principal)
        _espelhar_reserva(dados)
        return dados, principal, "", True
    if os.path.isdir(_pasta_dados()):
        # 1ª vez: importa a planilha antiga e cria o cadastro na rede
        dados = _importar_planilha() or {"revisao": 1, "itens": []}
        with _Trava():
            if not os.path.isfile(principal):
                _gravar_json_atomico(principal, dados)
                _registrar_historico("importar", "", None, None,
                                     {"origem": dados.get("importado_de", ""), "acoes": len(dados["itens"])})
        dados = _ler_json(principal)
        _espelhar_reserva(dados)
        return dados, principal, "", True
    if os.path.isfile(caminho_reserva_dados()):
        return (_ler_json(caminho_reserva_dados()), caminho_reserva_dados(),
                "pasta da rede inacessível — exibindo a última cópia (somente leitura)", False)
    dados = _importar_planilha()
    if dados:
        return dados, "", "pasta da rede inacessível — exibindo a planilha antiga (somente leitura)", False
    raise FileNotFoundError(f"Cadastro do Controle de Prazos não encontrado: {principal}")


def data_atualizacao_base():
    """Data da última gravação do cadastro (ou da cópia local)."""
    m = (_mtime(caminho_dados()) or _mtime(caminho_reserva_dados())
         or _mtime(caminho_origem()) or _mtime(caminho_reserva()))
    return dt.datetime.fromtimestamp(m) if m else None


def api_dados(pode_editar=False):
    try:
        dados, caminho, aviso, editavel = _carregar()
        m = _mtime(caminho) if caminho else None
        return {
            "ok": True,
            "itens": [_com_padroes(r) for r in dados.get("itens", [])],
            "aviso": aviso,
            "atualizado_em": dt.datetime.fromtimestamp(m).strftime("%d/%m/%Y às %Hh%Mmin") if m else "",
            "assinatura": dados.get("revisao"),
            "revisao": dados.get("revisao"),
            "pode_editar": bool(pode_editar and editavel),
            "editor_sem_rede": bool(pode_editar and not editavel),
            "opcoes": _opcoes(dados.get("itens", [])),
        }
    except Exception as exc:
        return {"ok": False, "mensagem": f"Erro ao ler o Controle de Prazos: {exc}"}


def api_assinatura():
    for caminho in (caminho_dados(), caminho_reserva_dados()):
        if os.path.isfile(caminho):
            try:
                return {"ok": True, "assinatura": _ler_json(caminho).get("revisao")}
            except OSError:
                break
    return {"ok": True, "assinatura": None}


def _opcoes(itens):
    """Valores já usados, para as listas do formulário."""
    def unicos(chave, base=()):
        vals = {str(r.get(chave) or "").strip() for r in itens} | set(base)
        return sorted((v for v in vals if v), key=lambda v: _normalizar(v))
    return {"ministerio": unicos("ministerio"), "eixo": unicos("eixo"), "projeto": unicos("projeto"),
            "responsavel": unicos("responsavel"),
            "status": unicos("status", ("A definir", "Aguardando", "Em andamento", "Concluído"))}


class _Trava:
    """Trava entre máquinas: arquivo .lock criado em modo exclusivo; trava
    esquecida (programa fechado no meio) expira em 30 s."""
    def __init__(self, espera=10.0):
        self.caminho = caminho_dados() + ".lock"
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
                idade = time.time() - (_mtime(self.caminho) or time.time())
                if idade > 30:
                    try:
                        os.remove(self.caminho)
                    except OSError:
                        pass
                    continue
                if time.time() > limite:
                    raise TimeoutError("outra pessoa está gravando o Controle de Prazos agora — tente de novo")
                time.sleep(0.25)

    def __exit__(self, *exc):
        try:
            os.close(self.fd)
        finally:
            try:
                os.remove(self.caminho)
            except OSError:
                pass


def _registrar_historico(acao, usuario, id_reg, antes, depois):
    linha = {"quando": dt.datetime.now().isoformat(timespec="seconds"), "usuario": usuario,
             "acao": acao, "id": id_reg, "antes": antes, "depois": depois}
    try:
        with open(os.path.join(_pasta_dados(), NOME_HISTORICO), "a", encoding="utf-8") as f:
            f.write(json.dumps(linha, ensure_ascii=False) + "\n")
    except OSError:
        pass


def _backup():
    origem = caminho_dados()
    if not os.path.isfile(origem):
        return
    pasta = os.path.join(_pasta_dados(), NOME_BACKUPS)
    try:
        os.makedirs(pasta, exist_ok=True)
        shutil.copy2(origem, os.path.join(pasta, f"dados_{dt.datetime.now():%Y%m%d_%H%M%S_%f}.json"))
        antigos = sorted(f for f in os.listdir(pasta) if f.startswith("dados_"))
        for f in antigos[:-MAX_BACKUPS]:
            os.remove(os.path.join(pasta, f))
    except OSError:
        pass


def _limpar_item(item: dict) -> dict:
    limpo = {}
    for c in CAMPOS_EDITAVEIS:
        v = item.get(c)
        v = "" if v is None else " ".join(str(v).split()) if c != "observacoes" else str(v).strip()
        limpo[c] = v[:LIMITE_TEXTO]
    d = limpo["data_prazo"]
    if d:
        try:
            limpo["data_prazo"] = dt.date.fromisoformat(d[:10]).isoformat()
        except ValueError:
            raise ValueError("data do prazo inválida")
    else:
        limpo["data_prazo"] = None
    if not limpo["descricao"]:
        raise ValueError("preencha a descrição da ação")
    if not limpo["projeto"]:
        raise ValueError("preencha o projeto/frente")
    return limpo


def _alterar(revisao_tela, usuario, nome, funcao):
    """Aplica `funcao(itens) -> (acao, id, antes, depois)` sob a trava,
    conferindo a revisão; devolve a resposta da API."""
    if not os.path.isfile(caminho_dados()):
        return {"ok": False, "mensagem": "Cadastro do Controle de Prazos inacessível (pasta da rede fora do ar)."}
    with _Trava():
        dados = _ler_json(caminho_dados())
        if revisao_tela is not None and dados.get("revisao") != revisao_tela:
            return {"ok": False, "conflito": True,
                    "mensagem": "Os dados foram alterados em outra máquina desde que a tela foi aberta. "
                                "A tela foi recarregada — confira e grave de novo."}
        acao, id_reg, antes, depois = funcao(dados["itens"])
        _backup()
        dados["revisao"] = int(dados.get("revisao") or 0) + 1
        dados["alterado_em"] = dt.datetime.now().strftime("%d/%m/%Y %H:%M")
        dados["alterado_por"] = nome
        _gravar_json_atomico(caminho_dados(), dados)
        _espelhar_reserva(dados)
        _registrar_historico(acao, usuario, id_reg, antes, depois)
    return {"ok": True, "id": id_reg, "revisao": dados["revisao"]}


def api_salvar(item, revisao, usuario, nome):
    """Cria (sem "id") ou edita (com "id") uma ação."""
    try:
        item = item or {}
        limpo = _limpar_item(item)
        carimbo = f"{dt.datetime.now():%d/%m/%Y %H:%M} · {nome}"

        def aplicar(itens):
            id_reg = str(item.get("id") or "").strip()
            if id_reg:
                for i, r in enumerate(itens):
                    if str(r.get("id")) == id_reg:
                        novo = dict(r, **limpo, ultima_atualizacao=carimbo)
                        itens[i] = novo
                        return "editar", id_reg, r, novo
                raise ValueError("ação não encontrada (pode ter sido excluída em outra máquina)")
            ids = [int(r["id"]) for r in itens if str(r.get("id", "")).isdigit()]
            novo = dict(limpo, id=str(max(ids or [0]) + 1), ultima_atualizacao=carimbo,
                        criado_em=f"{dt.datetime.now():%d/%m/%Y %H:%M}", criado_por=nome)
            itens.append(novo)
            return "criar", novo["id"], None, novo

        return _alterar(revisao, usuario, nome, aplicar)
    except (ValueError, TimeoutError, OSError) as exc:
        return {"ok": False, "mensagem": str(exc)}


def api_excluir(id_reg, revisao, usuario, nome):
    try:
        def aplicar(itens):
            for i, r in enumerate(itens):
                if str(r.get("id")) == str(id_reg):
                    itens.pop(i)
                    return "excluir", str(id_reg), r, None
            raise ValueError("ação não encontrada (pode ter sido excluída em outra máquina)")
        return _alterar(revisao, usuario, nome, aplicar)
    except (ValueError, TimeoutError, OSError) as exc:
        return {"ok": False, "mensagem": str(exc)}


def planilha_exportacao_bytes():
    """Excel com as ações do cadastro (mesmas colunas da planilha antiga,
    com Dias e Alerta calculados hoje) — só consulta, não volta para o sistema."""
    import io
    from openpyxl.styles import Alignment, Font, PatternFill
    dados, _c, _a, _e = _carregar()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = ABA_DADOS
    cab = ["ID", "Ministério/Órgão", "Eixo", "Projeto/Frente", "Descrição", "Prazo informado",
           "Data prazo (ref.)", "Status", "Responsável", "Dias", "Alerta", "Observações",
           "Fonte (slide)", "Última atualização"]
    ws.append(cab)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="C62828")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    hoje = dt.date.today()
    for r in sorted(dados.get("itens", []), key=lambda x: int(x["id"]) if str(x["id"]).isdigit() else 10**9):
        d = dt.date.fromisoformat(r["data_prazo"]) if r.get("data_prazo") else None
        concl = _normalizar(r.get("status")).startswith("conclu")
        dias = (d - hoje).days if d else None
        alerta = ("Concluído" if concl else "Sem data" if d is None else "Vencido" if dias < 0
                  else "Crítico" if dias <= 7 else "Atenção" if dias <= 30 else "OK")
        ws.append([int(r["id"]) if str(r["id"]).isdigit() else r["id"], r.get("ministerio"), r.get("eixo"),
                   r.get("projeto"), r.get("descricao"), r.get("prazo_informado"), d, r.get("status"),
                   r.get("responsavel"), None if concl else dias, alerta, r.get("observacoes"),
                   r.get("fonte"), r.get("ultima_atualizacao")])
        ws.cell(ws.max_row, 7).number_format = "DD/MM/YYYY"
    for col, larg in zip("ABCDEFGHIJKLMN", (6, 28, 16, 28, 50, 16, 13, 14, 20, 7, 11, 50, 14, 22)):
        ws.column_dimensions[col].width = larg
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---- HTML / CSS / JS (injetado no painel antes do </body>) ------------------
HTML_JANELA = r"""
<div id="prz-overlay" aria-hidden="true">
  <div id="prz-janela" role="dialog" aria-label="Controle de prazos">
    <div class="ra-cabecalho prz-cabecalho">
      <h2>CONTROLE DE PRAZOS
        <small>Pauta do Governador · Portos e Transporte<span id="przAtualizado"></span><span id="przAviso" class="prz-aviso-base"></span></small></h2>
      <div class="ra-botoes">
        <div class="seletor-secao troca-base chave-prz" data-tela="dash" role="tablist" title="Seção ativa (PAC, MONITORAMENTO ou PRAZOS). Atalho: Alt+1 / Alt+2 / Alt+3.">
          <button type="button" class="btn-casa" id="przInicio" title="Início (voltar à tela inicial)" aria-label="Início — voltar à tela inicial"><svg viewBox="0 0 24 24" width="17" height="17" fill="currentColor" aria-hidden="true"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"></path></svg></button>
          <button type="button" data-secao="PAC" role="tab" aria-selected="false" tabindex="-1">PAC<span class="contagem-base"></span></button>
          <button type="button" data-secao="MON" role="tab" aria-selected="false" tabindex="-1">MONITORAMENTO<span class="contagem-base"></span></button>
          <button type="button" class="ativo" data-secao="PRZ" role="tab" aria-selected="true" tabindex="0">PRAZOS<span class="contagem-base"></span></button>
        </div>
        <div class="ra-busca-geral prz-busca-geral">
          <input type="search" id="przBuscaGeral" autocomplete="off"
            placeholder="Buscar em todas as colunas (órgão, projeto, ação, responsável…)"
            title="Mostra só as ações que têm TODAS as palavras digitadas em qualquer campo das ações do Controle de Prazos">
          <button type="button" id="przBuscaGeralLimpar" title="Limpar a busca">×</button>
        </div>
        <button type="button" class="btn-inicio" id="przLimparBase" title="Limpar filtros (PRAZOS)">LIMPAR FILTROS (PRAZOS)</button>
        <!-- cadastro no sistema (25/11/2026): NOVA AÇÃO só para quem edita -->
        <button type="button" class="prz-btn-nova" id="przNova" hidden title="Cadastrar uma nova ação no Controle de Prazos">+ NOVA AÇÃO</button>
        <button type="button" id="przExportar" title="Baixa um Excel com as ações cadastradas (só para consulta — a fonte é o sistema)">EXPORTAR EXCEL</button>
        <button type="button" class="botao-tema ra-botao-tema" id="przTema" title="Tema escuro" aria-label="Ativar tema escuro" aria-pressed="false"><span class="icone-lua" aria-hidden="true"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg></span><span class="icone-sol" aria-hidden="true"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4.5"></circle><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"></path></svg></span></button>
        <span class="prz-relogio" title="Data e hora de agora (os alertas são recalculados ao vivo)"><span>HOJE</span><b id="przRelogio"></b></span>
      </div>
    </div>
    <div class="prz-area">
      <div class="prz-corpo" id="przCorpo">
        <div class="prz-msg" id="przMsg">Carregando o Controle de Prazos…</div>
        <div id="przConteudo" hidden>
          <div class="prz-destaques" id="przDestaques"></div>
          <div class="prz-radar">
            <div class="prz-radar-rot">PRÓXIMOS PRAZOS</div>
            <div class="prz-radar-cards" id="przRadar"></div>
          </div>
          <div class="prz-abas" role="tablist">
            <button type="button" class="prz-aba ativa" role="tab" aria-selected="true" data-visao="tempo">Linha do tempo</button>
            <button type="button" class="prz-aba" role="tab" aria-selected="false" data-visao="quadro">Quadro de alertas</button>
            <button type="button" class="prz-aba" role="tab" aria-selected="false" data-visao="mapa">Mapa mental</button>
            <span class="prz-abas-dica" id="przAbasDica"></span>
            <button type="button" class="prz-abas-btn" id="przMapaTudo" hidden>Expandir / recolher tudo</button>
          </div>
          <div class="prz-visao ativa" id="przVisTempo"></div>
          <div class="prz-visao" id="przVisQuadro"></div>
          <div class="prz-visao prz-viewport" id="przVisMapa">
            <svg class="prz-mapa-linhas" id="przMapaSvg"></svg>
            <div class="prz-mapa-canvas" id="przMapaCanvas"></div>
            <div class="prz-zoom">
              <button type="button" id="przZoomMais" title="Aproximar">+</button>
              <button type="button" id="przZoomMenos" title="Afastar">&minus;</button>
              <button type="button" id="przZoomCaber" title="Enquadrar tudo">&#8862;</button>
            </div>
          </div>
        </div>
      </div>
      <aside class="prz-painel" id="przPainel" aria-hidden="true">
        <button type="button" class="prz-painel-fechar" id="przPainelFechar" title="Fechar" aria-label="Fechar">×</button>
        <div id="przPainelConteudo"></div>
      </aside>
    </div>
  </div>
  <div class="prz-toast" id="przToast" role="status"></div>
  <!-- Formulário de cadastro/edição (25/11/2026) -->
  <div class="prz-form-fundo" id="przFormFundo" aria-hidden="true">
    <form class="prz-form" id="przForm" role="dialog" aria-modal="true" aria-labelledby="przFormTitulo" novalidate>
      <div class="prz-form-cab">
        <h3 id="przFormTitulo">Nova ação</h3>
        <button type="button" class="prz-form-x" id="przFormFechar" title="Fechar (Esc)" aria-label="Fechar">×</button>
      </div>
      <div class="prz-form-corpo">
        <label class="largo"><span>Descrição da ação <em>*</em></span>
          <textarea name="descricao" rows="2" required maxlength="4000" placeholder="O que precisa acontecer (ex.: assinatura do convênio…)"></textarea></label>
        <label>Ministério / Órgão
          <input name="ministerio" list="przListaMinisterio" autocomplete="off" maxlength="300"></label>
        <label>Eixo
          <input name="eixo" list="przListaEixo" autocomplete="off" maxlength="300"></label>
        <label class="largo"><span>Projeto / Frente <em>*</em></span>
          <input name="projeto" list="przListaProjeto" autocomplete="off" required maxlength="300"></label>
        <label>Data do prazo
          <input name="data_prazo" type="date"></label>
        <label>Prazo informado (texto original)
          <input name="prazo_informado" maxlength="300" placeholder="ex.: 2º semestre/2026, a definir…"></label>
        <label>Status
          <select name="status" id="przFormStatus"></select></label>
        <label>Responsável
          <input name="responsavel" list="przListaResponsavel" autocomplete="off" maxlength="300"></label>
        <label class="largo">Observações
          <textarea name="observacoes" rows="4" maxlength="4000"></textarea></label>
        <label class="largo">Fonte
          <input name="fonte" maxlength="300" placeholder="ex.: slide da pauta, ofício, reunião de 10/11…"></label>
      </div>
      <div class="prz-form-erro" id="przFormErro" role="alert"></div>
      <div class="prz-form-rodape">
        <span class="prz-form-dica">* obrigatório · o alerta é calculado pela data do prazo e pelo status</span>
        <button type="button" class="prz-form-cancelar" id="przFormCancelar">Cancelar</button>
        <button type="submit" class="prz-form-salvar" id="przFormSalvar">Salvar</button>
      </div>
      <datalist id="przListaMinisterio"></datalist><datalist id="przListaEixo"></datalist>
      <datalist id="przListaProjeto"></datalist><datalist id="przListaResponsavel"></datalist>
    </form>
  </div>
</div>
<style>
  /* ===== CONTROLE DE PRAZOS (25/11/2026) — mesmo esqueleto do dash do
     MONITORAMENTO (.ra-cabecalho/.ra-botoes de resumo_acoes.py) com a cor
     VERMELHA da seção (--cor-acento-prazos, definida no :root do painel). */
  #prz-overlay { display:none; position:fixed; inset:0; z-index:70;
    background:var(--cor-fundo); color:var(--cor-texto-primario); font-family:var(--fonte);
    --prz-venc:#C62828; --prz-crit:#E65100; --prz-aten:#C99400; --prz-ok:#2E7D32;
    --prz-sem:#8C97A6; --prz-concl:#0B5CAD; }
  html[data-tema="escuro"] #prz-overlay { --prz-venc:#EF5350; --prz-crit:#FF8A3D; --prz-aten:#F2C14E;
    --prz-ok:#4CAF50; --prz-sem:#8D9CB2; --prz-concl:#5B9BE6; }
  #prz-overlay.aberto { display:flex; }
  #prz-janela { display:flex; flex-direction:column; width:100%; height:100%; }
  #prz-overlay .ra-cabecalho { border-bottom-color:var(--cor-acento-prazos); }
  #prz-overlay .ra-botoes > button:hover:not(:disabled) { background:var(--cor-acento-prazos-hover); color:#fff; }
  #prz-overlay .ra-botoes > .ra-botao-tema:hover:not(:disabled) { background:var(--cor-acento-prazos); color:#fff; }
  #prz-overlay .ra-botoes .seletor-secao.troca-base button.ativo { background:var(--cor-acento-prazos); color:#fff; }
  #przBuscaGeral { width:100%; box-sizing:border-box; background:var(--cor-card);
    color:var(--cor-texto-primario); border:1px solid var(--cor-card-elevado);
    border-radius:8px; padding:7px 36px 7px 38px; font-size:13px; font-family:inherit;
    box-shadow:var(--sombra-card); outline:none;
    background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%23889' stroke-width='2.4' stroke-linecap='round'%3E%3Ccircle cx='11' cy='11' r='7'/%3E%3Cpath d='M20 20l-4-4'/%3E%3C/svg%3E");
    background-repeat:no-repeat; background-position:12px center; }
  #przBuscaGeral::-webkit-search-cancel-button { display:none; }
  #przBuscaGeral:focus, .prz-busca-geral.ativo #przBuscaGeral { border-color:var(--cor-acento-prazos); }
  #przBuscaGeralLimpar { position:absolute; right:8px; top:50%; transform:translateY(-50%);
    background:none; border:none; color:var(--cor-texto-secundario); font-size:18px;
    line-height:1; cursor:pointer; padding:2px 6px; display:none; border-radius:0; }
  .prz-busca-geral.ativo #przBuscaGeralLimpar { display:block; }
  #prz-overlay .ra-botoes > #przBuscaGeralLimpar:hover { background:none; }
  .prz-aviso-base:not(:empty) { margin-left:8px; padding:1px 8px; border-radius:999px;
    background:var(--cor-acento-prazos); color:#fff; font-weight:600; }
  .prz-relogio { margin-left:auto; text-align:right; font-size:10.5px; letter-spacing:.5px;
    color:var(--cor-texto-terciario); line-height:1.2; white-space:nowrap; }
  .prz-relogio b { display:block; font-size:13px; letter-spacing:0; color:var(--cor-texto-primario);
    font-variant-numeric:tabular-nums; }

  .prz-area { flex:1; min-height:0; position:relative; display:flex; }
  /* topo 62px: espaço do botão flutuante do painel de filtros (#fab-filtros) */
  .prz-corpo { flex:1; overflow:auto; padding:62px 18px 40px; }
  .prz-msg { font-size:13px; color:var(--cor-texto-secundario); min-height:18px; margin-bottom:10px; }
  .prz-msg:empty { display:none; }

  /* destaques: total do recorte + um cartão por alerta (clique filtra) */
  .prz-destaques { display:grid; grid-template-columns:minmax(200px,1.25fr) repeat(6, minmax(120px,1fr));
    gap:12px; margin-bottom:14px; }
  .prz-kpi { background:var(--cor-card); border-radius:8px; padding:12px 16px; box-shadow:var(--sombra-card);
    border:1px solid transparent; border-left:5px solid var(--c, var(--cor-acento-prazos));
    text-align:left; font:inherit; color:inherit; min-width:0; }
  button.prz-kpi { cursor:pointer; transition:transform .15s, border-color .15s; }
  button.prz-kpi:hover, button.prz-kpi:focus-visible { transform:translateY(-2px); border-color:var(--c); outline:none; }
  button.prz-kpi.ativo { border-color:var(--c); box-shadow:0 0 0 2px var(--c), var(--sombra-card); }
  button.prz-kpi.apagado { opacity:.45; }
  .prz-kpi span { display:block; font-size:12px; font-weight:700; letter-spacing:.4px; text-transform:uppercase;
    color:var(--cor-texto-secundario); }
  .prz-kpi b { display:block; font-size:30px; font-weight:900; line-height:1.15; color:var(--c); margin:2px 0;
    font-variant-numeric:tabular-nums; }
  .prz-kpi small { display:block; font-size:11.5px; color:var(--cor-texto-terciario); }
  .prz-kpi-total b { font-size:40px; color:var(--cor-texto-primario); }

  .prz-radar { display:flex; align-items:stretch; gap:10px; margin-bottom:16px; overflow-x:auto; padding-bottom:4px; }
  .prz-radar-rot { flex:0 0 auto; align-self:center; font-size:12px; font-weight:800; letter-spacing:.5px;
    color:var(--cor-acento-prazos-texto); writing-mode:horizontal-tb; }
  .prz-radar-cards { display:flex; gap:10px; }
  .prz-radar-card { flex:0 0 auto; min-width:200px; max-width:260px; background:var(--cor-card);
    border:1px solid var(--cor-card-elevado); border-left:5px solid var(--c); border-radius:8px;
    padding:8px 12px; cursor:pointer; text-align:left; font:inherit; color:inherit; transition:transform .12s; }
  .prz-radar-card:hover, .prz-radar-card:focus-visible { transform:translateY(-2px); box-shadow:var(--sombra-card);
    border-color:var(--c); outline:none; }
  .prz-radar-card .p { display:block; font-weight:700; font-size:12.5px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .prz-radar-card .o { display:block; color:var(--cor-texto-secundario); font-size:11px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .prz-radar-card .c { display:block; font-weight:800; margin-top:3px; color:var(--c); font-variant-numeric:tabular-nums; font-size:12.5px; }
  .prz-radar-vazio { font-size:12.5px; color:var(--cor-texto-secundario); align-self:center; }

  .prz-abas { display:flex; align-items:center; gap:2px; background:var(--cor-card); border-radius:8px 8px 0 0;
    border-bottom:1px solid var(--cor-card-elevado); padding:0 10px; }
  .prz-aba { border:none; background:none; color:var(--cor-texto-secundario); font:inherit; font-size:13px;
    font-weight:700; padding:11px 14px 9px; border-bottom:3px solid transparent; cursor:pointer; }
  .prz-aba:hover { color:var(--cor-texto-primario); }
  .prz-aba.ativa { color:var(--cor-texto-primario); border-bottom-color:var(--cor-acento-prazos); }
  .prz-aba:focus-visible { outline:2px solid var(--cor-acento-prazos); outline-offset:-2px; }
  .prz-abas-dica { margin-left:auto; font-size:11.5px; color:var(--cor-texto-terciario); }
  .prz-abas-btn { font:inherit; font-size:12px; font-weight:600; cursor:pointer; margin-left:12px;
    border:1px solid var(--cor-card-elevado); background:transparent; color:var(--cor-texto-primario);
    border-radius:6px; padding:5px 12px; }
  .prz-abas-btn:hover { border-color:var(--cor-acento-prazos); color:var(--cor-acento-prazos-texto); }
  .prz-visao { display:none; background:var(--cor-card); border-radius:0 0 8px 8px; box-shadow:var(--sombra-card); }
  .prz-visao.ativa { display:block; }
  .prz-vazio { padding:48px; text-align:center; color:var(--cor-texto-secundario); font-size:13px; }

  /* ---- linha do tempo ---- */
  #przVisTempo { overflow:auto; max-height:calc(100vh - 330px); min-height:360px; }
  .prz-tl { position:relative; min-height:100%; }
  .prz-tl-linha { display:flex; min-height:50px; border-bottom:1px solid var(--cor-card-elevado); position:relative; z-index:1; }
  .prz-tl-rot { flex:0 0 340px; width:340px; position:sticky; left:0; z-index:2; background:var(--cor-card);
    border-right:1px solid var(--cor-card-elevado); border-left:4px solid transparent; padding:7px 14px 7px 12px; }
  .prz-tl-trilha { flex:1 1 auto; position:relative; }
  .prz-tl-cab { position:sticky; top:0; z-index:4; }
  .prz-tl-cab .prz-tl-rot { z-index:5; background:var(--cor-card-elevado); }
  .prz-tl-cab .prz-tl-trilha { background:var(--cor-card-elevado); height:56px; }
  .prz-tl-cab-tit { font-size:11.5px; font-weight:800; text-transform:uppercase; letter-spacing:.5px; color:var(--cor-texto-secundario); }
  .prz-tl-cab-res { font-size:12px; margin-top:5px; color:var(--cor-texto-secundario); }
  .prz-tl-cab-res b { font-size:14px; font-variant-numeric:tabular-nums; }
  .prz-tl-mes { position:absolute; top:0; height:26px; line-height:26px; padding-left:8px; font-size:11px; font-weight:700;
    color:var(--cor-texto-secundario); text-transform:uppercase; letter-spacing:.5px;
    border-left:1px solid var(--cor-texto-terciario); white-space:nowrap; overflow:hidden; }
  .prz-tl-zfaixa { position:absolute; top:31px; height:18px; line-height:18px; padding-left:5px; font-size:10px; font-weight:700;
    white-space:nowrap; overflow:hidden; color:#fff; }
  .prz-tl-zfaixa.critica { background:var(--prz-crit); }
  .prz-tl-zfaixa.atencao { background:color-mix(in srgb, var(--prz-aten) 85%, transparent); color:#1b1b1b; }
  .prz-tl-zrot-hoje { position:absolute; top:31px; height:18px; line-height:18px; transform:translateX(-100%); margin-left:-1px;
    background:var(--cor-acento-prazos); color:#fff; font-size:10px; font-weight:800; padding:0 7px;
    border-radius:999px 0 0 999px; white-space:nowrap; }
  .prz-tl-fundo { position:absolute; top:0; bottom:0; left:340px; right:0; pointer-events:none; z-index:0; }
  .prz-tl-zona { position:absolute; top:0; bottom:0; }
  .prz-tl-zona.passado { background:color-mix(in srgb, var(--prz-venc) 5%, transparent); }
  .prz-tl-zona.critica { background:color-mix(in srgb, var(--prz-crit) 12%, transparent); }
  .prz-tl-zona.atencao { background:color-mix(in srgb, var(--prz-aten) 9%, transparent); }
  .prz-tl-grade { position:absolute; top:0; bottom:0; width:1px; background:var(--cor-card-elevado); }
  .prz-tl-grade.meio { opacity:.5; }
  .prz-tl-hoje { position:absolute; top:0; bottom:0; width:2px; margin-left:-1px; background:var(--cor-acento-prazos); }
  .prz-tl-grupo { min-height:34px; }
  .prz-tl-grupo .prz-tl-rot { background:var(--cor-card-elevado); font-weight:800; font-size:12.5px;
    border-left-color:var(--cor-acento-prazos); display:flex; align-items:center; flex-wrap:wrap; gap:4px 6px; }
  .prz-tl-grupo .resumo { font-weight:500; color:var(--cor-texto-secundario); font-size:11px; }
  .prz-tl-linha.item { cursor:pointer; }
  .prz-tl-linha.item:hover .prz-tl-rot, .prz-tl-linha.item:focus-visible .prz-tl-rot { background:var(--cor-card-elevado); }
  .prz-tl-linha.item:hover .prz-tl-trilha { background:color-mix(in srgb, var(--cor-texto-primario) 4%, transparent); }
  .prz-tl-linha.item:focus-visible { outline:2px solid var(--cor-acento-prazos); outline-offset:-2px; }
  .prz-tl-desc { font-size:12.5px; font-weight:600; line-height:1.3; display:-webkit-box; -webkit-line-clamp:2;
    -webkit-box-orient:vertical; overflow:hidden; color:var(--cor-texto-primario); }
  .prz-tl-sub { font-size:11px; color:var(--cor-texto-secundario); margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .prz-tl-linha.feito .prz-tl-desc { color:var(--cor-texto-secundario); }
  .prz-tl-marco { position:absolute; top:50%; width:14px; height:14px; margin:-7px 0 0 -7px; background:var(--c);
    transform:rotate(45deg); border-radius:2px; box-shadow:0 0 0 2px var(--cor-card); z-index:2; }
  .prz-tl-marco.feito { border-radius:50%; transform:none; opacity:.75; }
  .prz-tl-marco.pulsa::after { content:''; position:absolute; inset:-5px; border:2px solid var(--c); border-radius:3px;
    animation:przPulso 1.8s ease-out infinite; }
  @keyframes przPulso { from { opacity:.9; transform:scale(.7); } to { opacity:0; transform:scale(1.6); } }
  .prz-tl-barra { position:absolute; top:50%; height:6px; margin-top:-3px; border-radius:3px; }
  .prz-tl-barra.atraso { background:repeating-linear-gradient(135deg, var(--prz-venc) 0 6px,
    color-mix(in srgb, var(--prz-venc) 35%, transparent) 6px 12px); }
  .prz-tl-barra.futura { background:var(--c); opacity:.45; }
  .prz-tl-txt { position:absolute; top:50%; transform:translateY(-50%); font-size:11.5px; font-weight:700; white-space:nowrap;
    font-variant-numeric:tabular-nums; padding:1px 8px; border-radius:999px; background:var(--cor-card);
    border:1px solid var(--c); color:var(--c); z-index:3; }
  .prz-tl-semdata { position:absolute; top:50%; transform:translateY(-50%); left:14px; font-size:11.5px;
    color:var(--cor-texto-secundario); font-style:italic; border:1px dashed var(--cor-texto-terciario);
    padding:2px 10px; border-radius:999px; background:var(--cor-card); white-space:nowrap; }

  /* ---- quadro de alertas ---- */
  #przVisQuadro { overflow:auto; padding:14px; }
  .prz-qd { display:grid; grid-template-columns:repeat(6, minmax(200px, 1fr)); gap:12px; align-items:start; min-width:1250px; }
  .prz-qd-col { background:var(--cor-fundo); border:1px solid var(--cor-card-elevado); border-top:4px solid var(--c); border-radius:10px; }
  .prz-qd-cab { padding:10px 12px 8px; }
  .prz-qd-cab .t { display:flex; align-items:center; justify-content:space-between; font-weight:800; font-size:13.5px; }
  .prz-qd-cab .n { background:var(--c); color:#fff; border-radius:999px; padding:1px 9px; font-size:12px; font-variant-numeric:tabular-nums; }
  .prz-qd-cab .d { font-size:11px; color:var(--cor-texto-secundario); margin-top:2px; }
  .prz-qd-cards { padding:2px 10px 10px; display:flex; flex-direction:column; gap:8px; }
  .prz-qd-card { background:var(--cor-card); border:1px solid var(--cor-card-elevado); border-radius:8px; padding:9px 11px;
    cursor:pointer; transition:transform .12s, box-shadow .12s, border-color .12s; }
  .prz-qd-card:hover, .prz-qd-card:focus-visible { transform:translateY(-2px); box-shadow:var(--sombra-card);
    border-color:var(--c); outline:none; }
  .prz-qd-num { display:flex; align-items:baseline; gap:6px; color:var(--c); font-variant-numeric:tabular-nums; }
  .prz-qd-num b { font-size:22px; line-height:1.1; }
  .prz-qd-num span { font-size:11px; font-weight:600; }
  .prz-qd-desc { font-size:12.5px; font-weight:600; line-height:1.3; margin:4px 0 3px; color:var(--cor-texto-primario); }
  .prz-qd-meta { font-size:11px; color:var(--cor-texto-secundario); line-height:1.4; }
  .prz-qd-medidor { height:5px; border-radius:3px; background:var(--cor-card-elevado); margin-top:8px; overflow:hidden; }
  .prz-qd-medidor i { display:block; height:100%; background:var(--c); border-radius:3px; }
  .prz-qd-vazio { font-size:12px; color:var(--cor-texto-secundario); padding:4px 2px 6px; font-style:italic; }

  /* ---- mapa mental ---- */
  .prz-viewport { position:relative; overflow:hidden; cursor:grab; height:calc(100vh - 330px); min-height:420px;
    background:radial-gradient(circle, color-mix(in srgb, var(--cor-texto-terciario) 30%, transparent) 1px, transparent 1px) 0 0/22px 22px,
      var(--cor-fundo); }
  .prz-viewport.arrastando { cursor:grabbing; }
  .prz-mapa-canvas { position:absolute; top:0; left:0; transform-origin:0 0; }
  .prz-mapa-linhas { position:absolute; top:0; left:0; overflow:visible; pointer-events:none; transform-origin:0 0; }
  .prz-mapa-linhas path { stroke:var(--cor-texto-terciario); stroke-opacity:.55; }
  .prz-no { position:absolute; background:var(--cor-card); border:1px solid var(--cor-card-elevado); border-radius:10px;
    padding:8px 12px; box-shadow:var(--sombra-card); user-select:none; color:var(--cor-texto-primario);
    border-left:5px solid var(--cor-texto-terciario); }
  .prz-no.raiz { background:var(--cor-acento-prazos); color:#fff; border:none; padding:14px 20px; width:260px; cursor:pointer; }
  .prz-no.raiz .tit { font-size:15px; font-weight:800; letter-spacing:.3px; }
  .prz-no.raiz .sub { font-size:11px; opacity:.85; margin-top:2px; }
  .prz-no.ministerio { border-left-color:var(--cor-texto-primario); font-weight:800; font-size:13px; cursor:pointer; width:220px; }
  .prz-no.eixo { border-left-color:var(--cor-acento-teal); font-weight:700; font-size:12.5px; cursor:pointer; width:200px; }
  .prz-no.projeto { font-weight:700; font-size:12.5px; cursor:pointer; width:230px; }
  .prz-no.item { font-size:12px; width:250px; cursor:pointer; }
  .prz-no.item .desc { font-weight:600; margin-bottom:4px; line-height:1.35; }
  .prz-no .meta { color:var(--cor-texto-secundario); font-size:11px; margin-bottom:4px; font-weight:400; }
  .prz-no .seta { position:absolute; right:8px; top:50%; transform:translateY(-50%); font-size:10px; color:var(--cor-texto-secundario); }
  .prz-no.alerta-critico { border-left-color:var(--prz-crit); }
  .prz-no.alerta-vencido { border-left-color:var(--prz-venc); box-shadow:0 0 0 1.5px var(--prz-venc), var(--sombra-card); }
  .prz-no.alerta-semdata { border-left-color:var(--prz-sem); }
  .prz-no:not(.raiz):hover { border-color:var(--cor-acento-prazos); }
  .prz-zoom { position:absolute; right:14px; bottom:14px; display:flex; flex-direction:column; gap:6px; z-index:3; }
  .prz-zoom button { width:34px; height:34px; border-radius:8px; border:1px solid var(--cor-card-elevado);
    background:var(--cor-card); color:var(--cor-texto-primario); font-size:16px; cursor:pointer; box-shadow:var(--sombra-card); }
  .prz-zoom button:hover { border-color:var(--cor-acento-prazos); color:var(--cor-acento-prazos-texto); }

  .prz-selo { display:inline-block; padding:2px 8px; border-radius:999px; color:#fff; background:var(--c);
    font-size:10.5px; font-weight:800; letter-spacing:.2px; }
  .prz-contagem { font-variant-numeric:tabular-nums; font-weight:800; font-size:11.5px; margin-top:3px; color:var(--c); }

  /* ---- painel de detalhes ---- */
  .prz-painel { position:absolute; top:0; right:0; width:380px; max-width:92vw; height:100%; background:var(--cor-card);
    border-left:3px solid var(--cor-acento-prazos); box-shadow:-8px 0 24px rgba(0,0,0,.25); padding:22px 20px;
    overflow-y:auto; z-index:6; transform:translateX(105%); transition:transform .25s; visibility:hidden; }
  .prz-painel.aberto { transform:none; visibility:visible; }
  .prz-painel-fechar { position:absolute; top:12px; right:12px; border:none; background:none; font-size:22px; line-height:1;
    color:var(--cor-texto-secundario); cursor:pointer; padding:2px 6px; }
  .prz-painel-fechar:hover { color:var(--cor-acento-prazos-texto); }
  .prz-painel .trilha { font-size:11.5px; color:var(--cor-texto-secundario); margin:0 26px 8px 0; }
  .prz-painel h3 { font-size:16px; margin:0 0 14px; line-height:1.3; color:var(--cor-texto-primario); }
  .prz-painel .campo { margin-bottom:12px; }
  .prz-painel .campo b { display:block; font-size:11px; text-transform:uppercase; letter-spacing:.4px;
    color:var(--cor-texto-secundario); margin-bottom:3px; }
  .prz-painel .campo span { font-size:13.5px; color:var(--cor-texto-primario); white-space:pre-wrap; }
  .prz-painel .contagem-grande { font-size:22px; font-weight:800; font-variant-numeric:tabular-nums; color:var(--c); }

  /* ---- cadastro no sistema (25/11/2026) ---- */
  #prz-overlay .ra-botoes > .prz-btn-nova { background:var(--cor-acento-prazos); color:#fff; font-weight:800; }
  #prz-overlay .ra-botoes > .prz-btn-nova:hover:not(:disabled) { background:var(--cor-acento-prazos-hover); }
  #prz-overlay .ra-botoes > .prz-btn-nova:disabled { opacity:.45; cursor:not-allowed; }
  .prz-painel-aviso { margin-top:18px; padding:10px 12px; border-radius:8px; font-size:12.5px; line-height:1.4;
    border:1px dashed var(--cor-acento-prazos); color:var(--cor-texto-secundario); }
  .prz-painel-acoes { display:flex; gap:8px; margin-top:18px; padding-top:14px; border-top:1px solid var(--cor-card-elevado); }
  .prz-painel-acoes button { flex:1; font:inherit; font-size:12.5px; font-weight:800; letter-spacing:.3px; cursor:pointer;
    border-radius:8px; padding:9px 10px; border:1px solid var(--cor-acento-prazos); }
  .prz-painel-acoes .editar { background:var(--cor-acento-prazos); color:#fff; }
  .prz-painel-acoes .editar:hover { background:var(--cor-acento-prazos-hover); }
  .prz-painel-acoes .excluir { background:transparent; color:var(--cor-acento-prazos-texto); }
  .prz-painel-acoes .excluir:hover { background:color-mix(in srgb, var(--cor-acento-prazos) 12%, transparent); }
  .prz-form-fundo { display:none; position:fixed; inset:0; z-index:95; background:rgba(5,10,20,.6);
    align-items:center; justify-content:center; padding:20px; }
  .prz-form-fundo.aberto { display:flex; }
  .prz-form { background:var(--cor-card); color:var(--cor-texto-primario); border-radius:12px; width:min(760px, 100%);
    max-height:100%; display:flex; flex-direction:column; box-shadow:0 20px 60px rgba(0,0,0,.45);
    border-top:4px solid var(--cor-acento-prazos); }
  .prz-form-cab { display:flex; align-items:center; justify-content:space-between; padding:16px 20px 10px; }
  .prz-form-cab h3 { margin:0; font-size:18px; font-weight:800; }
  .prz-form-x { border:none; background:none; font-size:26px; line-height:1; color:var(--cor-texto-secundario); cursor:pointer; }
  .prz-form-x:hover { color:var(--cor-acento-prazos-texto); }
  .prz-form-corpo { display:grid; grid-template-columns:1fr 1fr; gap:12px 16px; padding:6px 20px 10px; overflow-y:auto; }
  .prz-form-corpo label { display:flex; flex-direction:column; gap:4px; font-size:11.5px; font-weight:700; letter-spacing:.3px;
    text-transform:uppercase; color:var(--cor-texto-secundario); }
  .prz-form-corpo label.largo { grid-column:1 / -1; }
  .prz-form-corpo label em { color:var(--cor-acento-prazos-texto); font-style:normal; }
  .prz-form-corpo input, .prz-form-corpo select, .prz-form-corpo textarea { font:inherit; font-size:13.5px; font-weight:400;
    text-transform:none; letter-spacing:0; color:var(--cor-texto-primario); background:var(--cor-fundo);
    border:1px solid var(--cor-card-elevado); border-radius:7px; padding:8px 10px; outline:none; resize:vertical; }
  .prz-form-corpo input:focus, .prz-form-corpo select:focus, .prz-form-corpo textarea:focus { border-color:var(--cor-acento-prazos); }
  .prz-form-corpo .invalido { border-color:var(--prz-venc); }
  html[data-tema="escuro"] .prz-form-corpo input[type="date"] { color-scheme:dark; }
  .prz-form-erro { color:var(--prz-venc); font-size:12.5px; font-weight:700; padding:0 20px; min-height:4px; }
  .prz-form-erro:empty { display:none; }
  .prz-form-rodape { display:flex; align-items:center; gap:10px; padding:12px 20px 16px; border-top:1px solid var(--cor-card-elevado); }
  .prz-form-dica { margin-right:auto; font-size:11.5px; color:var(--cor-texto-terciario); }
  .prz-form-rodape button { font:inherit; font-size:13px; font-weight:800; border-radius:8px; padding:9px 18px; cursor:pointer; }
  .prz-form-cancelar { background:transparent; color:var(--cor-texto-primario); border:1px solid var(--cor-card-elevado); }
  .prz-form-cancelar:hover { border-color:var(--cor-acento-prazos); }
  .prz-form-salvar { background:var(--cor-acento-prazos); color:#fff; border:1px solid var(--cor-acento-prazos); }
  .prz-form-salvar:hover { background:var(--cor-acento-prazos-hover); }
  .prz-form-salvar:disabled { opacity:.6; cursor:wait; }
  @media (max-width:700px) { .prz-form-corpo { grid-template-columns:1fr; } }

  .prz-toast { position:fixed; left:50%; bottom:22px; transform:translate(-50%, 20px); opacity:0; pointer-events:none; z-index:90;
    background:var(--cor-acento-prazos); color:#fff; padding:9px 16px; border-radius:999px; font-size:12.5px; font-weight:700;
    box-shadow:var(--sombra-card); transition:opacity .25s, transform .25s; }
  .prz-toast.visivel { opacity:1; transform:translate(-50%, 0); }

  /* ---- grade de filtros da seção (dentro do painel de filtros #app) ---- */
  #grade-prz { flex:1; display:none; grid-template-columns:repeat(4, 1fr); grid-auto-rows:1fr; gap:14px;
    overflow:hidden; min-height:0; }
  #grade-prz .bloco { border-color:color-mix(in srgb, var(--cor-acento-prazos) 25%, transparent); }
  #grade-prz .bloco:hover { border-color:var(--cor-acento-prazos); }
  #grade-prz label.item input[type="checkbox"]:checked { background:var(--cor-acento-prazos); border-color:var(--cor-acento-prazos); }
  #grade-prz .btn-mini:hover { border-color:var(--cor-acento-prazos); color:var(--cor-acento-prazos-texto); }
  #grade-prz .bloco-busca input:focus { border-color:var(--cor-acento-prazos); }
  #grade-prz .prz-item-n { margin-left:auto; padding-left:8px; font-size:11px; color:var(--cor-texto-terciario);
    font-variant-numeric:tabular-nums; }
  #grade-prz label.item { display:flex; }
  .prz-recorte { grid-column:span 2; background:var(--cor-card); border-radius:var(--raio-lg); padding:14px;
    display:flex; flex-direction:column; align-items:center; justify-content:center; gap:4px; text-align:center;
    font:inherit; color:inherit; border:1px solid transparent; cursor:pointer;
    transition:border-color .15s, box-shadow .15s, transform .15s; }
  .prz-recorte:hover, .prz-recorte:focus-visible { border-color:var(--cor-acento-prazos); box-shadow:var(--sombra-card);
    outline:none; transform:translateY(-1px); }
  .prz-recorte-rotulo { font-size:12px; font-weight:700; letter-spacing:.5px; color:var(--cor-texto-secundario); }
  .prz-recorte-num { font-size:34px; font-weight:800; line-height:1.1; color:var(--cor-acento-prazos-texto); }
  .prz-recorte-de { font-size:12px; color:var(--cor-texto-terciario); }
  .prz-recorte-alertas { display:flex; flex-wrap:wrap; justify-content:center; gap:6px 14px; margin-top:6px; font-size:12px;
    color:var(--cor-texto-secundario); }
  .prz-recorte-alertas b { color:var(--c); font-variant-numeric:tabular-nums; }
  .prz-recorte-ir { margin-top:6px; font-size:11.5px; font-weight:700; letter-spacing:.3px; color:var(--cor-acento-prazos-texto); }
  .prz-grade-aviso { grid-column:1 / -1; padding:20px; color:var(--cor-texto-secundario); }
  @media (max-width:1300px) { .prz-destaques { grid-template-columns:repeat(4, minmax(140px,1fr)); } }
  @media (max-width:900px) { #grade-prz { grid-template-columns:repeat(2, 1fr); grid-auto-rows:minmax(260px, auto); overflow:auto; }
    .prz-destaques { grid-template-columns:repeat(2, minmax(120px,1fr)); } .prz-tl-rot { flex-basis:220px; width:220px; }
    .prz-tl-fundo { left:220px; } }
  @media (prefers-reduced-motion:reduce) { .prz-tl-marco.pulsa::after { animation:none; } .prz-painel { transition:none; } }
</style>
<script>
(function () {
  "use strict";
  var overlay = document.getElementById("prz-overlay");
  if (!overlay) return;

  async function api(nome) {
    var args = Array.prototype.slice.call(arguments, 1);
    if (window.pywebview && window.pywebview.api && window.pywebview.api[nome]) {
      return await window.pywebview.api[nome].apply(null, args);
    }
    var r = await fetch("/api/" + nome, {method: "POST",
      headers: {"Content-Type": "application/json"}, body: JSON.stringify(args)});
    return await r.json();
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c];
    });
  }
  function norm(s) {
    return String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  // ---------------- regra de alerta (aba "Resumo e Legenda") ----------------
  var ALERTAS = ["Vencido", "Crítico", "Atenção", "OK", "Sem data", "Concluído"];
  var ALERTA_COR = {"Vencido": "var(--prz-venc)", "Crítico": "var(--prz-crit)", "Atenção": "var(--prz-aten)",
    "OK": "var(--prz-ok)", "Sem data": "var(--prz-sem)", "Concluído": "var(--prz-concl)"};
  var ALERTA_DESC = {"Vencido": "prazo já passou", "Crítico": "vence em até 7 dias", "Atenção": "vence em 8 a 30 dias",
    "OK": "vence em mais de 30 dias", "Sem data": "prazo a definir", "Concluído": "ações concluídas"};
  var DIA = 86400000;
  var MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  function hoje0() { var h = new Date(); h.setHours(0, 0, 0, 0); return h; }
  function dataLocal(iso) { return new Date(iso + "T00:00:00"); }
  function diasAte(d) { return Math.round((dataLocal(d.data_prazo) - hoje0()) / DIA); }
  function fmtData(t) { return t.toLocaleDateString("pt-BR", {day: "2-digit", month: "2-digit", year: "2-digit"}); }
  function concluido(d) { return norm(d.status).indexOf("conclu") === 0; }
  function alertaDe(d) {
    if (concluido(d)) return "Concluído";
    if (!d.data_prazo) return "Sem data";
    var dias = diasAte(d);
    if (dias < 0) return "Vencido";
    if (dias <= 7) return "Crítico";
    if (dias <= 30) return "Atenção";
    return "OK";
  }
  function contagemTxt(d) {
    var a = alertaDe(d);
    if (a === "Concluído") return "concluído";
    if (a === "Sem data") return "sem prazo definido";
    var dias = diasAte(d);
    if (dias < 0) return -dias + (dias === -1 ? " dia" : " dias") + " de atraso";
    if (dias === 0) return "vence hoje";
    return "faltam " + dias + (dias === 1 ? " dia" : " dias");
  }
  function prazoInformadoUtil(d) {
    var p = String(d.prazo_informado || "").trim();
    return p && norm(p) !== "a definir" ? p : "";
  }

  // ---------------- estado ----------------
  var BLOCOS = [
    {chave: "ministerio", titulo: "ÓRGÃO / MINISTÉRIO"},
    {chave: "eixo", titulo: "EIXO"},
    {chave: "projeto", titulo: "PROJETO / FRENTE", busca: "Buscar projeto…"},
    {chave: "responsavel", titulo: "RESPONSÁVEL"},
    {chave: "status", titulo: "STATUS"},
    {chave: "alerta", titulo: "ALERTA DE PRAZO"},
  ];
  var CHAVE_STORAGE = "przFiltrosV1";
  var PRZ = {itens: [], carregado: false, carregando: false, assinatura: null, sel: {}, busca: "",
             visao: "tempo", atualizado: "", aviso: "",
             // cadastro no sistema (25/11/2026): revisão do arquivo de dados,
             // se este usuário edita e as listas de valores do formulário
             revisao: null, podeEditar: false, opcoes: {}};
  BLOCOS.forEach(function (b) { PRZ.sel[b.chave] = new Set(); });
  function valorCampo(d, chave) {
    if (chave === "alerta") return alertaDe(d);
    if (chave === "responsavel") return d.responsavel || "Não informado";
    return d[chave] || "";
  }
  function salvarFiltros() {
    try {
      var o = {busca: PRZ.busca, visao: PRZ.visao};
      BLOCOS.forEach(function (b) { o[b.chave] = Array.from(PRZ.sel[b.chave]); });
      localStorage.setItem(CHAVE_STORAGE, JSON.stringify(o));
    } catch (e) {}
  }
  (function lerFiltros() {
    try {
      var o = JSON.parse(localStorage.getItem(CHAVE_STORAGE) || "null");
      if (!o) return;
      BLOCOS.forEach(function (b) { (o[b.chave] || []).forEach(function (v) { PRZ.sel[b.chave].add(v); }); });
      PRZ.busca = o.busca || "";
      if (o.visao) PRZ.visao = o.visao;
    } catch (e) {}
  })();

  function passaBusca(d) {
    var termos = norm(PRZ.busca).split(/\s+/).filter(Boolean);
    if (!termos.length) return true;
    var alvo = norm([d.id, d.ministerio, d.eixo, d.projeto, d.descricao, d.prazo_informado, d.status,
      d.responsavel, d.observacoes, d.fonte, alertaDe(d)].join(" "));
    return termos.every(function (t) { return alvo.indexOf(t) !== -1; });
  }
  function passa(d, chaveExcluida) {
    for (var i = 0; i < BLOCOS.length; i++) {
      var c = BLOCOS[i].chave;
      if (c === chaveExcluida) continue;
      var s = PRZ.sel[c];
      if (s.size && !s.has(valorCampo(d, c))) return false;
    }
    return passaBusca(d);
  }
  function itensFiltrados() { return PRZ.itens.filter(function (d) { return passa(d); }); }
  function nFiltros() {
    var n = 0;
    BLOCOS.forEach(function (b) { if (PRZ.sel[b.chave].size) n++; });
    if (PRZ.busca.trim()) n++;
    return n;
  }
  function filtrosMudaram() {
    salvarFiltros();
    if (window.agendarContadoresBase) window.agendarContadoresBase();
    renderizarGrade();
    if (overlay.classList.contains("aberto")) renderizarDash();
  }

  // ---------------- carga e atualização ao vivo ----------------
  function carregar(silencioso) {
    if (PRZ.carregando) return Promise.resolve();
    PRZ.carregando = true;
    if (!silencioso) document.getElementById("przMsg").textContent = "Carregando o Controle de Prazos…";
    return api("prazos_dados").then(function (r) {
      PRZ.carregando = false;
      if (!r || !r.ok) {
        document.getElementById("przMsg").textContent = (r && (r.mensagem || r.erro)) || "Não foi possível ler o Controle de Prazos.";
        var g = document.getElementById("grade-prz");
        if (g) g.innerHTML = '<div class="prz-grade-aviso">' + esc(document.getElementById("przMsg").textContent) + "</div>";
        return;
      }
      var antes = PRZ.assinatura;
      PRZ.itens = r.itens || [];
      PRZ.assinatura = r.assinatura;
      PRZ.atualizado = r.atualizado_em || "";
      PRZ.aviso = r.aviso || "";
      PRZ.revisao = r.revisao;
      PRZ.podeEditar = !!r.pode_editar;
      PRZ.opcoes = r.opcoes || {};
      PRZ.editorSemRede = !!r.editor_sem_rede;
      if (PRZ.editorSemRede) PRZ.aviso = (PRZ.aviso ? PRZ.aviso + " · " : "") + "edição indisponível fora da rede";
      // editor fora da rede: o botão aparece DESATIVADO, com o motivo (antes
      // sumia, e parecia defeito — pedido do usuário em 25/11/2026)
      var bNova = document.getElementById("przNova");
      bNova.hidden = !(PRZ.podeEditar || PRZ.editorSemRede);
      bNova.disabled = !PRZ.podeEditar;
      bNova.title = PRZ.podeEditar ? "Cadastrar uma nova ação no Controle de Prazos"
        : "Edição indisponível: esta máquina não acessa a pasta da rede (X:) onde fica o cadastro. Conecte-se à rede e abra de novo.";
      PRZ.carregado = true;
      document.getElementById("przMsg").textContent = "";
      document.getElementById("przConteudo").hidden = false;
      document.getElementById("przAtualizado").textContent = PRZ.atualizado ? " · cadastro atualizado em " + PRZ.atualizado : "";
      document.getElementById("przAviso").textContent = PRZ.aviso;
      if (window.atualizarCabecalhoFiltros) window.atualizarCabecalhoFiltros();
      renderizarGrade();
      renderizarDash();
      if (window.agendarContadoresBase) window.agendarContadoresBase();
      if (silencioso && antes !== null && antes !== PRZ.assinatura && !PRZ.gravouAgora) {
        avisar("Controle de Prazos atualizado às " + new Date().toLocaleTimeString("pt-BR", {hour: "2-digit", minute: "2-digit"}) +
          " — " + PRZ.itens.length + " ações");
      }
      PRZ.gravouAgora = false;
      // painel aberto: mostra a versão nova da mesma ação (ou fecha, se saiu)
      if (painel.classList.contains("aberto") && painel._id) {
        var atual = PRZ.itens.find(function (x) { return String(x.id) === painel._id; });
        if (atual) abrirPainel(atual); else fecharPainel();
      }
    }).catch(function (e) {
      PRZ.carregando = false;
      document.getElementById("przMsg").textContent = "Erro ao ler o Controle de Prazos: " + e;
    });
  }
  // Vigia: a cada 8 s (só com a seção PRAZOS ativa) confere a revisão do
  // cadastro e, se outra máquina gravou, relê sem perder aba, busca e filtros.
  setInterval(function () {
    if (window.baseAtiva !== "PRZ" || !PRZ.carregado || PRZ.carregando) return;
    api("prazos_assinatura").then(function (r) {
      if (r && r.ok && r.assinatura && r.assinatura !== PRZ.assinatura) setTimeout(function () { carregar(true); }, 1500);
    }).catch(function () {});
  }, 8000);
  var toastTimer = null;
  function avisar(msg) {
    var t = document.getElementById("przToast");
    t.textContent = msg;
    t.classList.add("visivel");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove("visivel"); }, 4500);
  }

  // ---------------- grade de filtros (painel de filtros #app) ----------------
  function garantirGrade() {
    var g = document.getElementById("grade-prz");
    if (g) return g;
    var app = document.getElementById("app");
    if (!app) return null;
    g = document.createElement("div");
    g.id = "grade-prz";
    g.innerHTML = '<div class="prz-grade-aviso">Carregando o Controle de Prazos…</div>';
    app.appendChild(g);
    return g;
  }
  garantirGrade();
  var buscaBloco = {};
  function renderizarGrade() {
    var g = garantirGrade();
    if (!g || !PRZ.carregado) return;
    var rolagens = {};
    g.querySelectorAll(".bloco").forEach(function (b) {
      var l = b.querySelector(".bloco-lista");
      if (l) rolagens[b.dataset.chave] = l.scrollTop;
    });
    g.innerHTML = "";
    BLOCOS.forEach(function (b) {
      var contagem = {};
      PRZ.itens.forEach(function (d) {
        if (!passa(d, b.chave)) return;
        var v = valorCampo(d, b.chave);
        contagem[v] = (contagem[v] || 0) + 1;
      });
      var opcoes;
      if (b.chave === "alerta") opcoes = ALERTAS.slice();
      else {
        opcoes = Array.from(new Set(PRZ.itens.map(function (d) { return valorCampo(d, b.chave); })));
        opcoes.sort(function (x, y) { return x.localeCompare(y, "pt-BR"); });
      }
      var el = document.createElement("div");
      el.className = "bloco";
      el.dataset.chave = b.chave;
      var t = document.createElement("div");
      t.className = "bloco-titulo";
      t.innerHTML = "<span>" + esc(b.titulo) + "</span>";
      el.appendChild(t);
      var botoes = document.createElement("div");
      botoes.className = "bloco-botoes";
      [["Marcar tudo", function () {
        el.querySelectorAll("label.item").forEach(function (l) {
          if (l.style.display !== "none" && !l.classList.contains("indisponivel")) PRZ.sel[b.chave].add(l.dataset.valor);
        });
        filtrosMudaram();
      }], ["Limpar", function () { PRZ.sel[b.chave].clear(); filtrosMudaram(); }]].forEach(function (par) {
        var bt = document.createElement("button");
        bt.type = "button";
        bt.className = "btn-mini";
        bt.textContent = par[0];
        bt.onclick = par[1];
        botoes.appendChild(bt);
      });
      el.appendChild(botoes);
      var lista = document.createElement("div");
      lista.className = "bloco-lista";
      if (b.busca) {
        var caixa = document.createElement("div");
        caixa.className = "bloco-busca";
        var campo = document.createElement("input");
        campo.type = "search";
        campo.placeholder = b.busca;
        campo.value = buscaBloco[b.chave] || "";
        campo.oninput = function () {
          buscaBloco[b.chave] = campo.value;
          var termo = norm(campo.value);
          lista.querySelectorAll("label.item").forEach(function (l) {
            l.style.display = !termo || norm(l.dataset.valor).indexOf(termo) !== -1 ? "" : "none";
          });
        };
        caixa.appendChild(campo);
        el.appendChild(caixa);
      }
      opcoes.forEach(function (op) {
        var marcado = PRZ.sel[b.chave].has(op);
        var label = document.createElement("label");
        label.className = "item" + (!contagem[op] && !marcado ? " indisponivel" : "");
        label.dataset.valor = op;
        var input = document.createElement("input");
        input.type = "checkbox";
        input.checked = marcado;
        input.onchange = function () {
          if (input.checked) PRZ.sel[b.chave].add(op); else PRZ.sel[b.chave].delete(op);
          filtrosMudaram();
        };
        var span = document.createElement("span");
        span.textContent = op;
        if (b.chave === "alerta") span.style.color = "";
        var n = document.createElement("span");
        n.className = "prz-item-n";
        n.textContent = contagem[op] || 0;
        label.appendChild(input);
        label.appendChild(span);
        label.appendChild(n);
        lista.appendChild(label);
      });
      el.appendChild(lista);
      g.appendChild(el);
      if (b.busca && buscaBloco[b.chave]) el.querySelector(".bloco-busca input").oninput();
      if (rolagens[b.chave]) lista.scrollTop = rolagens[b.chave];
    });
    // Ações no recorte (abre o dash) — mesmo papel do AÇÕES NO RECORTE do MONITORAMENTO
    var itens = itensFiltrados();
    var cont = {};
    ALERTAS.forEach(function (a) { cont[a] = 0; });
    itens.forEach(function (d) { cont[alertaDe(d)]++; });
    var rec = document.createElement("button");
    rec.type = "button";
    rec.className = "prz-recorte";
    rec.title = "Ver o dash do Controle de Prazos com este recorte";
    rec.innerHTML = '<span class="prz-recorte-rotulo">AÇÕES NO RECORTE</span>' +
      '<span class="prz-recorte-num">' + itens.length.toLocaleString("pt-BR") + "</span>" +
      '<span class="prz-recorte-de">de ' + PRZ.itens.length.toLocaleString("pt-BR") + " ações cadastradas</span>" +
      '<span class="prz-recorte-alertas">' + ALERTAS.filter(function (a) { return cont[a]; }).map(function (a) {
        return '<span style="--c:' + ALERTA_COR[a] + '"><b>' + cont[a] + "</b> " + esc(a.toLowerCase()) + "</span>";
      }).join("") + "</span>" +
      '<span class="prz-recorte-ir">Ver o dash ›</span>';
    rec.onclick = function () { if (window.trocarBase) window.trocarBase("PRZ", {tela: "dash"}); };
    g.appendChild(rec);
  }

  // ---------------- dash ----------------
  function renderizarDash() {
    if (!PRZ.carregado) return;
    renderizarDestaques();
    renderizarRadar();
    var caixa = document.getElementById("przBuscaGeral");
    if (document.activeElement !== caixa) caixa.value = PRZ.busca;
    caixa.parentNode.classList.toggle("ativo", !!PRZ.busca.trim());
    trocarVisao(PRZ.visao, true);
  }
  function renderizarDestaques() {
    var itens = itensFiltrados();
    var cont = {};
    ALERTAS.forEach(function (a) { cont[a] = 0; });
    // os cartões de alerta contam o recorte SEM o próprio filtro de alerta
    // (assim dá pra ligar/desligar cada um vendo quantas ações traz)
    PRZ.itens.forEach(function (d) { if (passa(d, "alerta")) cont[alertaDe(d)]++; });
    var selA = PRZ.sel.alerta;
    var html = '<div class="prz-kpi prz-kpi-total" style="--c:var(--cor-acento-prazos)"><span>Ações no recorte</span><b>' +
      itens.length.toLocaleString("pt-BR") + "</b><small>de " + PRZ.itens.length.toLocaleString("pt-BR") +
      " ações cadastradas" + (nFiltros() ? " · " + nFiltros() + " filtro(s) ativo(s)" : "") + "</small></div>";
    html += ALERTAS.map(function (a) {
      var classe = "prz-kpi" + (selA.has(a) ? " ativo" : (selA.size ? " apagado" : ""));
      return '<button type="button" class="' + classe + '" data-alerta="' + esc(a) + '" style="--c:' + ALERTA_COR[a] +
        '" aria-pressed="' + (selA.has(a) ? "true" : "false") + '" title="Clique para filtrar por ' + esc(a) +
        ' (clique de novo para tirar o filtro)"><span>' + esc(a) + "</span><b>" + cont[a] + "</b><small>" +
        esc(ALERTA_DESC[a]) + "</small></button>";
    }).join("");
    var el = document.getElementById("przDestaques");
    el.innerHTML = html;
    el.querySelectorAll("button[data-alerta]").forEach(function (b) {
      b.onclick = function () {
        var a = b.dataset.alerta;
        if (selA.has(a)) selA.delete(a); else selA.add(a);
        filtrosMudaram();
      };
    });
  }
  function renderizarRadar() {
    var rel = itensFiltrados().filter(function (d) { return ["Vencido", "Crítico", "Atenção"].indexOf(alertaDe(d)) !== -1; });
    rel.sort(function (a, b) { return a.data_prazo.localeCompare(b.data_prazo); });
    var el = document.getElementById("przRadar");
    el.innerHTML = rel.slice(0, 8).map(function (d) {
      var a = alertaDe(d);
      return '<button type="button" class="prz-radar-card" data-id="' + esc(d.id) + '" style="--c:' + ALERTA_COR[a] + '">' +
        '<span class="p">' + esc(d.projeto) + '</span><span class="o">' + esc(d.descricao || d.ministerio) + "</span>" +
        '<span class="c">' + esc(contagemTxt(d)) + " · " + fmtData(dataLocal(d.data_prazo)) + "</span></button>";
    }).join("") || '<span class="prz-radar-vazio">Nenhum prazo vencido ou nos próximos 30 dias no recorte.</span>';
    ligarCliques(el, ".prz-radar-card");
  }
  function ligarCliques(container, seletor) {
    container.querySelectorAll(seletor).forEach(function (el) {
      var abrir = function () {
        var d = PRZ.itens.find(function (x) { return String(x.id) === el.dataset.id; });
        if (d) abrirPainel(d);
      };
      el.addEventListener("click", abrir);
      el.addEventListener("keydown", function (e) {
        if ((e.key === "Enter" || e.key === " ") && el.tagName !== "BUTTON") { e.preventDefault(); abrir(); }
      });
    });
  }

  var VISOES = {tempo: "przVisTempo", quadro: "przVisQuadro", mapa: "przVisMapa"};
  var DICAS = {tempo: "Prazos na régua de datas · faixa vermelha = vence em até 7 dias · clique numa ação para ver os detalhes",
    quadro: "Ações agrupadas pelo alerta de prazo · clique num cartão para ver os detalhes",
    mapa: "Órgão › Eixo › Projeto › ação · clique nos ramos para abrir/fechar · arraste para mover, roda do mouse para zoom"};
  function trocarVisao(v, semSalvar) {
    if (!VISOES[v]) v = "tempo";
    PRZ.visao = v;
    if (!semSalvar) salvarFiltros();
    overlay.querySelectorAll(".prz-aba").forEach(function (b) {
      var at = b.dataset.visao === v;
      b.classList.toggle("ativa", at);
      b.setAttribute("aria-selected", at ? "true" : "false");
    });
    Object.keys(VISOES).forEach(function (k) { document.getElementById(VISOES[k]).classList.toggle("ativa", k === v); });
    document.getElementById("przAbasDica").textContent = DICAS[v];
    document.getElementById("przMapaTudo").hidden = v !== "mapa";
    if (v === "tempo") desenharTempo();
    else if (v === "quadro") desenharQuadro();
    else desenharMapa();
  }
  overlay.querySelectorAll(".prz-aba").forEach(function (b) {
    b.addEventListener("click", function () { trocarVisao(b.dataset.visao); });
  });

  // ---------------- linha do tempo ----------------
  var LARG_ROT = 340, PX_POR_DIA = 6, tempoJaRolado = false;
  function desenharTempo() {
    var el = document.getElementById("przVisTempo");
    var itens = itensFiltrados();
    if (!itens.length) { el.innerHTML = '<div class="prz-vazio">Nenhuma ação corresponde aos filtros.</div>'; return; }
    var larguraRot = window.innerWidth <= 900 ? 220 : LARG_ROT;
    var hoje = hoje0();
    var comData = itens.filter(function (d) { return d.data_prazo; })
      .sort(function (a, b) { return a.data_prazo.localeCompare(b.data_prazo); });
    var semData = itens.filter(function (d) { return !d.data_prazo; });
    var ini = new Date(+hoje - 21 * DIA), fim = new Date(+hoje + 45 * DIA);
    comData.forEach(function (d) {
      var t = dataLocal(d.data_prazo), antes = new Date(+t - 7 * DIA);
      if (antes < ini) ini = antes;
      if (t > fim) fim = t;
    });
    ini = new Date(ini.getFullYear(), ini.getMonth(), 1);
    fim = new Date(fim.getFullYear(), fim.getMonth() + 1, 1);
    var total = fim - ini;
    function pct(t) { return ((t - ini) / total * 100).toFixed(3) + "%"; }
    var pHoje = pct(hoje), p7 = pct(new Date(+hoje + 7 * DIA)), p30 = pct(new Date(+hoje + 30 * DIA));
    var larguraMin = larguraRot + Math.round(total / DIA * PX_POR_DIA);
    var meses = [], grades = [];
    for (var m = new Date(ini); m < fim; m = new Date(m.getFullYear(), m.getMonth() + 1, 1)) {
      var prox = new Date(m.getFullYear(), m.getMonth() + 1, 1);
      meses.push('<div class="prz-tl-mes" style="left:' + pct(m) + ";width:" + ((prox - m) / total * 100).toFixed(3) + '%">' +
        MESES[m.getMonth()] + "/" + String(m.getFullYear()).slice(2) + "</div>");
      grades.push('<div class="prz-tl-grade" style="left:' + pct(m) + '"></div>');
      grades.push('<div class="prz-tl-grade meio" style="left:' + pct(new Date(m.getFullYear(), m.getMonth(), 15)) + '"></div>');
    }
    var cont = {};
    ALERTAS.forEach(function (a) { cont[a] = 0; });
    itens.forEach(function (d) { cont[alertaDe(d)]++; });
    var cab = '<div class="prz-tl-linha prz-tl-cab"><div class="prz-tl-rot">' +
      '<div class="prz-tl-cab-tit">Ações e prazos</div>' +
      '<div class="prz-tl-cab-res"><b style="color:var(--prz-venc)">' + cont["Vencido"] + "</b> vencidas · " +
      '<b style="color:var(--prz-crit)">' + cont["Crítico"] + "</b> em até 7 dias · " +
      '<b style="color:var(--prz-aten)">' + cont["Atenção"] + "</b> em até 30 dias</div></div>" +
      '<div class="prz-tl-trilha">' + meses.join("") +
      '<div class="prz-tl-zfaixa critica" style="left:' + pHoje + ";width:calc(" + p7 + " - " + pHoje + ')" title="vence em até 7 dias">7 d</div>' +
      '<div class="prz-tl-zfaixa atencao" style="left:' + p7 + ";width:calc(" + p30 + " - " + p7 + ')" title="vence em até 30 dias">até 30 dias</div>' +
      '<div class="prz-tl-zrot-hoje" style="left:' + pHoje + '">HOJE ' + fmtData(hoje).slice(0, 5) + "</div></div></div>";
    var fundo = '<div class="prz-tl-fundo" style="left:' + larguraRot + 'px">' +
      '<div class="prz-tl-zona passado" style="left:0;width:' + pHoje + '"></div>' +
      '<div class="prz-tl-zona critica" style="left:' + pHoje + ";width:calc(" + p7 + " - " + pHoje + ')"></div>' +
      '<div class="prz-tl-zona atencao" style="left:' + p7 + ";width:calc(" + p30 + " - " + p7 + ')"></div>' +
      grades.join("") + '<div class="prz-tl-hoje" style="left:' + pHoje + '"></div></div>';
    function linhaItem(d) {
      var a = alertaDe(d), cor = ALERTA_COR[a], trilha = "", sub = esc(d.projeto), extra = "";
      if (d.data_prazo) {
        var t = dataLocal(d.data_prazo), dias = diasAte(d), pT = pct(t), txt, pTxt;
        sub += " · " + fmtData(t);
        if (a === "Concluído") { txt = "✓ concluído"; pTxt = "calc(" + pT + " + 12px)"; }
        else if (dias < 0) {
          trilha += '<div class="prz-tl-barra atraso" style="left:' + pT + ";width:calc(" + pHoje + " - " + pT + ')"></div>';
          txt = -dias + (dias === -1 ? " dia" : " dias") + " de atraso"; pTxt = "calc(" + pHoje + " + 10px)";
        } else {
          trilha += '<div class="prz-tl-barra futura" style="left:' + pHoje + ";width:calc(" + pT + " - " + pHoje + ");--c:" + cor + '"></div>';
          txt = dias === 0 ? "vence hoje" : "faltam " + dias + (dias === 1 ? " dia" : " dias"); pTxt = "calc(" + pT + " + 12px)";
        }
        var marco = a === "Concluído" ? "feito" : (a === "Vencido" || a === "Crítico") ? "pulsa" : "";
        trilha += '<div class="prz-tl-marco ' + marco + '" style="left:' + pT + ";--c:" + cor + '" title="prazo: ' + fmtData(t) + '"></div>';
        trilha += '<div class="prz-tl-txt" style="left:' + pTxt + ";--c:" + cor + '">' + txt + "</div>";
      } else {
        var inf = prazoInformadoUtil(d) ? " · informado: " + esc(prazoInformadoUtil(d)) : "";
        trilha = '<div class="prz-tl-semdata">' + (a === "Concluído" ? "✓ concluído (sem data)" : "prazo a definir") + inf + "</div>";
        sub = esc(d.ministerio) + " · " + sub;
      }
      if (a === "Concluído") extra = " feito";
      return '<div class="prz-tl-linha item' + extra + '" data-id="' + esc(d.id) + '" tabindex="0" role="button">' +
        '<div class="prz-tl-rot" style="border-left-color:' + cor + '"><div class="prz-tl-desc">' + esc(d.descricao || d.projeto) +
        '</div><div class="prz-tl-sub">' + sub + "</div></div>" +
        '<div class="prz-tl-trilha">' + trilha + "</div></div>";
    }
    var grupos = new Map();
    comData.forEach(function (d) {
      if (!grupos.has(d.ministerio)) grupos.set(d.ministerio, []);
      grupos.get(d.ministerio).push(d);
    });
    var ordem = Array.from(new Set(PRZ.itens.map(function (d) { return d.ministerio; })));
    var corpo = "";
    ordem.filter(function (o) { return grupos.has(o); }).forEach(function (o) {
      var lista = grupos.get(o);
      var venc = lista.filter(function (d) { return alertaDe(d) === "Vencido"; }).length;
      var prox = lista.filter(function (d) { return ["Crítico", "Atenção"].indexOf(alertaDe(d)) !== -1; }).length;
      var res = [venc ? '<span style="color:var(--prz-venc)">' + venc + " vencida(s)</span>" : "",
        prox ? '<span style="color:var(--prz-crit)">' + prox + " próxima(s)</span>" : ""].filter(Boolean).join(" · ");
      corpo += '<div class="prz-tl-linha prz-tl-grupo"><div class="prz-tl-rot">' + esc(o) +
        '<span class="resumo">' + lista.length + " ação(ões)" + (res ? " · " + res : "") + "</span></div><div class=\"prz-tl-trilha\"></div></div>";
      corpo += lista.map(linhaItem).join("");
    });
    if (semData.length) {
      corpo += '<div class="prz-tl-linha prz-tl-grupo"><div class="prz-tl-rot">Sem prazo definido' +
        '<span class="resumo">' + semData.length + " ação(ões)</span></div><div class=\"prz-tl-trilha\"></div></div>";
      corpo += semData.map(linhaItem).join("");
    }
    var rolagem = el.scrollTop;
    el.innerHTML = '<div class="prz-tl" style="min-width:' + larguraMin + 'px;--larg-rot:' + larguraRot + 'px">' + fundo + cab + corpo + "</div>";
    el.querySelectorAll(".prz-tl-rot").forEach(function (r) { r.style.flexBasis = larguraRot + "px"; r.style.width = larguraRot + "px"; });
    el.scrollTop = rolagem;
    ligarCliques(el, ".prz-tl-linha.item");
    if (!tempoJaRolado && el.clientWidth) {
      tempoJaRolado = true;
      var largTrilha = el.querySelector(".prz-tl").offsetWidth - larguraRot;
      var xHoje = (hoje - ini) / total * largTrilha, vis = el.clientWidth - larguraRot;
      if (xHoje > vis * 0.6) el.scrollLeft = xHoje - vis * 0.6;
    }
  }

  // ---------------- quadro de alertas ----------------
  function desenharQuadro() {
    var el = document.getElementById("przVisQuadro");
    var itens = itensFiltrados();
    el.innerHTML = '<div class="prz-qd">' + ALERTAS.map(function (a) {
      var lista = itens.filter(function (d) { return alertaDe(d) === a; })
        .sort(function (x, y) { return (x.data_prazo || "9999").localeCompare(y.data_prazo || "9999"); });
      var cards = lista.map(function (d) {
        var num, rot, medidor = "";
        if (a === "Concluído") { num = "✓"; rot = d.data_prazo ? "em " + fmtData(dataLocal(d.data_prazo)) : "concluído"; }
        else if (a === "Sem data") { num = "—"; rot = prazoInformadoUtil(d) ? "informado: " + esc(prazoInformadoUtil(d)) : "prazo a definir"; }
        else {
          var dias = diasAte(d);
          if (dias < 0) { num = -dias; rot = (dias === -1 ? "dia" : "dias") + " de atraso"; }
          else if (dias === 0) { num = "0"; rot = "vence hoje"; }
          else { num = dias; rot = dias === 1 ? "dia restante" : "dias restantes"; }
          if (dias >= 0) {
            var f = Math.max(4, Math.min(100, (60 - dias) / 60 * 100));
            medidor = '<div class="prz-qd-medidor" title="proximidade do prazo"><i style="width:' + f.toFixed(0) + '%"></i></div>';
          }
        }
        var data = d.data_prazo && a !== "Concluído" ? " · prazo " + fmtData(dataLocal(d.data_prazo)) : "";
        return '<div class="prz-qd-card" data-id="' + esc(d.id) + '" tabindex="0" role="button">' +
          '<div class="prz-qd-num"><b>' + num + "</b><span>" + rot + "</span></div>" +
          '<div class="prz-qd-desc">' + esc(d.descricao || d.projeto) + "</div>" +
          '<div class="prz-qd-meta">' + esc(d.projeto) + " · " + esc(d.ministerio) + data + "</div>" + medidor + "</div>";
      }).join("") || '<div class="prz-qd-vazio">nenhuma ação</div>';
      return '<div class="prz-qd-col" style="--c:' + ALERTA_COR[a] + '"><div class="prz-qd-cab"><div class="t">' + esc(a) +
        '<span class="n">' + lista.length + '</span></div><div class="d">' + esc(ALERTA_DESC[a]) + "</div></div>" +
        '<div class="prz-qd-cards">' + cards + "</div></div>";
    }).join("") + "</div>";
    ligarCliques(el, ".prz-qd-card");
  }

  // ---------------- mapa mental ----------------
  var COL_W = 320, ROW_H = 96, ROW_GAP = 14;
  var LARGURA_NO = {raiz: 260, ministerio: 220, eixo: 200, projeto: 230, item: 250};
  var colapsado = new Set(), mapaIniciado = false;
  var zoom = 1, panX = 40, panY = 30;
  var elCanvas = document.getElementById("przMapaCanvas"), elSvg = document.getElementById("przMapaSvg");
  var viewport = document.getElementById("przVisMapa");
  function arvore() {
    var raiz = {nome: "Controle de Prazos", tipo: "raiz", filhos: [], _id: "r"};
    var idx = {};
    itensFiltrados().forEach(function (d) {
      var caminho = [["ministerio", d.ministerio], ["eixo", d.eixo], ["projeto", d.projeto]];
      var no = raiz, chave = "r";
      caminho.forEach(function (p) {
        chave += "|" + p[1];
        if (!idx[chave]) { idx[chave] = {nome: p[1], tipo: p[0], filhos: [], _id: chave}; no.filhos.push(idx[chave]); }
        no = idx[chave];
      });
      no.filhos.push({nome: d.descricao || d.projeto, tipo: "item", dados: d, filhos: [], _id: chave + "|#" + d.id});
    });
    return raiz;
  }
  function piorAlerta(no) {
    if (no.tipo === "item") { no._pior = alertaDe(no.dados); return [no._pior]; }
    var fs = [];
    no.filhos.forEach(function (f) { fs = fs.concat(piorAlerta(f)); });
    no._pior = fs.indexOf("Vencido") !== -1 ? "Vencido" : fs.indexOf("Crítico") !== -1 ? "Crítico"
      : (fs.length && fs.every(function (a) { return a === "Sem data"; })) ? "Sem data" : null;
    return fs;
  }
  function aberto(no) { return no.tipo === "raiz" || (!colapsado.has(no._id) && no.filhos.length > 0); }
  function conteudoNo(no) {
    if (no.tipo === "raiz") return '<div class="tit">CONTROLE DE PRAZOS</div><div class="sub">' + itensFiltrados().length +
      " ações · clique nos ramos para abrir/fechar</div>";
    if (no.tipo === "item") {
      var d = no.dados, a = alertaDe(d);
      return '<div class="desc">' + esc(d.descricao || d.projeto) + '</div><div class="meta">' + esc(d.projeto) +
        (d.responsavel ? " · " + esc(d.responsavel) : "") + '</div><span class="prz-selo" style="--c:' + ALERTA_COR[a] + '">' +
        esc(a) + '</span><div class="prz-contagem" style="--c:' + ALERTA_COR[a] + '">' + esc(contagemTxt(d)) + "</div>";
    }
    var rot = no.filhos.length + (no.tipo === "ministerio" ? " eixo(s)" : no.tipo === "eixo" ? " frente(s)" : " ação(ões)");
    return "<div>" + esc(no.nome) + '</div><div class="meta">' + rot + "</div>";
  }
  function divNo(no) {
    var classe = "prz-no " + no.tipo;
    if (no._pior === "Vencido") classe += " alerta-vencido";
    else if (no._pior === "Crítico") classe += " alerta-critico";
    else if (no._pior === "Sem data") classe += " alerta-semdata";
    var seta = no.tipo !== "raiz" && no.tipo !== "item" ? '<span class="seta">' + (aberto(no) ? "▾" : "▸") + "</span>" : "";
    return '<div class="' + classe + '" style="left:' + no._x + "px;top:" + (no._y || 0) + 'px" data-id="' + esc(no._id) + '">' +
      conteudoNo(no) + seta + "</div>";
  }
  var raizAtual = null;
  function desenharMapa(enquadrar) {
    var raiz = arvore();
    raizAtual = raiz;
    if (!mapaIniciado) {
      mapaIniciado = true;
      (function rec(no) { if (no.tipo !== "raiz" && no.tipo !== "item") colapsado.add(no._id); no.filhos.forEach(rec); })(raiz);
    }
    if (!raiz.filhos.length) {
      elCanvas.innerHTML = '<div class="prz-vazio" style="position:absolute;left:40px;top:40px">Nenhuma ação corresponde aos filtros.</div>';
      elSvg.innerHTML = "";
      return;
    }
    piorAlerta(raiz);
    (function defX(no, p) { no._x = p * COL_W; if (aberto(no)) no.filhos.forEach(function (f) { defX(f, p + 1); }); })(raiz, 0);
    var folhas = [];
    (function col(no) { if (aberto(no)) no.filhos.forEach(col); else folhas.push(no); })(raiz);
    elCanvas.style.visibility = "hidden";
    elCanvas.innerHTML = folhas.map(divNo).join("");
    var cursor = 0;
    var alturas = {};
    elCanvas.querySelectorAll(".prz-no").forEach(function (el) { alturas[el.dataset.id] = el.offsetHeight; });
    folhas.forEach(function (no) { no._y = cursor; cursor += Math.max(alturas[no._id] || ROW_H, ROW_H) + ROW_GAP; });
    (function centro(no) {
      if (aberto(no)) {
        no.filhos.forEach(centro);
        var ys = no.filhos.map(function (f) { return f._y; });
        no._y = (Math.min.apply(null, ys) + Math.max.apply(null, ys)) / 2;
      }
    })(raiz);
    var html = [], linhas = [], maxX = 0, maxY = 0;
    (function montar(no) {
      html.push(divNo(no));
      maxX = Math.max(maxX, no._x + 290); maxY = Math.max(maxY, no._y + 100);
      if (aberto(no)) no.filhos.forEach(function (f) {
        var x1 = no._x + LARGURA_NO[no.tipo], y1 = no._y + 26, x2 = f._x, y2 = f._y + 26, mx = (x1 + x2) / 2;
        linhas.push('<path d="M' + x1 + "," + y1 + " C" + mx + "," + y1 + " " + mx + "," + y2 + " " + x2 + "," + y2 +
          '" fill="none" stroke-width="2"/>');
        montar(f);
      });
    })(raiz);
    elCanvas.innerHTML = html.join("");
    elCanvas.style.visibility = "visible";
    elSvg.setAttribute("width", maxX); elSvg.setAttribute("height", maxY);
    elSvg.innerHTML = linhas.join("");
    viewport._tamanho = [maxX, maxY];
    if (enquadrar) enquadrarMapa(); else aplicarTransform();
    elCanvas.querySelectorAll(".prz-no").forEach(function (el) {
      el.addEventListener("click", function (ev) {
        ev.stopPropagation();
        if (arrastou) return;
        var no = acharNo(raizAtual, el.dataset.id);
        if (!no) return;
        if (no.tipo === "item") abrirPainel(no.dados);
        else if (no.tipo !== "raiz" && no.filhos.length) {
          if (colapsado.has(no._id)) colapsado.delete(no._id); else colapsado.add(no._id);
          desenharMapa();
        }
      });
    });
  }
  function acharNo(no, id) {
    if (no._id === id) return no;
    for (var i = 0; i < no.filhos.length; i++) { var r = acharNo(no.filhos[i], id); if (r) return r; }
    return null;
  }
  function aplicarTransform() {
    var t = "translate(" + panX + "px," + panY + "px) scale(" + zoom + ")";
    elCanvas.style.transform = t; elSvg.style.transform = t;
  }
  function enquadrarMapa() {
    var tam = viewport._tamanho || [800, 600], vw = viewport.clientWidth, vh = viewport.clientHeight, mg = 30;
    if (!vw || !vh) return;
    zoom = Math.min(1.4, Math.max(0.25, Math.min((vw - mg * 2) / tam[0], (vh - mg * 2) / tam[1])));
    panX = Math.max(mg, (vw - tam[0] * zoom) / 2); panY = Math.max(mg, (vh - tam[1] * zoom) / 2);
    aplicarTransform();
  }
  var arrastando = false, ax = 0, ay = 0, arrastou = false;
  viewport.addEventListener("mousedown", function (e) {
    if (e.target.closest(".prz-zoom")) return;
    arrastando = true; arrastou = false; ax = e.clientX; ay = e.clientY; viewport.classList.add("arrastando");
  });
  window.addEventListener("mousemove", function (e) {
    if (!arrastando) return;
    var dx = e.clientX - ax, dy = e.clientY - ay;
    if (Math.abs(dx) + Math.abs(dy) > 3) arrastou = true;
    panX += dx; panY += dy; ax = e.clientX; ay = e.clientY;
    aplicarTransform();
  });
  window.addEventListener("mouseup", function () {
    if (!arrastando) return;
    arrastando = false; viewport.classList.remove("arrastando");
    setTimeout(function () { arrastou = false; }, 0);
  });
  viewport.addEventListener("wheel", function (e) {
    e.preventDefault();
    var novo = Math.min(2.2, Math.max(0.25, zoom * (e.deltaY < 0 ? 1.08 : 0.93)));
    if (novo === zoom) return;
    var r = viewport.getBoundingClientRect(), mx = e.clientX - r.left, my = e.clientY - r.top, k = novo / zoom;
    panX = mx - (mx - panX) * k; panY = my - (my - panY) * k; zoom = novo;
    aplicarTransform();
  }, {passive: false});
  document.getElementById("przZoomMais").onclick = function () { zoom = Math.min(2.2, zoom * 1.15); aplicarTransform(); };
  document.getElementById("przZoomMenos").onclick = function () { zoom = Math.max(0.25, zoom * 0.87); aplicarTransform(); };
  document.getElementById("przZoomCaber").onclick = enquadrarMapa;
  document.getElementById("przMapaTudo").onclick = function () {
    var raiz = arvore(), ramos = [];
    (function rec(no) { if (no.tipo !== "raiz" && no.tipo !== "item") ramos.push(no._id); no.filhos.forEach(rec); })(raiz);
    var algumFechado = ramos.some(function (id) { return colapsado.has(id); });
    colapsado.clear();
    if (!algumFechado) ramos.forEach(function (id) { colapsado.add(id); });
    desenharMapa(true);
  };

  // ---------------- painel de detalhes ----------------
  var painel = document.getElementById("przPainel");
  function abrirPainel(d) {
    var a = alertaDe(d), cor = ALERTA_COR[a];
    function campo(rot, val) { return '<div class="campo"><b>' + rot + "</b><span>" + val + "</span></div>"; }
    document.getElementById("przPainelConteudo").innerHTML =
      '<div class="trilha">' + esc(d.ministerio) + " › " + esc(d.eixo) + " › " + esc(d.projeto) + "</div>" +
      "<h3>" + esc(d.descricao || d.projeto) + "</h3>" +
      '<div class="campo"><b>Alerta</b><span class="prz-selo" style="--c:' + cor + '">' + esc(a) + "</span></div>" +
      '<div class="campo" style="--c:' + cor + '"><b>Contagem</b><span class="contagem-grande">' + esc(contagemTxt(d)) + "</span></div>" +
      campo("Status", esc(d.status)) +
      campo("Prazo informado (original)", esc(d.prazo_informado || "a definir")) +
      campo("Data de referência", d.data_prazo ? dataLocal(d.data_prazo).toLocaleDateString("pt-BR") : "não definida") +
      campo("Responsável", esc(d.responsavel || "não atribuído")) +
      campo("Observações", esc(d.observacoes || "nenhuma")) +
      (d.fonte ? campo("Fonte", esc(d.fonte)) : "") +
      (d.ultima_atualizacao ? campo("Última atualização", esc(d.ultima_atualizacao)) : "") +
      campo("ID", esc(d.id)) +
      (PRZ.podeEditar ? '<div class="prz-painel-acoes"><button type="button" class="editar">EDITAR</button>' +
        '<button type="button" class="excluir">EXCLUIR</button></div>'
        : PRZ.editorSemRede ? '<div class="prz-painel-aviso">Edição indisponível: esta máquina não acessa a pasta da rede ' +
          '(X:) onde fica o cadastro. Conecte-se à rede e abra o Controle de Prazos de novo.</div>' : "");
    painel._id = String(d.id);
    var acoes = document.querySelector("#przPainelConteudo .prz-painel-acoes");
    if (acoes) {
      acoes.querySelector(".editar").onclick = function () { abrirForm(d); };
      acoes.querySelector(".excluir").onclick = function () { excluir(d); };
    }
    painel.classList.add("aberto");
    painel.setAttribute("aria-hidden", "false");
  }
  function fecharPainel() { painel.classList.remove("aberto"); painel.setAttribute("aria-hidden", "true"); painel._id = null; }

  // ---------------- cadastro no sistema (25/11/2026) ----------------
  // Só aparece para quem edita (PRZ.podeEditar, decidido no Python); o
  // Python confere a permissão de novo a cada gravação.
  var formFundo = document.getElementById("przFormFundo"), form = document.getElementById("przForm");
  var formErro = document.getElementById("przFormErro"), formSalvar = document.getElementById("przFormSalvar");
  var formEditando = null;
  function preencherListas() {
    [["przListaMinisterio", "ministerio"], ["przListaEixo", "eixo"], ["przListaProjeto", "projeto"],
     ["przListaResponsavel", "responsavel"]].forEach(function (p) {
      document.getElementById(p[0]).innerHTML = (PRZ.opcoes[p[1]] || []).map(function (v) {
        return '<option value="' + esc(v) + '"></option>';
      }).join("");
    });
  }
  function abrirForm(d) {
    if (!PRZ.podeEditar) return;
    formEditando = d || null;
    preencherListas();
    var sel = document.getElementById("przFormStatus");
    var status = (PRZ.opcoes.status || ["A definir", "Aguardando", "Em andamento", "Concluído"]).slice();
    var atual = d ? (d.status || "A definir") : "A definir";
    if (status.indexOf(atual) === -1) status.push(atual);
    sel.innerHTML = status.map(function (s) { return '<option' + (s === atual ? " selected" : "") + ">" + esc(s) + "</option>"; }).join("");
    var vazio = function (v, padrao) { return v === padrao ? "" : (v || ""); };
    form.descricao.value = d ? d.descricao || "" : "";
    form.ministerio.value = d ? vazio(d.ministerio, "Sem órgão definido") : "";
    form.eixo.value = d ? vazio(d.eixo, "Sem eixo") : "";
    form.projeto.value = d ? vazio(d.projeto, "Sem projeto/frente") : "";
    form.data_prazo.value = d && d.data_prazo ? d.data_prazo : "";
    form.prazo_informado.value = d ? d.prazo_informado || "" : "";
    form.responsavel.value = d ? d.responsavel || "" : "";
    form.observacoes.value = d ? d.observacoes || "" : "";
    form.fonte.value = d ? d.fonte || "" : "";
    form.querySelectorAll(".invalido").forEach(function (c) { c.classList.remove("invalido"); });
    formErro.textContent = "";
    document.getElementById("przFormTitulo").textContent = d ? "Editar ação · ID " + d.id : "Nova ação";
    formSalvar.disabled = false;
    formSalvar.textContent = d ? "Salvar alterações" : "Cadastrar ação";
    formFundo.classList.add("aberto");
    formFundo.setAttribute("aria-hidden", "false");
    setTimeout(function () { form.descricao.focus(); }, 30);
  }
  function fecharForm() {
    formFundo.classList.remove("aberto");
    formFundo.setAttribute("aria-hidden", "true");
    formEditando = null;
  }
  function depoisDeGravar(r, msgOk, idAbrir) {
    if (r && r.ok) {
      PRZ.gravouAgora = true;
      avisar(msgOk);
      if (idAbrir) painel._id = String(idAbrir);
      return carregar(true).then(function () {
        if (idAbrir) {
          var d = PRZ.itens.find(function (x) { return String(x.id) === String(idAbrir); });
          if (d) abrirPainel(d);
        }
      });
    }
    if (r && r.conflito) carregar(true);
    throw new Error((r && (r.mensagem || r.erro)) || "Não foi possível gravar.");
  }
  form.addEventListener("submit", function (e) {
    e.preventDefault();
    var faltando = ["descricao", "projeto"].filter(function (c) { return !form[c].value.trim(); });
    form.querySelectorAll(".invalido").forEach(function (c) { c.classList.remove("invalido"); });
    faltando.forEach(function (c) { form[c].classList.add("invalido"); });
    if (faltando.length) { formErro.textContent = "Preencha a descrição da ação e o projeto/frente."; form[faltando[0]].focus(); return; }
    var item = {};
    ["descricao", "ministerio", "eixo", "projeto", "data_prazo", "prazo_informado", "status", "responsavel",
     "observacoes", "fonte"].forEach(function (c) { item[c] = form[c].value; });
    if (formEditando) item.id = formEditando.id;
    formSalvar.disabled = true;
    formErro.textContent = "";
    var editando = !!formEditando;
    api("prazos_salvar", item, PRZ.revisao).then(function (r) {
      return depoisDeGravar(r, editando ? "Ação atualizada" : "Ação cadastrada", r && r.id).then(fecharForm);
    }).catch(function (err) {
      formSalvar.disabled = false;
      formErro.textContent = String(err && err.message || err);
    });
  });
  function excluir(d) {
    if (!PRZ.podeEditar) return;
    if (!window.confirm('Excluir a ação "' + (d.descricao || d.projeto) + '" (ID ' + d.id + ')?\n\n' +
        "Ela sai do Controle de Prazos para todos. Fica registrada no histórico e na cópia de segurança.")) return;
    api("prazos_excluir", d.id, PRZ.revisao).then(function (r) {
      fecharPainel();
      return depoisDeGravar(r, "Ação excluída", null);
    }).catch(function (err) { window.alert(String(err && err.message || err)); });
  }
  document.getElementById("przNova").onclick = function () { abrirForm(null); };
  document.getElementById("przFormFechar").onclick = fecharForm;
  document.getElementById("przFormCancelar").onclick = fecharForm;
  formFundo.addEventListener("mousedown", function (e) { if (e.target === formFundo) fecharForm(); });
  // ao digitar num campo marcado como faltando, a marca e o aviso somem
  form.addEventListener("input", function (e) {
    if (e.target.classList.contains("invalido") && e.target.value.trim()) {
      e.target.classList.remove("invalido");
      if (!form.querySelector(".invalido")) formErro.textContent = "";
    }
  });
  document.getElementById("przExportar").onclick = function () {
    var b = this;
    b.disabled = true;
    var fim = function () { b.disabled = false; };
    if (window.pywebview && window.pywebview.api && window.pywebview.api.prazos_exportar_excel) {
      api("prazos_exportar_excel").then(function (r) { fim(); avisar((r && r.mensagem) || "Planilha exportada"); })
        .catch(function (e) { fim(); avisar("Erro ao exportar: " + e); });
      return;
    }
    fetch("/api/prazos_exportar_excel", {method: "POST", headers: {"Content-Type": "application/json"}, body: "[]"})
      .then(function (resp) {
        var tipo = resp.headers.get("Content-Type") || "";
        if (tipo.indexOf("json") !== -1) return resp.json().then(function (r) { throw new Error(r.mensagem || r.erro || "erro"); });
        var nome = ((resp.headers.get("Content-Disposition") || "").match(/filename="([^"]+)"/) || [])[1] || "Controle de Prazos.xlsx";
        return resp.blob().then(function (blob) {
          var a = document.createElement("a");
          a.href = URL.createObjectURL(blob); a.download = nome; document.body.appendChild(a); a.click();
          setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
        });
      }).then(fim).catch(function (e) { fim(); avisar("Erro ao exportar: " + (e.message || e)); });
  };
  document.getElementById("przPainelFechar").onclick = fecharPainel;
  document.getElementById("przCorpo").addEventListener("click", function (e) {
    if (!painel.classList.contains("aberto") || arrastou) return;
    if (e.target.closest("[data-id]")) return;
    fecharPainel();
  });

  // ---------------- cabeçalho ----------------
  var caixaBusca = document.getElementById("przBuscaGeral");
  var timerBusca = null;
  function definirBusca(texto) {
    PRZ.busca = String(texto || "");
    if (document.activeElement !== caixaBusca) caixaBusca.value = PRZ.busca;
    caixaBusca.parentNode.classList.toggle("ativo", !!PRZ.busca.trim());
    clearTimeout(timerBusca);
    timerBusca = setTimeout(filtrosMudaram, 150);
  }
  caixaBusca.addEventListener("input", function () { definirBusca(caixaBusca.value); });
  caixaBusca.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && caixaBusca.value) { e.stopPropagation(); caixaBusca.value = ""; definirBusca(""); }
  });
  document.getElementById("przBuscaGeralLimpar").onclick = function () { caixaBusca.value = ""; definirBusca(""); caixaBusca.focus(); };
  document.getElementById("przLimparBase").onclick = function () {
    if (window.limparFiltrosBaseAtiva) window.limparFiltrosBaseAtiva(false); else limpar();
  };
  var btnTema = document.getElementById("przTema");
  if (window.alternarTema) {
    btnTema.onclick = function () { window.alternarTema(); };
    if (window.sincronizarBotoesTema) window.sincronizarBotoesTema();
  } else btnTema.style.display = "none";
  var btnInicio = document.getElementById("przInicio");
  if (btnInicio && window.irParaHome) btnInicio.onclick = window.irParaHome;
  var seletor = overlay.querySelector(".seletor-secao.troca-base");
  if (seletor && window.configurarSeletorSecao) window.configurarSeletorSecao(seletor);
  function relogio() {
    var a = new Date();
    document.getElementById("przRelogio").textContent = a.toLocaleDateString("pt-BR") + " " +
      a.toLocaleTimeString("pt-BR", {hour: "2-digit", minute: "2-digit"});
  }
  relogio();
  setInterval(relogio, 15000);
  // a cada minuto, com a seção à vista, refaz contagens/cores (o "hoje" anda)
  setInterval(function () {
    if (overlay.classList.contains("aberto") && PRZ.visao !== "mapa") renderizarDash();
  }, 60000);
  document.addEventListener("keydown", function (e) {
    if (e.key !== "Escape" || !overlay.classList.contains("aberto")) return;
    if (formFundo.classList.contains("aberto")) { e.stopPropagation(); fecharForm(); return; }
    if (painel.classList.contains("aberto")) { e.stopPropagation(); fecharPainel(); }
  });

  function limpar() {
    BLOCOS.forEach(function (b) { PRZ.sel[b.chave].clear(); });
    PRZ.busca = "";
    buscaBloco = {};
    caixaBusca.value = "";
    filtrosMudaram();
  }

  // ---------------- ganchos usados por trocarBase() e cia. ("CGAPE - BALANÇO PAC.py") ----------------
  window.przGarantirCarga = function () { if (!PRZ.carregado && !PRZ.carregando) carregar(false); };
  window.przAbrirDash = function () {
    overlay.classList.add("aberto");
    overlay.setAttribute("aria-hidden", "false");
    if (!PRZ.carregado) carregar(false);
    else renderizarDash();
    if (PRZ.visao === "mapa") setTimeout(function () { desenharMapa(true); }, 30);
  };
  window.przFecharDash = function () {
    overlay.classList.remove("aberto");
    overlay.setAttribute("aria-hidden", "true");
    fecharPainel();
  };
  window.przContarFiltros = nFiltros;
  window.przLimpar = limpar;
  window.przBuscaAtual = function () { return PRZ.busca; };
  window.przDefinirBusca = function (t) { caixaBusca.value = String(t || ""); definirBusca(t); };
  window.przCabecalho = function () {
    return {titulo: "CONTROLE DE PRAZOS",
      data: PRZ.atualizado ? "Controle de Prazos (Pauta do Governador) atualizado em " + PRZ.atualizado
        : "Controle de Prazos" + (PRZ.carregando ? " — carregando…" : "")};
  };
})();
</script>
"""
