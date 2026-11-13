# -*- coding: utf-8 -*-
"""Resumo das Ações Prioritárias e PAC — janela própria do dash.

Portado do Relatório SSP (pedido do usuário em 12/11/2026: "incluir esse
painel no dash do panorama pac conforme feito no relatório SSP"). Base
própria, a mesma do MONITORAMENTO ("Ações do Cronograma.xlsx", lida da
pasta do programa); independe da planilha e dos filtros do PAC.

Ligado ao módulo principal pelo mesmo padrão do gerar_pptx_pac:
configurar(sys.modules[__name__]).
"""
import math
import os
import re
import unicodedata
from datetime import datetime

import pandas as pd

pac = None


def configurar(modulo):
    global pac
    pac = modulo


RA_PLANILHA = "Ações do Cronograma.xlsx"
RA_COL_EIXO, RA_COL_STATUS = "Eixo", "Status"
RA_COL_VALOR, RA_COL_ORCAMENTO = "Valor de Investimento", "Valor do Orçamento"
RA_EIXOS = ["Mobilidade", "Educação", "Saneamento",
            "Prevenção a Desastres Naturais", "Turismo e Cultura", "Saúde"]
RA_CORES_TOP3, RA_COR_TOTAL = ["2E7D32", "F2B632", "2F7F92"], "0B1F3F"
RA_AZUL, RA_CLARO = "1F5C85", "EEF3F8"
# painel de filtros: (chave, coluna da base, rótulo)
RA_FILTROS = [("eixo", RA_COL_EIXO, "Eixo"),
              ("status", RA_COL_STATUS, "Status"),
              ("territorio", "Território Identidade", "Território"),
              ("municipio", "Município", "Município"),
              ("executor", "Executor", "Executor"),
              ("programa", "Programa de Transferência de Recursos", "Programa"),
              ("escopo", "Escopo", "Escopo"),
              ("objeto", "Objeto", "Objeto"),
              ("prazo", "Prazo Contratual", "Prazo contratual")]
# filtros de data (intervalo de/até). Com um intervalo preenchido, ações sem
# data nessa coluna ficam de fora.
RA_DATAS = [("atualizacao", "Última Atualização", "Última atualização"),
            ("estudos", "Conclusão dos Estudos", "Conclusão dos estudos"),
            ("projeto", "Conclusão do Projeto", "Conclusão do projeto"),
            ("aviso", "Aviso de Licitação", "Aviso de licitação"),
            ("licitacao", "Licitação / Contratação",
             "Licitação / contratação"),
            ("os", "Ordem de Serviço", "Ordem de serviço"),
            ("entrega", "Entrega Obra", "Entrega da obra"),
            ("equipamentos", "Instalação Equipamentos",
             "Instalação de equipamentos"),
            ("funcionalidade", "Funcionalidade", "Funcionalidade")]
RA_SEM_INFO = "(não informado)"

# Filtro "30 maiores municípios" (pedido do usuário em 12/11/2026). Critério
# escolhido: POPULAÇÃO — estimativa IBGE 2026 (SIDRA, tabela 6579). Comparado
# com o eleitorado (TSE, aptos no 1º turno de 2024) e com o Censo 2022: os
# três dão as mesmas 30 cidades. Lista fixa, em ordem de população; revisar
# quando sair nova estimativa do IBGE (a 31ª hoje é Bom Jesus da Lapa).
RA_TOP30_FONTE = "30 maiores municípios por população (IBGE, estimativa 2026)"
RA_TOP30 = [
    "Salvador", "Feira de Santana", "Vitória da Conquista", "Camaçari",
    "Juazeiro", "Lauro de Freitas", "Itabuna", "Ilhéus", "Porto Seguro",
    "Barreiras", "Jequié", "Alagoinhas", "Teixeira de Freitas", "Eunápolis",
    "Simões Filho", "Luís Eduardo Magalhães", "Paulo Afonso",
    "Santo Antônio de Jesus", "Guanambi", "Valença", "Jacobina", "Serrinha",
    "Irecê", "Senhor do Bonfim", "Casa Nova", "Dias d'Ávila",
    "Campo Formoso", "Candeias", "Brumado", "Conceição do Coité"]


def _chave_municipio(nome):
    """Só letras e números, sem acento e em maiúsculas: "Dias D'Ávila" e
    "DIAS D AVILA" viram a mesma chave."""
    s = unicodedata.normalize("NFKD", str(nome or ""))
    return "".join(c for c in s if c.isalnum()
                   and not unicodedata.combining(c)).upper()


RA_TOP30_CHAVES = {_chave_municipio(m) for m in RA_TOP30}
_ra_cache = None   # (caminho, data do arquivo, base)

# PPTX (mesmas cores/fontes do slide do SSP)
PE_FONTE = "Segoe UI"
PE_FONTE_BLACK = "Segoe UI Black"
PE_COR = {"navy": "0B1F3F", "cinza": "5B6576", "gov": "4FA3E0",
          "subcab": "C9D6EA"}


# ---- utilitários ------------------------------------------------------------
def caminho_planilha():
    caminho = pac.caminho_recurso(RA_PLANILHA) if pac else RA_PLANILHA
    return caminho if os.path.exists(caminho) else None


def pasta_downloads():
    pasta = getattr(pac, "PASTA_DOWNLOADS_PADRAO", None)
    if pasta and os.path.isdir(pasta):
        return pasta
    padrao = os.path.join(os.path.expanduser("~"), "Downloads")
    return padrao if os.path.isdir(padrao) else os.getcwd()


def _normalizar(texto):
    if texto is None or (isinstance(texto, float) and math.isnan(texto)):
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in s if not unicodedata.combining(c)).strip().lower()


def _pe_larg(texto, fs, estilo="n"):
    """Largura estimada (px) do texto em Segoe UI (o PowerPoint não mede)."""
    k = {"n": 0.52, "b": 0.56, "k": 0.64}[estilo]
    return sum(fs * k * (1.18 if c.isupper() else
                         (0.45 if c in " .,:;|/·'" else 1.0))
               for c in str(texto))


def _pe_fs(texto, largura, fs, estilo="n", minimo=11):
    while fs > minimo and _pe_larg(texto, fs, estilo) > largura:
        fs -= 1
    return fs


def _ra_numero(v):
    """Valor da base ("R$ 12.752.146,77", número ou vazio) em float."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = re.sub(r"[^\d,.-]", "", str(v))
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _ra_rs(v):
    if v >= 999.95e6:
        # "BI", "MI" e "mil" (pedido do usuário em 13/11/2026)
        return "R$ " + f"{v / 1e9:.2f}".replace(".", ",") + " BI"
    if v >= 1e6:
        return "R$ " + f"{v / 1e6:.1f}".replace(".", ",") + " MI"
    if v > 0:
        return f"R$ {v / 1e3:.0f} mil"
    return "R$ 0"


def _ra_qtd(q):
    return "1 ação" if q == 1 else f"{q:,} ações".replace(",", ".")


def _ra_larguras(dados, disponivel, fs_max=19.0):
    """Maior fonte (até fs_max) em que os valores cabem em 1 linha, com a
    largura de cada coluna (eixos + Total)."""
    n = len(dados["eixos"])
    linhas = dados["linhas"] + [dados["total"]]
    valores = [[lin["celulas"][j]["txt"] for lin in linhas
                if lin["celulas"][j]] for j in range(n)]
    palavras = [e.split() for e in dados["eixos"]]
    totais = [lin["total"]["txt"] for lin in linhas if lin["total"]]

    def largura(txts, fs, cab):
        maior = max([_pe_larg(t, fs, "b") for t in txts]
                    + [_pe_larg(p, min(fs, 12), "n") for p in cab] + [0])
        return maior + 10

    fs = fs_max
    while True:
        cols = [largura(valores[j], fs, palavras[j]) for j in range(n)]
        cols.append(largura(totais, fs + 2, ["Total"]))
        if sum(cols) <= disponivel or fs <= 10:
            break
        fs -= 0.5
    if sum(cols) > disponivel:
        cols = [c * disponivel / sum(cols) for c in cols]
    else:
        cols = [c + (disponivel - sum(cols)) / len(cols) for c in cols]
    return fs, cols


def _ra_opcao_filtro(serie):
    """Vazio vira "(não informado)" e grafias que só diferem em caixa ou
    espaços viram uma só."""
    s = serie.map(lambda v: " ".join(str(v).split()) if pd.notna(v) else "")
    s = s.where(s != "", RA_SEM_INFO)
    chave = s.str.casefold()
    return chave.map(s.groupby(chave).agg(lambda g: g.mode().iloc[0]))


def _ra_data(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return pd.NaT
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v).normalize()
    t = str(v).strip()
    if not t:
        return pd.NaT
    formato = "%Y-%m-%d" if re.match(r"\d{4}-\d{2}-\d{2}", t) else None
    d = pd.to_datetime(t[:10] if formato else t, format=formato,
                       dayfirst=True, errors="coerce")
    return d.normalize() if pd.notna(d) else pd.NaT


def _ra_ordem(v):
    """Ordem natural: "36 Meses" antes de "120 Meses"."""
    return [(0, int(p)) if p.isdigit() else (1, p)
            for p in re.split(r"(\d+)", _normalizar(v))]


def _ra_base(caminho):
    """Planilha já preparada; só é lida de novo se o arquivo tiver mudado."""
    global _ra_cache
    data_arquivo = os.path.getmtime(caminho)
    if _ra_cache and _ra_cache[:2] == (caminho, data_arquivo):
        return _ra_cache[2]
    df = pd.read_excel(caminho, dtype=object)
    faltando = [c for c in (RA_COL_EIXO, RA_COL_STATUS, RA_COL_VALOR,
                            RA_COL_ORCAMENTO) if c not in df.columns]
    if faltando:
        raise ValueError("coluna(s) ausente(s): " + ", ".join(faltando))
    df = df[df[RA_COL_EIXO].notna() & df[RA_COL_STATUS].notna()].copy()
    df = df.reset_index(drop=True)
    df["_eixo"] = df[RA_COL_EIXO].astype(str).str.strip()
    # sem valor de investimento: usa o orçamento e marca para o asterisco
    investimento = df[RA_COL_VALOR].map(_ra_numero)
    orcamento = df[RA_COL_ORCAMENTO].map(_ra_numero)
    df["_orc"] = (investimento == 0) & (orcamento > 0)
    df["_valor"] = investimento.where(~df["_orc"], orcamento)
    for chave, coluna, _ in RA_FILTROS:
        if coluna in df.columns:
            df["_f_" + chave] = _ra_opcao_filtro(df[coluna])
    for chave, coluna, _ in RA_DATAS:
        if coluna in df.columns:
            df["_d_" + chave] = pd.to_datetime(df[coluna].map(_ra_data),
                                               errors="coerce")
    df["_top30"] = (df["Município"].map(_chave_municipio).isin(RA_TOP30_CHAVES)
                    if "Município" in df.columns else False)
    _ra_cache = (caminho, data_arquivo, df)
    return df


# Seletor "ENTREGUE" (pedido do usuário em 13/11/2026): Obras Concluídas e
# Obras Inauguradas viram um só status, somando ações e valores. Não é
# filtro: não restringe o recorte e não é zerado pelo LIMPAR FILTROS.
RA_STATUS_ENTREGUE = "01 - Entregue"


def _ra_juntar_entregues(base):
    """Cópia da base com os status de concluída/inaugurada trocados por
    RA_STATUS_ENTREGUE, e a lista dos status que foram juntados."""
    st = base["_f_status"]
    chave = st.map(lambda v: _normalizar(v).casefold())
    juntar = chave.str.contains("conclu") | chave.str.contains("inaugur")
    juntados = sorted(st[juntar].unique(), key=_ra_ordem)
    if not juntados:
        return base, []
    return base.assign(_f_status=st.where(~juntar, RA_STATUS_ENTREGUE)), juntados


# ---- dados ------------------------------------------------------------------
def resumo_acoes_dados(caminho, filtros=None):
    """Quadro por eixo e status, já formatado para a página e para o slide.
    filtros: {chave de RA_FILTROS: [valores], "d_<chave>": {de, ate}}."""
    base = _ra_base(caminho)
    filtros = filtros or {}
    entregue = bool(filtros.get("entregue"))
    juntados = []
    if entregue:
        base, juntados = _ra_juntar_entregues(base)
    campos = [f for f in RA_FILTROS if "_f_" + f[0] in base.columns]
    sel = {c: filtros.get(c) for c, _, _ in campos if filtros.get(c)}
    if juntados and sel.get("status"):
        sel["status"] = list(dict.fromkeys(
            RA_STATUS_ENTREGUE if v in juntados else v for v in sel["status"]))

    campos_data = [f for f in RA_DATAS if "_d_" + f[0] in base.columns]
    sel_datas = {}
    for c, _, _ in campos_data:
        faixa = filtros.get("d_" + c) or {}
        de = pd.to_datetime(faixa.get("de") or None, errors="coerce")
        ate = pd.to_datetime(faixa.get("ate") or None, errors="coerce")
        if pd.notna(de) or pd.notna(ate):
            sel_datas[c] = (de, ate)

    top30 = bool(filtros.get("top30"))

    def marca_datas(com_top30=True):
        marca = pd.Series(True, index=base.index)
        if top30 and com_top30:
            marca &= base["_top30"]
        for c, (de, ate) in sel_datas.items():
            d = base["_d_" + c]
            marca &= d.notna()
            if pd.notna(de):
                marca &= d >= de
            if pd.notna(ate):
                marca &= d <= ate
        return marca

    def recorte(exceto=None, com_top30=True):
        marca = marca_datas(com_top30)
        for c, valores in sel.items():
            if c != exceto:
                marca &= base["_f_" + c].isin(valores)
        return base[marca]

    df = recorte()

    # opções de cada filtro: o que existe dado o que os OUTROS filtros têm;
    # o que já foi escolhido fica na lista mesmo zerado, para poder desmarcar
    painel = []
    for c, _, rotulo in campos:
        qtd = recorte(c)["_f_" + c].value_counts()
        nomes = sorted(set(qtd.index) | set(sel.get(c, [])), key=_ra_ordem)
        painel.append({"chave": c, "rotulo": rotulo,
                       "opcoes": [{"v": v, "n": int(qtd.get(v, 0))}
                                  for v in nomes]})
    painel_datas = []
    for c, _, rotulo in campos_data:
        col = base["_d_" + c].dropna()
        painel_datas.append({
            "chave": c, "rotulo": rotulo,
            "min": f"{col.min():%Y-%m-%d}" if len(col) else "",
            "max": f"{col.max():%Y-%m-%d}" if len(col) else "",
            "n": int(df["_d_" + c].notna().sum())})

    # sem filtro: todos os eixos, mesmo sem ação; com filtro, só os que têm
    # ação; sempre do maior para o menor investimento
    por_eixo = df.groupby("_eixo")["_valor"].sum()
    filtrado = bool(sel or sel_datas or top30)
    eixos = sorted(set(df["_eixo"]) if filtrado
                   else set(RA_EIXOS) | set(df["_eixo"]),
                   key=lambda e: (-por_eixo.get(e, 0), e))

    def celula(parte):
        if not len(parte):
            return None
        return {"txt": _ra_rs(float(parte["_valor"].sum()))
                + ("*" if parte["_orc"].any() else ""),
                "qtd": _ra_qtd(len(parte))}

    linhas = [{"status": st,
               "celulas": [celula(df[(df["_f_status"] == st)
                                     & (df["_eixo"] == e)]) for e in eixos],
               "total": celula(df[df["_f_status"] == st])}
              for st in sorted(df["_f_status"].unique())]
    total = {"status": "TOTAL",
             "celulas": [celula(df[df["_eixo"] == e]) for e in eixos],
             "total": celula(df)}
    # quadro Status × Objeto (pedido do usuário em 12/11/2026): colunas
    # agrupadas pelo eixo a que o objeto pertence, na MESMA ordem de eixos
    # do quadro acima; dentro do eixo, do maior para o menor investimento.
    # Objeto presente em mais de um eixo aparece em cada um deles.
    grupos_obj, colunas_obj = [], []
    if "_f_objeto" in df.columns:
        for e in eixos:
            parte_e = df[df["_eixo"] == e]
            if not len(parte_e):
                continue
            por_obj = parte_e.groupby("_f_objeto")["_valor"].agg(["sum", "size"])
            nomes = sorted(por_obj.index,
                           key=lambda o: (-por_obj.loc[o, "sum"], -por_obj.loc[o, "size"], o))
            grupos_obj.append({"eixo": e, "objetos": nomes})
            colunas_obj += [(e, o) for o in nomes]
    vazio = df.iloc[0:0]
    por_seo = ({k: p for k, p in df.groupby(["_f_status", "_eixo", "_f_objeto"])}
               if colunas_obj else {})
    por_eo = ({k: p for k, p in df.groupby(["_eixo", "_f_objeto"])}
              if colunas_obj else {})
    linhas_obj = [{"status": l["status"],
                   "celulas": [celula(por_seo.get((l["status"], e, o), vazio))
                               for e, o in colunas_obj],
                   "total": l["total"]}
                  for l in linhas] if colunas_obj else []
    total_obj = {"status": "TOTAL",
                 "celulas": [celula(por_eo.get((e, o), vazio)) for e, o in colunas_obj],
                 "total": total["total"]}
    quadro_objeto = {"grupos": grupos_obj, "linhas": linhas_obj, "total": total_obj}

    # Cards de valor macro (pedido do usuário em 13/11/2026): sem os cards
    # dos 3 maiores eixos. VALOR TOTAL = estado todo (segue os demais
    # filtros, mas NÃO a chave dos 30 maiores); com a chave ligada, ao lado
    # direito entra o total dos 30 maiores municípios (= quadros, mapa, lista).
    vazio_card = {"txt": "R$ 0", "qtd": _ra_qtd(0)}
    df_estado = recorte(com_top30=False) if top30 else df

    def nota_asterisco(parte):
        # explicação do "*" logo abaixo do card (pedido do usuário em 13/11/2026)
        orc = parte[parte["_orc"]]
        if not len(orc):
            return ""
        return (f"* Inclui {RA_COL_ORCAMENTO} em {_ra_qtd(len(orc))} sem {RA_COL_VALOR} "
                f"({_ra_rs(float(orc['_valor'].sum()))}).")

    cards = [{"rotulo": "VALOR TOTAL · ESTADO" if top30 else "VALOR TOTAL", "cor": RA_COR_TOTAL,
              **(celula(df_estado) or vazio_card), "nota": nota_asterisco(df_estado)}]
    aviso_cards = ""
    if top30:
        cards.append({"rotulo": "VALOR TOTAL · 30 MAIORES MUNICÍPIOS", "cor": RA_CORES_TOP3[2],
                      **(celula(df) or vazio_card), "nota": nota_asterisco(df)})
        aviso_cards = ("O total do estado não segue a chave dos 30 maiores municípios; o total "
                       "dos 30 maiores, o mapa, os quadros e a lista seguem.")

    data_base = datetime.fromtimestamp(os.path.getmtime(caminho))
    nota = (f"Fonte: Ações do Cronograma, coluna {RA_COL_VALOR} "
            f"(base de {data_base:%d/%m/%Y}).")
    filtros_txt = ""
    if filtrado:
        def faixa_txt(de, ate):
            if pd.notna(de) and pd.notna(ate):
                return f"{de:%d/%m/%Y} a {ate:%d/%m/%Y}"
            return (f"a partir de {de:%d/%m/%Y}" if pd.notna(de)
                    else f"até {ate:%d/%m/%Y}")

        def resumo_filtros(limite):
            return "; ".join(
                ([RA_TOP30_FONTE] if top30 else [])
                + [f"{rotulo}: " + (", ".join(sel[c]) if len(sel[c]) <= limite
                                  else f"{len(sel[c])} selecionado(s)")
                 for c, _, rotulo in campos if c in sel]
                + [f"{rotulo}: {faixa_txt(*sel_datas[c])}"
                   for c, _, rotulo in campos_data if c in sel_datas])
        escolhidos = resumo_filtros(3)
        if len(escolhidos) > 220:
            escolhidos = resumo_filtros(0)
        # lista completa (todos os valores escolhidos) para o topo do HTML
        filtros_txt = resumo_filtros(10 ** 6)
        nota = f"Filtros: {escolhidos}. {nota}"
        if top30:
            nota += " Valor total · estado: todos os municípios (sem a chave dos 30 maiores)."
    if juntados:
        nota += (" Status Entregue = "
                 + " + ".join(re.sub(r"^\d+\s*-\s*", "", s) for s in juntados) + ".")
    com_orcamento = df[df["_orc"]]
    if len(com_orcamento):
        nota += (f" * Inclui {RA_COL_ORCAMENTO} em "
                 f"{_ra_qtd(len(com_orcamento))} sem {RA_COL_VALOR} "
                 f"({_ra_rs(float(com_orcamento['_valor'].sum()))}).")
    sem_valor = int((df["_valor"] == 0).sum())
    if sem_valor:
        nota += (f" {_ra_qtd(sem_valor)} sem nenhum dos dois valores "
                 "somam R$ 0.")
    return {"eixos": eixos, "cards": cards,
            "linhas": linhas, "total": total, "nota": nota,
            "quadro_objeto": quadro_objeto,
            "filtros": painel, "datas": painel_datas,
            "filtrado": filtrado, "vazio": df.empty, "top30": top30,
            "entregue": bool(juntados), "aviso_cards": aviso_cards, "filtros_txt": filtros_txt,
            "top30_fonte": RA_TOP30_FONTE,
            "ids": [int(i) for i in df.index],
            "versao": str(os.path.getmtime(caminho))}


def _ra_colunas_tabela(base):
    """[(rótulo, tipo, chave do filtro ou "", série)] da tabela de ações."""
    cols = [(rot, "txt", "", base[col].map(
                lambda v: " ".join(str(v).split()) if pd.notna(v) else ""))
            for col, rot in (("Código", "Código"), ("Ação", "Ação"))
            if col in base.columns]
    cols += [(rot, "txt", c, base["_f_" + c].where(
                 base["_f_" + c] != RA_SEM_INFO, ""))
             for c, _, rot in RA_FILTROS if "_f_" + c in base.columns]
    cols.append(("Valor (R$)", "valor", "", base["_valor"]))
    cols += [(rot, "data", "d_" + c, base["_d_" + c])
             for c, _, rot in RA_DATAS if "_d_" + c in base.columns]
    return cols


def resumo_acoes_tabela(caminho):
    """Todas as linhas da base para a tabela da página."""
    base = _ra_base(caminho)
    cols = _ra_colunas_tabela(base)
    series = []
    for _rot, tipo, _c, s in cols:
        if tipo == "data":
            s = s.map(lambda d: f"{d:%Y-%m-%d}" if pd.notna(d) else "")
        elif tipo == "valor":
            s = s.astype(float).round(2)
        series.append(s.tolist())
    return {"colunas": [{"rotulo": r, "tipo": t, "filtro": c}
                        for r, t, c, _ in cols],
            "linhas": [list(v) for v in zip(*series)],
            "orc": base["_orc"].astype(bool).tolist(),
            "versao": str(os.path.getmtime(caminho))}


def gerar_xlsx_resumo_acoes(caminho_base, dados, caminho):
    """Planilha (.xlsx) das ações do recorte."""
    base = _ra_base(caminho_base)
    linhas = base.iloc[dados["ids"]]
    tabela = pd.DataFrame({rot: s.loc[linhas.index]
                           for rot, _t, _c, s in _ra_colunas_tabela(base)})
    tabela.insert(tabela.columns.get_loc("Valor (R$)") + 1,
                  "Valor do orçamento?",
                  linhas["_orc"].map({True: "Sim", False: ""}))
    with pd.ExcelWriter(caminho, engine="openpyxl") as escritor:
        tabela.to_excel(escritor, index=False, sheet_name="Ações")
        pd.DataFrame({"Resumo": [dados["nota"]]}).to_excel(
            escritor, index=False, sheet_name="Filtros")
        ws = escritor.sheets["Ações"]
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = ws.dimensions
        tipos = {rot: t for rot, t, _c, _s in _ra_colunas_tabela(base)}
        for i, nome in enumerate(tabela.columns, start=1):
            letra = ws.cell(row=1, column=i).column_letter
            tipo = tipos.get(nome, "txt")
            largura = {"data": 13, "valor": 18}.get(tipo)
            if largura is None:
                maior = tabela[nome].astype(str).str.len().quantile(.9)
                largura = max(len(nome) + 2, min(60, int(maior or 0) + 2))
            ws.column_dimensions[letra].width = largura
            formato = {"data": "DD/MM/YYYY",
                       "valor": '"R$" #,##0.00'}.get(tipo)
            if formato:
                for cel in ws[letra][1:]:
                    cel.number_format = formato
        escritor.sheets["Filtros"].column_dimensions["A"].width = 120


def gerar_pptx_resumo_acoes(dados, caminho):
    """Slide único (16:9, editável, com tabela nativa) do resumo."""
    from lxml import etree
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Pt

    navy, branco = PE_COR["navy"], "FFFFFF"
    cinza, linha_tab = PE_COR["cinza"], "D9DFE8"
    px = lambda v: Emu(int(v * 9525))   # noqa: E731 — slide de 1280 x 720 px

    def escrever(par, txt, tam, cor, estilo="n", alin=PP_ALIGN.LEFT):
        par.alignment = alin
        r = par.add_run()
        r.text = txt
        r.font.size = Pt(tam * 0.75)
        r.font.color.rgb = RGBColor.from_string(cor)
        r.font.name = PE_FONTE_BLACK if estilo == "k" else PE_FONTE
        r.font.bold = estilo == "b"
        return r

    def texto(x, y, w, h, linhas, alin=PP_ALIGN.LEFT, vert=MSO_ANCHOR.TOP):
        tf = slide.shapes.add_textbox(px(x), px(y), px(w), px(h)).text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = vert
        for i, (txt, tam, cor, estilo) in enumerate(linhas):
            par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            escrever(par, txt, tam, cor, estilo, alin)
        return tf

    def retangulo(x, y, w, h, fundo, nome, raio=0):
        tipo = (MSO_SHAPE.ROUNDED_RECTANGLE if raio else MSO_SHAPE.RECTANGLE)
        s = slide.shapes.add_shape(tipo, px(x), px(y), px(w), px(h))
        if raio:
            s.adjustments[0] = min(0.5, raio / min(w, h))
        s.fill.solid()
        s.fill.fore_color.rgb = RGBColor.from_string(fundo)
        s.line.fill.background()
        s.shadow.inherit = False
        s.name = nome

    def celula(cel, linhas, fundo, alin=PP_ALIGN.CENTER):
        tf = cel.text_frame
        tf.word_wrap = True
        for i, (txt, tam, cor, estilo) in enumerate(linhas):
            par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            escrever(par, txt, tam, cor, estilo, alin)
        cel.vertical_anchor = MSO_ANCHOR.MIDDLE
        cel.margin_left = px(8 if alin == PP_ALIGN.LEFT else 4)
        cel.margin_right = px(4)
        cel.margin_top = cel.margin_bottom = px(1 if compacto else 2)
        cel.fill.solid()
        cel.fill.fore_color.rgb = RGBColor.from_string(fundo)
        # as bordas precisam vir antes do preenchimento no XML da célula
        tc_pr = cel._tc.get_or_add_tcPr()
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        for i, lado in enumerate(("lnL", "lnR", "lnT", "lnB")):
            ln = etree.Element(f"{{{ns}}}{lado}", w="9525", cap="flat",
                               cmpd="sng", algn="ctr")
            sol = etree.SubElement(ln, f"{{{ns}}}solidFill")
            etree.SubElement(sol, f"{{{ns}}}srgbClr", val=linha_tab)
            tc_pr.insert(i, ln)

    def conteudo(c, tam, cor=navy, cor_qtd=cinza):
        if not c:
            return [("–", 10 if compacto else 15, cinza, "n")]
        tam_qtd = 9 if compacto else max(10.5, min(14, tam - 1.5))
        return [(c["txt"], tam, cor, "b"), (c["qtd"], tam_qtd, cor_qtd, "n")]

    compacto = len(dados["linhas"]) > 5

    prs = Presentation()
    prs.slide_width, prs.slide_height = px(1280), px(720)
    slide = prs.slides.add_slide(prs.slide_layouts[5])   # só título

    retangulo(0, 0, 1280, 92, navy, "Cabeçalho")
    titulo = slide.shapes.title
    titulo.left, titulo.top = px(36), px(16)
    titulo.width, titulo.height = px(960), px(38)
    tf = titulo.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    escrever(tf.paragraphs[0], "AÇÕES EM MONITORAMENTO", 26,
             branco, "k")
    arvore = slide.shapes._spTree
    arvore.remove(titulo._element)
    arvore.append(titulo._element)
    texto(36, 56, 960, 18,
          [("COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS ESTRATÉGICAS",
            11.5, PE_COR["subcab"], "b")])
    texto(994, 26, 250, 40,
          [("GOVERNO DO ESTADO DA BAHIA", 12, PE_COR["gov"], "b"),
           ("CASA CIVIL", 12, PE_COR["gov"], "b")], PP_ALIGN.RIGHT)

    cards = dados["cards"]
    larg = min(400, (1208 - 16 * (len(cards) - 1)) / len(cards))
    for i, c in enumerate(cards):
        x = 36 + i * (larg + 16)
        y, h = (104, 82) if compacto else (110, 122)
        retangulo(x, y, larg, h, RA_CLARO, f"Card {c['rotulo']}", raio=8)
        retangulo(x, y, 6, h, c["cor"], f"Marca {c['rotulo']}")
        if compacto:
            texto(x + 24, y + 6, larg - 36, 16, [(c["rotulo"], 12, cinza, "b")])
            texto(x + 24, y + 22, larg - 36, 36,
                  [(c["txt"], _pe_fs(c["txt"], larg - 36, 28, "k", 16),
                    RA_AZUL, "k")], vert=MSO_ANCHOR.MIDDLE)
            texto(x + 24, y + 59, larg - 36, 20, [(c["qtd"], 15, cinza, "b")])
        else:
            texto(x + 24, y + 12, larg - 36, 20,
                  [(c["rotulo"], 14, cinza, "b")])
            texto(x + 24, y + 32, larg - 36, 58,
                  [(c["txt"], _pe_fs(c["txt"], larg - 36, 40, "k", 20),
                    RA_AZUL, "k")], vert=MSO_ANCHOR.MIDDLE)
            texto(x + 24, y + 92, larg - 36, 24, [(c["qtd"], 17, cinza, "b")])

    eixos, linhas = dados["eixos"], dados["linhas"]
    n_lin, n_col = len(linhas) + 2, len(eixos) + 2
    if compacto:
        topo, cab, rod = 194, 42, 44
        alturas = [cab] + [(656 - topo - cab - rod) / len(linhas)] * len(linhas)
        alturas.append(rod)
    else:
        topo = 246
        alturas = [64] + [min(84, 252 / len(linhas))] * len(linhas) + [88]
    grade = slide.shapes.add_table(n_lin, n_col, px(36), px(topo), px(1208),
                                   px(sum(alturas)))
    grade.name = "Quadro resumo"
    tab = grade.table
    tab.first_row = tab.horz_banding = False
    larg_status = 200 if compacto else 140
    fs, larguras = _ra_larguras(dados, 1208 - larg_status,
                                10.5 if compacto else 19.0)
    fs_cab = min(fs, 12)
    fs_st, fs_rot = (10.5, 12) if compacto else (13, 15)
    tab.columns[0].width = px(larg_status)
    for j, w in enumerate(larguras, 1):
        tab.columns[j].width = px(w)
    for i, a in enumerate(alturas):
        tab.rows[i].height = px(a)

    celula(tab.cell(0, 0), [("Status", fs_rot, navy, "b")], branco,
           PP_ALIGN.LEFT)
    for j, e in enumerate(eixos, 1):
        celula(tab.cell(0, j), [(e, fs_cab, navy, "n")], branco)
    celula(tab.cell(0, n_col - 1), [("Total", fs_rot, branco, "b")], navy)

    for i, l in enumerate(linhas, 1):
        celula(tab.cell(i, 0), [(l["status"], fs_st, navy, "n")], branco,
               PP_ALIGN.LEFT)
        for j, c in enumerate(l["celulas"], 1):
            celula(tab.cell(i, j), conteudo(c, fs, RA_AZUL), branco)
        celula(tab.cell(i, n_col - 1), conteudo(l["total"], fs), branco)

    ult, tot = n_lin - 1, dados["total"]
    celula(tab.cell(ult, 0), [("TOTAL", fs_rot, navy, "b")], RA_CLARO,
           PP_ALIGN.LEFT)
    for j, c in enumerate(tot["celulas"], 1):
        celula(tab.cell(ult, j), conteudo(c, fs), RA_CLARO)
    celula(tab.cell(ult, n_col - 1),
           conteudo(tot["total"], fs + 2, branco, branco), navy)

    nota = texto(36, 664, 1208, 22, [(dados["nota"], 14, cinza, "n")])
    for par in nota.paragraphs:
        for r in par.runs:
            r.font.italic = True

    # ---- Status × Objeto: slide(s) seguinte(s), objetos agrupados pelo
    # eixo (mesma ordem do quadro Status × Eixo). Com muitos objetos o
    # quadro é dividido em partes, sem cortar um eixo ao meio; o Total vai
    # na última parte.
    q = dados.get("quadro_objeto") or {}
    if q.get("grupos"):
        compacto = True
        larg_st, larg_tot, fs_v = 170, 96, 11
        linhas_q = q["linhas"] + [q["total"]]

        def larg_col(k, nome):
            txts = [l["celulas"][k]["txt"] for l in linhas_q if l["celulas"][k]]
            palavras = [_pe_larg(p, 9.5, "n") for p in str(nome).split()] or [0]
            return max(max([_pe_larg(t, fs_v, "b") for t in txts] + [0]),
                       max(palavras)) + 12

        blocos, k = [], 0
        for g in q["grupos"]:
            ix = list(range(k, k + len(g["objetos"])))
            k += len(g["objetos"])
            blocos.append((g, ix, [max(58, larg_col(i, o))
                                   for i, o in zip(ix, g["objetos"])]))
        disp = 1208 - larg_st - larg_tot
        partes, atual, soma = [], [], 0
        for b in blocos:
            w = sum(b[2])
            if atual and soma + w > disp:
                partes.append(atual)
                atual, soma = [], 0
            atual.append(b)
            soma += w
        if atual:
            partes.append(atual)

        for n_parte, parte in enumerate(partes, 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])   # em branco
            retangulo(0, 0, 1280, 92, navy, "Cabeçalho")
            texto(36, 16, 960, 38, [("AÇÕES EM MONITORAMENTO", 26,
                                     branco, "k")], vert=MSO_ANCHOR.MIDDLE)
            texto(36, 56, 960, 18,
                  [("COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS "
                    "ESTRATÉGICAS", 11.5, PE_COR["subcab"], "b")])
            texto(994, 26, 250, 40,
                  [("GOVERNO DO ESTADO DA BAHIA", 12, PE_COR["gov"], "b"),
                   ("CASA CIVIL", 12, PE_COR["gov"], "b")], PP_ALIGN.RIGHT)
            sufixo = (f"  (parte {n_parte} de {len(partes)})"
                      if len(partes) > 1 else "")
            texto(36, 102, 1208, 22,
                  [("QUADRO RESUMO — STATUS × OBJETO" + sufixo, 15, navy, "b")])

            ultimo = n_parte == len(partes)
            ix = [i for _g, idx, _w in parte for i in idx]
            ws = [w for _g, _idx, wl in parte for w in wl]
            n_col = 1 + len(ix) + (1 if ultimo else 0)
            n_lin = 2 + len(q["linhas"]) + 1
            topo = 132
            h_g, h_o, h_t = 24, 44, 40
            h_l = min(46, (652 - topo - h_g - h_o - h_t) / max(1, len(q["linhas"])))
            alturas = [h_g, h_o] + [h_l] * len(q["linhas"]) + [h_t]
            # o que sobrar da largura vai para as colunas de objeto
            usado = larg_st + sum(ws) + (larg_tot if ultimo else 0)
            extra = max(0, 1208 - usado) / (len(ws) + (1 if ultimo else 0))
            grade = slide.shapes.add_table(n_lin, n_col, px(36), px(topo),
                                           px(1208), px(sum(alturas)))
            grade.name = "Quadro Status x Objeto"
            tab = grade.table
            tab.first_row = tab.horz_banding = False
            tab.columns[0].width = px(larg_st)
            for j, w in enumerate(ws, 1):
                tab.columns[j].width = px(w + extra)
            if ultimo:
                tab.columns[n_col - 1].width = px(larg_tot + extra)
            for i, a in enumerate(alturas):
                tab.rows[i].height = px(a)

            fs_st = 10.5 if len(q["linhas"]) > 5 else 13
            tab.cell(0, 0).merge(tab.cell(1, 0))
            celula(tab.cell(0, 0), [("Status", 12, navy, "b")], branco, PP_ALIGN.LEFT)
            j = 1
            for g, idx, _w in parte:
                if len(idx) > 1:
                    tab.cell(0, j).merge(tab.cell(0, j + len(idx) - 1))
                celula(tab.cell(0, j), [(g["eixo"], 10.5, navy, "b")], RA_CLARO)
                for o in g["objetos"]:
                    celula(tab.cell(1, j), [(o, 9.5, cinza, "n")], branco)
                    j += 1
            if ultimo:
                tab.cell(0, n_col - 1).merge(tab.cell(1, n_col - 1))
                celula(tab.cell(0, n_col - 1), [("Total", 12, branco, "b")], navy)

            for r, l in enumerate(q["linhas"], 2):
                celula(tab.cell(r, 0), [(l["status"], fs_st, navy, "n")], branco,
                       PP_ALIGN.LEFT)
                for j, i in enumerate(ix, 1):
                    celula(tab.cell(r, j), conteudo(l["celulas"][i], fs_v, RA_AZUL), branco)
                if ultimo:
                    celula(tab.cell(r, n_col - 1), conteudo(l["total"], fs_v), branco)
            ult = n_lin - 1
            celula(tab.cell(ult, 0), [("TOTAL", 12, navy, "b")], RA_CLARO, PP_ALIGN.LEFT)
            for j, i in enumerate(ix, 1):
                celula(tab.cell(ult, j), conteudo(q["total"]["celulas"][i], fs_v), RA_CLARO)
            if ultimo:
                celula(tab.cell(ult, n_col - 1),
                       conteudo(q["total"]["total"], fs_v + 1, branco, branco), navy)
            nota = texto(36, 664, 1208, 22, [(dados["nota"], 14, cinza, "n")])
            for par in nota.paragraphs:
                for r in par.runs:
                    r.font.italic = True
    prs.save(caminho)


# ---- HTML autônomo: mapa de investimento + quadro Status × Eixo ------------
# Pedido do usuário em 12/11/2026 (botão GERAR HTML da janela). Mesmas faixas
# do mapa do MONITORAMENTO (mesma base): até R$ 30 Mi, R$ 30–60 Mi e acima.
# "Gerar PDF" do HTML abre o detalhamento destes primeiros do ranking
RA_PDF_DETALHADOS = 4
RA_FAIXAS_MAPA = [(60e6, None, "#D32F45"), (30e6, 60e6, "#F0834E"),
                  (0, 30e6, "#2E75B6")]
RA_COR_SEM_ACAO = "#DCE1E8"


def _esc(t):
    return (str(t if t is not None else "").replace("&", "&amp;")
            .replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def _rs_cheio(v):
    return ("R$ " + f"{v:,.2f}").replace(",", "X").replace(".", ",").replace("X", ".")


def _cor_faixa(v):
    for minimo, maximo, cor in RA_FAIXAS_MAPA:
        if v > 0 and v >= minimo and (maximo is None or v < maximo):
            return cor
    return RA_COR_SEM_ACAO


def _investimento_por_municipio(caminho_base, dados):
    """{chave normalizada do geojson: {nome, valor, qtd, orc, eixos}} do
    recorte atual, e a lista dos que não casaram com o geojson."""
    import relatorio_monitoramento as mon
    base = _ra_base(caminho_base)
    if dados.get("entregue"):
        base, _ = _ra_juntar_entregues(base)
    df = base.iloc[dados["ids"]]
    if "Município" not in df.columns:
        return {}, []
    chave = df["Município"].map(mon.normalizar).map(
        lambda k: mon.APELIDOS_MUNICIPIO.get(k, k))
    saida = {}
    for k, parte in df.groupby(chave):
        if not k:
            continue
        por_eixo = parte.groupby("_eixo")["_valor"].sum().sort_values(ascending=False)
        saida[k] = {"valor": float(parte["_valor"].sum()), "qtd": int(len(parte)),
                    "orc": bool(parte["_orc"].any()),
                    "nome": parte["Município"].astype(str).str.strip().mode().iloc[0],
                    "eixos": [(e, float(v)) for e, v in por_eixo.items() if v > 0][:4],
                    "acoes": _acoes_do_municipio(parte)}
    return saida, []


def _acoes_do_municipio(parte):
    """Detalhamento do município no ranking, em 2 níveis (pedido do usuário
    em 13/11/2026): Município > Status > Objeto. Um grupo por status (na
    ordem dos status, 01, 02...), com investimento e quantidade somados;
    dentro dele, uma linha por objeto, do maior investimento para o menor."""
    objeto = (parte["_f_objeto"] if "_f_objeto" in parte.columns
              else pd.Series(RA_SEM_INFO, index=parte.index))
    tab = pd.DataFrame({"o": objeto, "s": parte["_f_status"], "v": parte["_valor"],
                        "orc": parte["_orc"]})
    saida = []
    for st in sorted(tab["s"].unique(), key=_ra_ordem):
        p = tab[tab["s"] == st]
        objetos = (p.groupby("o").agg(v=("v", "sum"), n=("v", "size"), orc=("orc", "any"))
                   .reset_index().sort_values(["v", "n"], ascending=False))
        saida.append({
            "status": re.sub(r"^\d+\s*-\s*", "", st),
            "valor": _ra_rs(float(p["v"].sum())) + ("*" if p["orc"].any() else ""),
            "qtd": _ra_qtd(len(p)),
            "objetos": [[g.o, _ra_qtd(int(g.n)) if g.n > 1 else "",
                         _ra_rs(float(g.v)) + ("*" if g.orc else "")]
                        for g in objetos.itertuples()]})
    return saida


def _svg_mapa(por_mun, largura=640):
    """Bahia em SVG (paths já projetados), cada município com os dados do
    tooltip em data-*. Devolve (svg, nomes do recorte fora do geojson)."""
    import relatorio_monitoramento as mon
    geo = pac.carregar_geojson_municipios() if pac else None
    if not geo:
        return '<div class="sem-mapa">Mapa indisponível: falta municipios_bahia.geojson</div>', []
    feats = geo["features"]
    aneis = lambda f: mon._poligonos(f)   # noqa: E731
    pts = [p for f in feats for pol in aneis(f) for anel in pol for p in anel]
    minlon, maxlon = min(p[0] for p in pts), max(p[0] for p in pts)
    minlat, maxlat = min(p[1] for p in pts), max(p[1] for p in pts)
    cx = math.cos(math.radians((minlat + maxlat) / 2))
    esc_x = largura / ((maxlon - minlon) * cx)
    altura = (maxlat - minlat) * esc_x
    caminhos, usados = [], set()
    for f in feats:
        nome = str(f["properties"].get("name", ""))
        k = mon.normalizar(nome)
        info = por_mun.get(k)
        usados.add(k)
        d = []
        for pol in aneis(f):
            for anel in pol:
                d.append("M" + "L".join(
                    f"{(lon - minlon) * cx * esc_x:.1f},{(maxlat - lat) * esc_x:.1f}"
                    for lon, lat in anel) + "Z")
        if info:
            eixos = "|".join(f"{_esc(e)}~{_ra_rs(v)}" for e, v in info["eixos"])
            attrs = (f' data-n="{_esc(nome)}" data-v="{_ra_rs(info["valor"])}'
                     f'{"*" if info["orc"] else ""}" data-q="{_ra_qtd(info["qtd"])}"'
                     f' data-e="{eixos}"')
            cor = _cor_faixa(info["valor"])
        else:
            attrs = f' data-n="{_esc(nome)}" data-v="" data-q="sem ações no recorte"'
            cor = RA_COR_SEM_ACAO
        caminhos.append(f'<path d="{"".join(d)}" fill="{cor}"{attrs}/>')
    fora = sorted(info["nome"] for k, info in por_mun.items() if k not in usados)
    svg = (f'<svg class="mapa" viewBox="0 0 {largura:.0f} {altura:.0f}" '
           f'role="img" aria-label="Mapa de investimento por município">'
           + "".join(caminhos) + "</svg>")
    return svg, fora


def gerar_html_resumo_acoes(caminho_base, dados):
    """Página única (abre sem internet) com cards, mapa de investimento por
    município e o quadro Status × Eixo do recorte atual."""
    por_mun, _ = _investimento_por_municipio(caminho_base, dados)
    svg, fora = _svg_mapa(por_mun)

    # legenda: só as faixas que existem no recorte, com a contagem
    valores = [i["valor"] for i in por_mun.values() if i["valor"] > 0]
    legenda = []
    for minimo, maximo, cor in reversed(RA_FAIXAS_MAPA):
        n = sum(1 for v in valores if v >= minimo and (maximo is None or v < maximo))
        if not n:
            continue
        rot = (f"até {_ra_rs(maximo)}" if minimo == 0 else
               f"acima de {_ra_rs(minimo)}" if maximo is None else
               f"{_ra_rs(minimo)} a {_ra_rs(maximo)}")
        legenda.append(f'<li><i style="background:{cor}"></i>{rot}'
                       f'<b>{n} município{"s" if n > 1 else ""}</b></li>')
    sem_valor = sum(1 for i in por_mun.values() if i["valor"] <= 0)
    legenda.append(f'<li><i style="background:{RA_COR_SEM_ACAO}"></i>sem ação'
                   + (f" ou sem valor ({sem_valor} com ação e R$ 0)" if sem_valor else "")
                   + "</li>")

    # Lista = os 30 maiores municípios por POPULAÇÃO (RA_TOP30, já em ordem
    # de população), não mais os 15 maiores investimentos — pedido do usuário
    # em 13/11/2026. Município sem ação no recorte aparece zerado, sem "+".
    por_chave = {_chave_municipio(i["nome"]): i for i in por_mun.values()}
    ranking = [por_chave.get(_chave_municipio(nome))
               or {"nome": nome, "valor": 0.0, "qtd": 0, "orc": False, "acoes": []}
               for nome in RA_TOP30]
    valores_rank = sorted((m["valor"] for m in ranking), reverse=True)
    vmax = valores_rank[0] if valores_rank and valores_rank[0] > 0 else 1
    # um município muito acima dos demais (Salvador) achataria todas as
    # outras barras: a escala passa a ser a do 2º maior e a barra do maior
    # leva uma marca de quebra
    quebra = len(valores_rank) > 1 and valores_rank[1] > 0 and vmax > 3 * valores_rank[1]
    escala = valores_rank[1] * 1.15 if quebra else vmax
    # detalhamento de cada município abre na própria lista pelo "+" (mesmo
    # botão das seções do HTML do PAC), sem caixa de diálogo
    # Município > Status > Objeto: cada status é um subtítulo com o seu
    # próprio "+"; tudo fica oculto enquanto o município estiver recolhido
    def detalhe_rank(i, m):
        blocos = []
        for j, g in enumerate(m["acoes"]):
            tabela = ('<table><thead><tr><th>Objeto</th><th class="r">Investimento</th>'
                      '</tr></thead><tbody>'
                      + "".join(f'<tr><td>{_esc(o)}{f"<small>{_esc(q)}</small>" if q else ""}</td>'
                                f'<td class="r">{_esc(v)}</td></tr>' for o, q, v in g["objetos"])
                      + "</tbody></table>")
            blocos.append(
                f'<div class="rk-st"><div class="rk-st-linha"><button type="button" class="det det-st" '
                f'aria-expanded="false" aria-controls="rk-det-{i}-{j}" '
                f'title="Ver objetos — {_esc(g["status"])}">+</button>'
                f'<span class="rk-st-nome">{_esc(g["status"])}<small>{_esc(g["qtd"])}</small></span>'
                f'<b>{_esc(g["valor"])}</b></div>'
                f'<div class="rk-st-obj oculto" id="rk-det-{i}-{j}">{tabela}</div></div>')
        return "".join(blocos)

    def linha_rank(i, m):
        largura = (min(100, max(2, 100 * m["valor"] / escala)) if m["valor"] > 0 else 0)
        botao = (f'<button type="button" class="det" aria-expanded="false" aria-controls="rk-det-{i}"'
                 f' title="Ver detalhamento de {_esc(m["nome"])}">+</button>' if m["acoes"]
                 else '<span></span>')
        return (f'<li><div class="rk-linha{"" if m["qtd"] else " vazio"}"><span class="pos">{i + 1}</span>'
                f'<span class="nm">{_esc(m["nome"])}'
                f'<small>{_ra_qtd(m["qtd"]) if m["qtd"] else "sem ação no recorte"}</small></span>'
                f'<span class="barra{" quebra" if quebra and m["valor"] == vmax else ""}"><span style="width:'
                f'{largura:.1f}%;background:{_cor_faixa(m["valor"])}"></span></span>'
                f'<b>{_ra_rs(m["valor"]) if m["qtd"] else "–"}{"*" if m["orc"] else ""}</b>{botao}</div>'
                + (f'<div class="rk-det oculto" id="rk-det-{i}">{detalhe_rank(i, m)}</div>' if m["acoes"] else "")
                + "</li>")

    linhas_rank = "".join(linha_rank(i, m) for i, m in enumerate(ranking))
    nota_rank = ('<div class="nota-mapa">Em ordem de população (IBGE, estimativa 2026).'
                 + (" Barra do maior investimento interrompida (≈) para não achatar as demais."
                    if quebra else "") + "</div>")

    cards = "".join(
        f'<div><div class="card" style="--c:#{c["cor"]}"><span>{_esc(c["rotulo"])}</span>'
        f'<b>{_esc(c["txt"])}</b><small>{_esc(c["qtd"])}</small></div>'
        + (f'<div class="card-nota">{_esc(c["nota"])}</div>' if c.get("nota") else "")
        + "</div>"
        for c in dados["cards"])
    filtros_html = (f'<b>Recorte filtrado.</b> Filtros: {_esc(dados.get("filtros_txt") or "")}. '
                    if dados["filtrado"] else "<b>Base completa.</b> ")

    def cel(c, cls=""):
        cl = f' class="{cls}"' if cls else ""
        return (f"<td{cl}><b>{_esc(c['txt'])}</b><small>{_esc(c['qtd'])}</small></td>"
                if c else f'<td{cl}><span class="tr">–</span></td>')

    def linha(l, cls=""):
        return (f"<td>{_esc(l['status'])}</td>" + "".join(cel(c, cls) for c in l["celulas"])
                + cel(l["total"], "tot"))

    quadro = ('<table class="quadro"><thead><tr><th>Status</th>'
              + "".join(f"<th>{_esc(e)}</th>" for e in dados["eixos"])
              + '<th class="tot">Total</th></tr></thead><tbody>'
              + "".join(f"<tr>{linha(l)}</tr>" for l in dados["linhas"])
              + f'</tbody><tfoot><tr>{linha(dados["total"], "rod")}</tr></tfoot></table>')

    q = dados.get("quadro_objeto") or {}
    quadro_obj = ""
    if q.get("grupos"):
        inicio, n = set(), 0
        for g in q["grupos"]:
            inicio.add(n)
            n += len(g["objetos"])

        def linha_obj(l, cls=""):
            return (f"<td>{_esc(l['status'])}</td>"
                    + "".join(cel(c, (cls + (" ig" if i in inicio else "")).strip())
                              for i, c in enumerate(l["celulas"]))
                    + cel(l["total"], "tot ig"))

        quadro_obj = (
            '<table class="quadro quadro-obj"><thead><tr><th rowspan="2">Status</th>'
            + "".join(f'<th class="grupo ig" colspan="{len(g["objetos"])}">{_esc(g["eixo"])}</th>'
                      for g in q["grupos"])
            + '<th class="tot ig" rowspan="2">Total</th></tr><tr>'
            + "".join(f'<th class="obj{" ig" if j == 0 else ""}">{_esc(o)}</th>'
                      for g in q["grupos"] for j, o in enumerate(g["objetos"]))
            + "</tr></thead><tbody>"
            + "".join(f"<tr>{linha_obj(l)}</tr>" for l in q["linhas"])
            + f'</tbody><tfoot><tr>{linha_obj(q["total"], "rod")}</tr></tfoot></table>')

    total_mun = sum(1 for i in por_mun.values())
    nota_mapa = (f"{total_mun} municípios com ação no recorte."
                 + (f" Fora do mapa (nome não casa com o IBGE): {', '.join(fora)}." if fora else ""))
    agora = datetime.now()
    titulo = "AÇÕES EM MONITORAMENTO"
    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ações em Monitoramento</title>
<style>
:root{{--navy:#0B1F3F;--azul:#1F5C85;--cinza:#5B6576;--linha:#D9DFE8;--claro:#EEF3F8;--fundo:#F4F6F9}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--fundo);color:#1A2433;font-family:"Segoe UI",Roboto,system-ui,sans-serif}}
header{{background:var(--navy);color:#fff;padding:18px 36px;display:flex;justify-content:space-between;align-items:center;gap:20px}}
header h1{{margin:0;font-size:24px;font-weight:900;letter-spacing:.3px}}
header p{{margin:4px 0 0;color:#C9D6EA;font-size:12px;font-weight:600;letter-spacing:.4px}}
header .gov{{text-align:right;color:#4FA3E0;font-weight:700;font-size:13px;line-height:1.4}}
main{{max-width:1500px;margin:0 auto;padding:20px 36px 36px}}
.filtros{{font-size:13px;color:var(--cinza);margin-bottom:14px}}
.filtros b{{color:var(--navy)}}
.card-nota{{font-size:12px;font-style:italic;color:var(--cinza);margin-top:6px;padding-left:4px}}
.aviso-cards{{display:inline-block;background:#FFF4D6;border:1px solid #F2B632;color:#6B4E00;border-radius:8px;padding:6px 12px;font-size:13px;font-weight:600;margin-bottom:10px}}
.cards{{display:grid;grid-template-columns:repeat({len(dados["cards"])},minmax(0,440px));gap:16px;margin-bottom:22px}}
.card{{background:var(--claro);border-radius:10px;padding:14px 20px;border-left:6px solid var(--c)}}
.card span{{display:block;font-size:13px;font-weight:700;color:var(--cinza)}}
.card b{{display:block;font-size:34px;font-weight:900;color:var(--azul);margin:4px 0;white-space:nowrap}}
.card small{{font-size:17px;font-weight:600;color:var(--cinza)}}
section{{background:#fff;border-radius:12px;box-shadow:0 4px 24px rgba(11,31,63,.08);padding:20px 24px;margin-bottom:22px}}
h2{{margin:0 0 14px;font-size:15px;letter-spacing:.6px;color:var(--navy)}}
.mapa-grade{{display:grid;grid-template-columns:minmax(0,1fr) minmax(320px,1.2fr);gap:28px;align-items:start}}
.mapa{{width:100%;max-width:560px;height:auto;display:block;margin:0 auto}}
.rk-linha.vazio{{opacity:.55}}
.mapa path{{stroke:#fff;stroke-width:.6;transition:opacity .1s}}
.mapa path:hover{{stroke:var(--navy);stroke-width:1.6;opacity:.85}}
.legenda{{list-style:none;padding:0;margin:0 0 18px;font-size:13px}}
.legenda li{{display:flex;align-items:center;gap:8px;margin:5px 0}}
.legenda i{{width:16px;height:16px;border-radius:4px;flex:none}}
.legenda b{{margin-left:auto;color:var(--cinza);font-weight:600}}
h3{{font-size:13px;margin:0 0 8px;color:var(--navy);letter-spacing:.4px}}
.ranking{{list-style:none;padding:0;margin:0;font-size:14.5px}}
.ranking li{{border-bottom:1px solid #EEF1F5}}
.ranking .rk-linha{{display:grid;grid-template-columns:24px minmax(0,1.3fr) minmax(0,1fr) 108px 28px;align-items:center;gap:8px;padding:4px 0}}
.ranking .pos{{color:var(--cinza);text-align:right}}
.ranking .nm{{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.ranking .nm small{{color:var(--cinza);margin-left:6px}}
.ranking .barra{{height:10px;background:#EEF1F5;border-radius:5px;overflow:hidden}}
.ranking .barra span{{display:block;height:100%;border-radius:5px}}
.ranking .barra{{position:relative}}
.ranking .barra.quebra::after{{content:"";position:absolute;top:-2px;bottom:-2px;right:14%;width:6px;background:#fff;transform:skewX(-20deg)}}
.ranking b{{text-align:right;color:var(--navy)}}
.nota-mapa{{font-size:12px;color:var(--cinza);margin-top:10px}}
.ranking .det{{width:26px;height:26px;border-radius:50%;border:1px solid #D8DCE3;background:#F2F4F7;color:var(--navy);font-size:16px;font-weight:700;line-height:1;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;padding:0;font-family:inherit}}
.ranking .det:hover{{background:#E6E9EE}}
.ranking .det:focus-visible{{outline:2px solid var(--azul);outline-offset:1px}}
.oculto{{display:none}}
.rk-det{{margin:2px 0 10px 30px;background:#F8FAFC;border-left:3px solid var(--azul);border-radius:0 6px 6px 0;padding:4px 10px 6px;max-height:420px;overflow:auto;font-size:15px}}
.rk-st{{border-bottom:1px solid #E8ECF2}}
.rk-st:last-child{{border-bottom:none}}
.rk-st-linha{{display:grid;grid-template-columns:24px minmax(0,1fr) auto;align-items:center;gap:8px;padding:5px 0}}
.rk-st-nome{{font-weight:700;color:var(--navy)}}
.rk-st-nome small{{font-weight:400;color:var(--cinza);margin-left:8px;font-size:13px}}
.rk-st-linha b{{color:var(--navy);white-space:nowrap}}
.ranking .det.det-st{{width:22px;height:22px;font-size:14px}}
.rk-st-obj{{margin:0 0 8px 32px}}
.rk-det table{{width:100%;border-collapse:collapse;font-size:15px}}
.rk-det th{{text-align:left;color:var(--navy);padding:6px 6px 4px;border-bottom:2px solid var(--navy);font-size:14px}}
.rk-det td{{padding:5px 6px;border-bottom:1px solid #E8ECF2;vertical-align:top}}
.rk-det td small{{display:block;color:var(--cinza);font-size:12.5px}}
.rk-det .r{{text-align:right;white-space:nowrap}}
.quadro{{width:100%;border-collapse:collapse;font-size:12px}}
.quadro th,.quadro td{{border:1px solid var(--linha);padding:8px 4px;text-align:center}}
.quadro th{{color:var(--navy);font-weight:600}}
.quadro th:first-child,.quadro td:first-child{{text-align:left;padding-left:8px;min-width:150px;color:var(--navy)}}
.quadro th.tot{{background:var(--navy);color:#fff;font-weight:700}}
.quadro td b{{display:block;color:var(--azul);font-size:13px;white-space:nowrap}}
.quadro td small{{color:var(--cinza);font-size:11px;white-space:nowrap}}
.quadro td.tot b{{color:var(--navy)}}
.quadro tfoot td{{background:var(--claro);font-weight:700}}
.quadro tfoot td b{{color:var(--navy)}}
.quadro tfoot td.tot{{background:var(--navy)}}
.quadro tfoot td.tot b,.quadro tfoot td.tot small{{color:#fff}}
.tr{{color:#AAB2BF}}
h2 small{{font-weight:400;font-size:12px;color:var(--cinza);letter-spacing:0;margin-left:8px}}
.quadro-obj th.grupo{{background:var(--claro);color:var(--navy);font-weight:700}}
.quadro-obj th.obj{{font-size:11.5px;min-width:96px;color:var(--cinza)}}
.quadro-obj .ig{{border-left:2px solid var(--azul)}}
.quadro-obj td:first-child,.quadro-obj th:first-child{{position:sticky;left:0;background:#fff;z-index:1}}
.quadro-obj tfoot td:first-child{{background:var(--claro)}}
.quadro-obj thead tr+tr th:first-child{{position:static;text-align:center;min-width:96px;padding-left:4px;color:var(--cinza);background:transparent}}
@media print{{.rolagem-obj{{overflow:visible}}}}
.rolagem{{overflow-x:auto}}
.nota{{font-size:12.5px;font-style:italic;color:var(--cinza);margin-top:12px}}
footer{{text-align:center;font-size:11.5px;color:var(--cinza);padding:0 0 24px}}
#dica{{position:fixed;pointer-events:none;background:var(--navy);color:#fff;border-radius:8px;padding:9px 12px;font-size:12.5px;box-shadow:0 6px 20px rgba(0,0,0,.25);display:none;max-width:280px;z-index:5}}
#dica b{{display:block;font-size:13.5px;margin-bottom:2px}}
#dica .v{{color:#9FD3FF;font-weight:700}}
#dica ul{{list-style:none;padding:6px 0 0;margin:6px 0 0;border-top:1px solid rgba(255,255,255,.2)}}
#dica li{{display:flex;justify-content:space-between;gap:12px;color:#C9D6EA}}
@media (max-width:900px){{.mapa-grade{{grid-template-columns:1fr}}.cards{{grid-template-columns:1fr 1fr}}main,header{{padding-left:16px;padding-right:16px}}}}
@media print{{body{{background:#fff}}section{{box-shadow:none;break-inside:avoid}}#dica{{display:none!important}}}}
.btn-pdf{{position:fixed;top:18px;right:24px;z-index:20;padding:11px 22px;font:700 14px "Segoe UI",Roboto,system-ui,sans-serif;color:#fff;background:#0B5CAD;border:none;border-radius:8px;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.3)}}
.btn-pdf:hover{{background:#0A4E92}}
@media print{{.btn-pdf,.ranking .det{{display:none!important}}.rk-det{{max-height:none;overflow:visible}}section.sec-mapa{{break-inside:auto}}.rk-det tr{{break-inside:avoid}}
.mapa-grade{{grid-template-columns:minmax(0,1fr) minmax(0,1.15fr);gap:18px}}.mapa{{max-width:330px;margin:0}}.mapa-grade>div:first-child{{break-inside:avoid}}
.ranking .rk-linha{{grid-template-columns:18px minmax(0,1.9fr) minmax(0,.7fr) 96px;gap:6px}}.ranking .nm small{{margin-left:4px}}}}
</style></head><body>
<button type="button" class="btn-pdf" onclick="gerarPdf()" title="Imprime/salva em PDF com o detalhamento dos {RA_PDF_DETALHADOS} maiores municípios aberto">Gerar PDF</button>
<header><div><h1>{titulo}</h1><p>COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS ESTRATÉGICAS</p></div>
<div class="gov">GOVERNO DO ESTADO DA BAHIA<br>CASA CIVIL</div></header>
<main>
<div class="filtros">{filtros_html}Gerado em {agora:%d/%m/%Y às %Hh%M}.</div>
{f'<div class="aviso-cards">{_esc(dados.get("aviso_cards"))}</div>' if dados.get("aviso_cards") else ""}
<div class="cards">{cards}</div>
<section class="sec-mapa"><h2>MAPA DE INVESTIMENTO POR MUNICÍPIO</h2>
<div class="mapa-grade"><div>{svg}<div class="nota-mapa">{_esc(nota_mapa)} Passe o mouse sobre um município para ver o detalhe.</div></div>
<div><h3>FAIXAS DE INVESTIMENTO</h3><ul class="legenda">{"".join(legenda)}</ul>
<h3>30 MAIORES MUNICÍPIOS (POPULAÇÃO)</h3><ol class="ranking">{linhas_rank}</ol>{nota_rank}</div></div></section>
<section><h2>QUADRO RESUMO — STATUS × EIXO</h2><div class="rolagem">{quadro}</div>
<div class="nota">{_esc(dados["nota"])}</div></section>
{f'<section><h2>QUADRO RESUMO — STATUS × OBJETO <small>objetos agrupados pelo eixo, na mesma ordem do quadro acima</small></h2><div class="rolagem rolagem-obj">{quadro_obj}</div><div class="nota">{_esc(dados["nota"])}</div></section>' if quadro_obj else ""}
</main>
<footer>CGAPE · Casa Civil · Base Ações do Cronograma</footer>
<div id="dica"></div>
<script>
(function(){{
  // "+" / "–" de cada município: abre/fecha o detalhamento na própria lista
  function alternar(b,abrir){{
    var alvo=document.getElementById(b.getAttribute("aria-controls"));
    if(abrir===undefined)abrir=alvo.classList.contains("oculto");
    alvo.classList.toggle("oculto",!abrir);
    b.textContent=abrir?"–":"+";
    b.setAttribute("aria-expanded",abrir?"true":"false");
    b.title=abrir?b.title.replace("Ver","Ocultar"):b.title.replace("Ocultar","Ver");
  }}
  var botoes=Array.prototype.slice.call(document.querySelectorAll(".ranking .det"));
  botoes.forEach(function(b){{b.addEventListener("click",function(){{alternar(b);}});}});
  var botoesMun=Array.prototype.slice.call(document.querySelectorAll(".ranking .rk-linha .det"));
  // Gerar PDF (pedido do usuário em 13/11/2026): abre o detalhamento dos
  // {RA_PDF_DETALHADOS} primeiros municípios (com os status e objetos deles)
  // para a impressão e devolve a lista ao estado anterior quando a janela de
  // impressão fecha.
  window.gerarPdf=function(){{
    var antes=botoes.map(function(b){{return b.getAttribute("aria-expanded")==="true";}});
    var abrir=new Set(botoesMun.slice(0,{RA_PDF_DETALHADOS}));
    botoes.forEach(function(b){{
      var mun=b.classList.contains("det-st")?b.closest("li").querySelector(".rk-linha .det"):b;
      alternar(b,abrir.has(mun));
    }});
    function restaurar(){{
      botoes.forEach(function(b,i){{alternar(b,antes[i]);}});
      window.removeEventListener("afterprint",restaurar);
    }}
    window.addEventListener("afterprint",restaurar);
    window.print();
  }};
  var dica=document.getElementById("dica");
  document.querySelectorAll(".mapa path").forEach(function(p){{
    p.addEventListener("mousemove",function(ev){{
      var e=p.dataset.e?p.dataset.e.split("|").map(function(x){{var a=x.split("~");return "<li><span>"+a[0]+"</span><span>"+a[1]+"</span></li>";}}).join(""):"";
      dica.innerHTML="<b>"+p.dataset.n+"</b>"+(p.dataset.v?'<span class="v">'+p.dataset.v+"</span> · ":"")+p.dataset.q+(e?"<ul>"+e+"</ul>":"");
      dica.style.display="block";
      var x=ev.clientX+14,y=ev.clientY+14,r=dica.getBoundingClientRect();
      if(x+r.width>innerWidth-8)x=ev.clientX-r.width-14;
      if(y+r.height>innerHeight-8)y=ev.clientY-r.height-14;
      dica.style.left=x+"px";dica.style.top=y+"px";
    }});
    p.addEventListener("mouseleave",function(){{dica.style.display="none";}});
  }});
}})();
</script>
</body></html>"""


# ---- API (chamada pelo painel) ----------------------------------------------
def _sem_planilha():
    return f"Planilha '{RA_PLANILHA}' não encontrada na pasta do programa."


def api_dados(filtros=None):
    try:
        caminho = caminho_planilha()
        if not caminho:
            return {"ok": False, "mensagem": _sem_planilha()}
        return {"ok": True, "dados": resumo_acoes_dados(caminho, filtros)}
    except Exception as exc:
        return {"ok": False, "mensagem": f"Erro ao ler '{RA_PLANILHA}': {exc}"}


def api_base():
    try:
        caminho = caminho_planilha()
        if not caminho:
            return {"ok": False, "mensagem": _sem_planilha()}
        return {"ok": True, **resumo_acoes_tabela(caminho)}
    except Exception as exc:
        return {"ok": False, "mensagem": f"Erro ao ler '{RA_PLANILHA}': {exc}"}


def _caminho_livre(caminho):
    """Arquivo aberto (PowerPoint/Excel): grava com outro nome."""
    if os.path.exists(caminho):
        try:
            with open(caminho, "ab"):
                pass
        except OSError:
            raiz, ext = os.path.splitext(caminho)
            caminho = f"{raiz}_{datetime.now():%Hh%Mmin}{ext}"
    return caminho


def _abrir(caminho):
    try:
        if os.name == "nt":
            os.startfile(caminho)  # noqa
    except Exception:
        pass


def api_pptx(filtros=None):
    """Grava o slide em Downloads e o abre."""
    try:
        caminho_base = caminho_planilha()
        if not caminho_base:
            return {"mensagem": _sem_planilha()}
        dados = resumo_acoes_dados(caminho_base, filtros)
        if dados["vazio"]:
            return {"mensagem": "Nenhuma ação com os filtros escolhidos: "
                                "apresentação não gerada.", "dados": dados}
        sufixo = "_Filtrado" if dados["filtrado"] else ""
        nome = ("Acoes_em_Monitoramento_"
                f"{datetime.now():%d-%m-%Y}{sufixo}.pptx")
        caminho = _caminho_livre(os.path.join(pasta_downloads(), nome))
        gerar_pptx_resumo_acoes(dados, caminho)
        _abrir(caminho)
        return {"mensagem": "Apresentação salva em Downloads: "
                            + os.path.basename(caminho), "dados": dados}
    except Exception as exc:
        return {"mensagem": f"Erro ao gerar a apresentação: {exc}"}


def api_xlsx(filtros=None):
    """Grava em Downloads a planilha das ações filtradas e a abre."""
    try:
        caminho_base = caminho_planilha()
        if not caminho_base:
            return {"mensagem": _sem_planilha()}
        dados = resumo_acoes_dados(caminho_base, filtros)
        if dados["vazio"]:
            return {"mensagem": "Nenhuma ação com os filtros escolhidos: "
                                "planilha não gerada."}
        sufixo = "_Filtrado" if dados["filtrado"] else ""
        nome = (f"Acoes_em_Monitoramento_{datetime.now():%d-%m-%Y_%Hh%Mmin}"
                f"{sufixo}.xlsx")
        caminho = _caminho_livre(os.path.join(pasta_downloads(), nome))
        gerar_xlsx_resumo_acoes(caminho_base, dados, caminho)
        _abrir(caminho)
        return {"mensagem": "Planilha salva em Downloads: "
                            + os.path.basename(caminho)}
    except Exception as exc:
        return {"mensagem": f"Erro ao gerar a planilha: {exc}"}


def api_html(filtros=None, devolver_conteudo=False):
    """Página HTML (mapa + quadro) do recorte. Desktop: grava em Downloads e
    abre no navegador padrão. Web (devolver_conteudo): devolve o HTML para o
    navegador abrir numa aba."""
    try:
        caminho_base = caminho_planilha()
        if not caminho_base:
            return {"ok": False, "mensagem": _sem_planilha()}
        dados = resumo_acoes_dados(caminho_base, filtros)
        if dados["vazio"]:
            return {"ok": False, "mensagem": "Nenhuma ação com os filtros "
                                             "escolhidos: HTML não gerado."}
        html = gerar_html_resumo_acoes(caminho_base, dados)
        if devolver_conteudo:
            return {"ok": True, "html": html,
                    "mensagem": "HTML aberto em uma nova aba."}
        sufixo = "_Filtrado" if dados["filtrado"] else ""
        nome = ("Acoes_em_Monitoramento_"
                f"{datetime.now():%d-%m-%Y_%Hh%Mmin}{sufixo}.html")
        caminho = _caminho_livre(os.path.join(pasta_downloads(), nome))
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(html)
        _abrir(caminho)
        return {"ok": True, "mensagem": "HTML salvo em Downloads: "
                                        + os.path.basename(caminho)}
    except Exception as exc:
        return {"ok": False, "mensagem": f"Erro ao gerar o HTML: {exc}"}


# ---- HTML / CSS / JS da janela (injetado no painel antes do </body>) --------
HTML_JANELA = r"""
<div id="ra-overlay" aria-hidden="true">
  <div id="ra-janela" role="dialog" aria-label="Ações em monitoramento">
    <div class="ra-cabecalho">
      <h2>AÇÕES EM MONITORAMENTO
        <small>base Ações do Cronograma</small></h2>
      <div class="ra-botoes">
        <button type="button" id="raHtml" title="Gera uma página HTML com o mapa de investimento por município e o quadro Status × Eixo do recorte atual">GERAR HTML</button>
        <button type="button" id="raGerar" title="Lê de novo a planilha Ações do Cronograma e gera o slide em Downloads">GERAR APRESENTAÇÃO</button>
        <button type="button" class="ra-fechar" id="raFechar">FECHAR</button>
      </div>
    </div>
    <div class="ra-corpo">
      <div class="ra-filtros" id="raFiltros"></div>
      <div class="ra-msg" id="raMsg"></div>
      <div id="raQuadro"></div>
      <div id="raTabela"></div>
    </div>
  </div>
</div>
<style>
  #ra-overlay { display:none; position:fixed; inset:0; z-index:200;
    background:var(--cor-fundo); color:var(--cor-texto-primario);
    font-family:var(--fonte); }
  #ra-overlay.aberto { display:flex; }
  #ra-janela { display:flex; flex-direction:column; width:100%; height:100%; }
  .ra-cabecalho { display:flex; justify-content:space-between; align-items:center;
    gap:12px; padding:10px 18px; border-bottom:2px solid var(--cor-acento-teal);
    background:var(--cor-card); }
  .ra-cabecalho h2 { margin:0; font-size:16px; letter-spacing:.5px; }
  .ra-cabecalho h2 small { font-weight:400; font-size:12px; margin-left:8px;
    color:var(--cor-texto-terciario); letter-spacing:0; }
  .ra-botoes { display:flex; gap:8px; align-items:center; }
  .ra-botoes button, #raLimpar, .ra-mais { background:transparent;
    border:1px solid var(--cor-texto-terciario); color:var(--cor-texto-primario);
    border-radius:6px; padding:6px 14px; font-size:12px; cursor:pointer;
    font-family:inherit; }
  .ra-botoes button:hover, #raLimpar:hover, .ra-mais:hover {
    border-color:var(--cor-acento-teal); color:var(--cor-acento-teal-texto); }
  #raGerar, #raXlsx { background:var(--cor-acento-teal); color:#10251F;
    border-color:var(--cor-acento-teal); font-weight:700; }
  #raGerar:hover, #raXlsx:hover { background:var(--cor-acento-teal-hover);
    color:#10251F; }
  #raGerar:disabled, #raXlsx:disabled, #raHtml:disabled { opacity:.6; cursor:default; }
  .ra-corpo { flex:1; overflow:auto; padding:14px 18px 40px; }
  .ra-msg { font-size:13px; color:var(--cor-texto-secundario); min-height:18px;
    margin-bottom:10px; }
  .ra-cards { display:grid; gap:12px; margin-bottom:14px; }
  .ra-aviso-cards { display:inline-block; margin-bottom:8px; padding:5px 10px; border-radius:6px;
    font-size:12px; font-weight:600; color:#F2B632; background:rgba(242,182,50,.12);
    border:1px solid rgba(242,182,50,.5); }
  .ra-card { background:var(--cor-card); border-radius:8px; padding:12px 16px;
    border-left:5px solid var(--ra-cor, var(--cor-acento-teal));
    box-shadow:var(--sombra-card); }
  .ra-card span { display:block; font-size:12px; color:var(--cor-texto-secundario); }
  .ra-card span + b + span { font-size:15px; font-weight:600; }
  .ra-card b { display:block; font-size:28px; color:var(--cor-acento-mint-texto);
    margin:4px 0; white-space:nowrap; }
  .ra-rolagem { overflow-x:auto; }
  .ra-tab { width:100%; border-collapse:collapse; font-size:12px;
    background:var(--cor-card); }
  .ra-tab th, .ra-tab td { border:1px solid var(--cor-card-elevado);
    padding:9px 6px; text-align:center; }
  .ra-tab th { color:var(--cor-texto-primario); }
  .ra-tab th:first-child, .ra-tab td:first-child { text-align:left; min-width:130px; }
  .ra-tab tfoot td, .ra-tab .tot { background:var(--cor-card-elevado); }
  .ra-tab b { display:block; font-size:13px; white-space:nowrap; }
  .ra-tab small { color:var(--cor-texto-secundario); }
  .ra-nota { font-size:12px; font-style:italic; color:var(--cor-texto-secundario);
    margin-top:10px; }
  .ra-quadro-titulo { margin:22px 0 8px; font-size:13px; font-weight:700;
    letter-spacing:.5px; color:var(--cor-texto-primario); }
  .ra-quadro-titulo small { font-weight:400; font-size:12px; letter-spacing:0;
    color:var(--cor-texto-terciario); }
  .ra-tab-obj th.ra-grupo { background:var(--cor-card-elevado);
    color:var(--cor-acento-mint-texto); font-size:12px; letter-spacing:.3px; }
  .ra-tab-obj th.ra-obj { font-weight:600; font-size:11.5px; min-width:92px;
    color:var(--cor-texto-secundario); }
  .ra-tab-obj .ra-ini-grupo { border-left:2px solid var(--cor-acento-teal); }
  .ra-tab-obj td:first-child, .ra-tab-obj th:first-child { position:sticky; left:0;
    z-index:1; background:var(--cor-card); }
  .ra-tab-obj tfoot td:first-child { background:var(--cor-card-elevado); }
  .ra-tab-obj thead tr + tr th:first-child { position:static; text-align:center;
    min-width:92px; background:transparent; }
  .ra-vazio { padding:40px 0; text-align:center; font-size:14px;
    color:var(--cor-texto-secundario); }
  .ra-filtros { background:var(--cor-card); border-radius:8px;
    padding:10px 14px 12px; margin-bottom:10px; box-shadow:var(--sombra-card); }
  .ra-filtros:empty { display:none; }
  .ra-filtros-topo { display:flex; justify-content:space-between; align-items:center;
    min-height:26px; margin-bottom:6px; font-size:12px; letter-spacing:.5px;
    color:var(--cor-texto-secundario); }
  .ra-filtros-topo b { flex:1; }
  .ra-top30 { display:inline-flex; align-items:center; gap:8px; margin-right:14px;
    cursor:pointer; font-size:12px; font-weight:600; color:var(--cor-texto-secundario);
    user-select:none; }
  .ra-top30 small { font-weight:400; color:var(--cor-texto-terciario); }
  .ra-top30 input { position:absolute; opacity:0; pointer-events:none; }
  .ra-top30-trilho { position:relative; width:32px; height:18px; border-radius:9px;
    background:var(--cor-card-elevado); border:1px solid var(--cor-texto-terciario);
    transition:background .15s ease; flex:none; }
  .ra-top30-trilho span { position:absolute; top:2px; left:2px; width:12px; height:12px;
    border-radius:50%; background:var(--cor-texto-secundario); transition:transform .15s ease; }
  .ra-top30.ativo { color:var(--cor-acento-teal-texto); }
  .ra-top30.ativo .ra-top30-trilho { background:var(--cor-acento-teal);
    border-color:var(--cor-acento-teal); }
  .ra-top30.ativo .ra-top30-trilho span { transform:translateX(14px); background:#10251F; }
  .ra-top30 input:focus-visible + .ra-top30-trilho { outline:2px solid var(--cor-acento-teal);
    outline-offset:2px; }
  .ra-filtros-grade { display:grid; align-items:start; gap:10px 12px;
    grid-template-columns:repeat(8, minmax(0, 1fr)); }
  @media (max-width:1100px) {
    .ra-filtros-grade { grid-template-columns:repeat(4, minmax(0, 1fr)); } }
  .ra-ms { position:relative; min-width:0; }
  .ra-ms h3 { margin:0 0 4px; font-size:11px; font-weight:600;
    color:var(--cor-texto-secundario); text-transform:uppercase; letter-spacing:.3px; }
  .ra-ms-botao { width:100%; text-align:left; background:var(--cor-fundo);
    color:var(--cor-texto-primario); border:1px solid var(--cor-card-elevado);
    border-radius:6px; padding:6px 22px 6px 10px; font-size:12px; cursor:pointer;
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis; position:relative;
    font-family:inherit; }
  .ra-ms-botao::after { content:"\25BE"; position:absolute; right:8px; top:5px;
    color:var(--cor-texto-terciario); }
  .ra-ms.marcado .ra-ms-botao { border-color:var(--cor-acento-teal); }
  .ra-ms-lista { display:none; position:absolute; z-index:5; top:100%; left:0;
    min-width:100%; width:max-content; max-width:340px; margin-top:4px;
    background:var(--cor-card); border:1px solid var(--cor-card-elevado);
    border-radius:8px; box-shadow:var(--sombra-card); padding:6px; }
  .ra-ms.aberto .ra-ms-lista { display:block; }
  .ra-ms-lista input[type=text] { width:100%; margin-bottom:6px;
    background:var(--cor-fundo); color:var(--cor-texto-primario);
    border:1px solid var(--cor-card-elevado); border-radius:6px; padding:4px 8px;
    font-size:12px; font-family:inherit; }
  .ra-ms-opcoes { max-height:260px; overflow-y:auto; }
  .ra-ms-opcoes label { display:flex; align-items:center; gap:6px; padding:3px 4px;
    font-size:12px; cursor:pointer; border-radius:4px; }
  .ra-ms-opcoes label:hover { background:var(--cor-card-elevado); }
  .ra-ms-opcoes label span { flex:1; }
  .ra-ms-opcoes label small { color:var(--cor-texto-terciario); }
  .ra-ms-opcoes label.zerado { opacity:.5; }
  .ra-chips { display:flex; flex-wrap:wrap; gap:4px; margin-top:4px;
    max-height:78px; overflow-y:auto; }
  .ra-chip { display:inline-flex; align-items:center; gap:2px; max-width:100%;
    background:var(--cor-card-elevado); border-radius:10px; padding:1px 4px 1px 8px;
    font-size:11px; }
  .ra-chip span { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .ra-chip button { background:none; border:none; color:var(--cor-texto-secundario);
    cursor:pointer; font-size:13px; line-height:1; padding:0 2px; }
  .ra-datas-topo { font-size:12px; letter-spacing:.5px; color:var(--cor-texto-secundario);
    font-weight:700; margin:14px 0 6px; padding-top:10px;
    border-top:1px solid var(--cor-card-elevado); }
  .ra-datas-topo small { font-weight:400; letter-spacing:0; }
  .ra-datas-grade { display:grid; align-items:start; gap:10px 12px;
    grid-template-columns:repeat(5, minmax(0, 1fr)); }
  @media (max-width:1100px) {
    .ra-datas-grade { grid-template-columns:repeat(3, minmax(0, 1fr)); } }
  .ra-data { min-width:0; border:1px solid var(--cor-card-elevado);
    border-radius:8px; padding:6px 8px 8px; }
  .ra-data.ativo { border-color:var(--cor-acento-teal); }
  .ra-data h3 { display:flex; align-items:baseline; gap:6px; margin:0 0 6px;
    font-size:11px; color:var(--cor-texto-secundario); text-transform:uppercase;
    letter-spacing:.3px; }
  .ra-data h3 span { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis;
    white-space:nowrap; }
  .ra-data h3 small { font-weight:400; text-transform:none; letter-spacing:0;
    white-space:nowrap; color:var(--cor-texto-terciario); }
  .ra-data-x { background:none; border:none; color:var(--cor-texto-secundario);
    cursor:pointer; font-size:14px; line-height:1; padding:0 2px; visibility:hidden; }
  .ra-data.ativo .ra-data-x { visibility:visible; }
  .ra-mandatos { display:flex; gap:6px; margin-bottom:6px; }
  .ra-mandatos button { flex:1; background:var(--cor-fundo); color:var(--cor-texto-secundario);
    border:1px solid var(--cor-card-elevado); border-radius:6px; padding:4px 6px; font-size:11px;
    font-weight:700; letter-spacing:.3px; font-family:inherit; cursor:pointer; }
  .ra-mandatos button:hover { color:var(--cor-texto-primario); border-color:var(--cor-acento-teal); }
  .ra-mandatos button.ativo { background:var(--cor-acento-teal); border-color:var(--cor-acento-teal);
    color:#10251F; }
  .ra-data-campos { display:flex; flex-direction:column; gap:4px; }
  .ra-data-campos label { display:flex; align-items:center; gap:6px; font-size:11px;
    color:var(--cor-texto-secundario); }
  .ra-data-campos input { flex:1; min-width:0; background:var(--cor-fundo);
    color:var(--cor-texto-primario); border:1px solid var(--cor-card-elevado);
    border-radius:6px; padding:3px 6px; font-size:12px; font-family:inherit; }
  html:not([data-tema="claro"]) .ra-data-campos input { color-scheme:dark; }
  #raTabela { margin-top:18px; }
  .ra-tabela-topo { display:flex; align-items:center; gap:12px; margin-bottom:8px;
    font-size:12px; color:var(--cor-texto-secundario); }
  .ra-tabela-topo b { color:var(--cor-texto-primario); letter-spacing:.5px; }
  .ra-tabela-topo span { flex:1; }
  #raXlsx { border-radius:6px; padding:6px 12px; font-size:12px; cursor:pointer;
    border:none; font-family:inherit; }
  .ra-tabela-wrap { overflow:auto; max-height:70vh; border-radius:8px;
    border:1px solid var(--cor-card-elevado); }
  .ra-acoes { width:100%; border-collapse:collapse; font-size:12px;
    background:var(--cor-card); }
  .ra-acoes th { position:sticky; top:0; background:var(--cor-card-elevado);
    padding:7px 8px; text-align:left; cursor:pointer; white-space:nowrap; z-index:1; }
  .ra-acoes td { padding:5px 8px; border-top:1px solid var(--cor-card-elevado);
    white-space:nowrap; }
  .ra-acoes td:nth-child(3) { white-space:normal; min-width:260px; }
  .ra-acoes .dir { text-align:right; }
  .ra-acoes .centro { text-align:center; }
  .ra-acoes th.ra-col-filtro { color:var(--cor-acento-teal-texto);
    box-shadow:inset 0 -2px 0 var(--cor-acento-teal); }
  .ra-mais { display:block; margin:10px auto 0; }
</style>
<script>
(function () {
  function esc(t) {
    return String(t == null ? "" : t).replace(/[&<>"']/g, function (c) {
      return {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c];
    });
  }
  function norm(t) {
    return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  }
  async function api(nome) {
    var args = Array.prototype.slice.call(arguments, 1);
    if (window.pywebview && window.pywebview.api && window.pywebview.api[nome]) {
      return await window.pywebview.api[nome].apply(null, args);
    }
    var r = await fetch("/api/" + nome, {method: "POST",
      headers: {"Content-Type": "application/json"}, body: JSON.stringify(args)});
    return await r.json();
  }

  var raFiltros = {};
  var raPedido = 0;
  var overlay = document.getElementById("ra-overlay");

  function raCelula(c, classe) {
    var cl = classe ? ' class="' + classe + '"' : "";
    return c ? "<td" + cl + "><b>" + esc(c.txt) + "</b><small>" + esc(c.qtd) + "</small></td>"
             : "<td" + cl + ">–</td>";
  }
  function raAlternar(chave, valor, marcar) {
    var atual = raFiltros[chave] || [];
    var nova = marcar ? atual.concat([valor]) : atual.filter(function (v) { return v !== valor; });
    if (nova.length) raFiltros[chave] = nova; else delete raFiltros[chave];
    raConsultar();
  }
  // "entregue" é configuração do quadro, não filtro: o LIMPAR não mexe nele
  function raChavesFiltro() {
    return Object.keys(raFiltros).filter(function (c) { return c !== "entregue"; });
  }
  function raLimparFiltros() {
    raChavesFiltro().forEach(function (c) { delete raFiltros[c]; });
    raConsultar();
  }
  function raBuscar(bloco) {
    var termo = norm(bloco.querySelector(".ra-ms-lista input").value);
    bloco.querySelectorAll(".ra-ms-opcoes label").forEach(function (l) {
      l.style.display = norm(l.querySelector("span").textContent).indexOf(termo) >= 0 ? "" : "none";
    });
  }
  function raMontarBloco(f) {
    var bloco = document.createElement("div");
    bloco.className = "ra-ms";
    bloco.dataset.ra = f.chave;
    bloco.innerHTML = "<h3>" + esc(f.rotulo) + "</h3>" +
      '<button type="button" class="ra-ms-botao">Escolher opções…</button>' +
      '<div class="ra-ms-lista"><input type="text" placeholder="Buscar…">' +
      '<div class="ra-ms-opcoes"></div></div><div class="ra-chips"></div>';
    bloco.querySelector(".ra-ms-botao").onclick = function (e) {
      e.stopPropagation();
      var abrir = !bloco.classList.contains("aberto");
      document.querySelectorAll("#ra-overlay .ra-ms.aberto").forEach(function (b) { b.classList.remove("aberto"); });
      if (abrir) { bloco.classList.add("aberto"); bloco.querySelector(".ra-ms-lista input").focus(); }
    };
    bloco.querySelector(".ra-ms-lista").onclick = function (e) { e.stopPropagation(); };
    bloco.querySelector(".ra-ms-lista input").oninput = function () { raBuscar(bloco); };
    return bloco;
  }
  function raOpcoes(bloco, f) {
    var marcados = raFiltros[f.chave] || [];
    var lista = bloco.querySelector(".ra-ms-opcoes");
    var rolagem = lista.scrollTop;
    lista.innerHTML = "";
    f.opcoes.forEach(function (o) {
      var rotulo = document.createElement("label");
      if (!o.n) rotulo.className = "zerado";
      var caixa = document.createElement("input");
      caixa.type = "checkbox";
      caixa.checked = marcados.indexOf(o.v) >= 0;
      caixa.onchange = function (e) { e.stopPropagation(); raAlternar(f.chave, o.v, caixa.checked); };
      var nome = document.createElement("span");
      nome.textContent = o.v;
      var qtd = document.createElement("small");
      qtd.textContent = o.n;
      rotulo.append(caixa, nome, qtd);
      lista.appendChild(rotulo);
    });
    lista.scrollTop = rolagem;
    raBuscar(bloco);
    bloco.classList.toggle("marcado", marcados.length > 0);
    bloco.querySelector(".ra-ms-botao").textContent = marcados.length
      ? marcados.length + " escolhido(s)" : "Escolher opções…";
    var chips = bloco.querySelector(".ra-chips");
    chips.innerHTML = "";
    marcados.forEach(function (v) {
      var chip = document.createElement("span");
      chip.className = "ra-chip";
      var nome = document.createElement("span");
      nome.textContent = v; nome.title = v;
      var x = document.createElement("button");
      x.type = "button"; x.textContent = "×"; x.title = "Remover";
      x.onclick = function (e) { e.stopPropagation(); raAlternar(f.chave, v, false); };
      chip.append(nome, x);
      chips.appendChild(chip);
    });
  }
  function raData(chave) {
    var bloco = document.querySelector('.ra-data[data-ra="' + chave + '"]');
    var de = bloco.querySelector(".de").value;
    var ate = bloco.querySelector(".ate").value;
    if (de || ate) raFiltros["d_" + chave] = {de: de, ate: ate};
    else delete raFiltros["d_" + chave];
    raConsultar();
  }
  // Atalhos de período por mandato na Entrega da Obra (pedido do usuário em
  // 13/11/2026): um clique preenche de/até; clicar no ativo limpa.
  var RA_MANDATOS = [
    {rotulo: "JERÔNIMO 01", de: "2023-01-01", ate: "2026-12-31"},
    {rotulo: "JERÔNIMO 02", de: "2027-01-01", ate: "2030-12-31"}];
  var RA_DATAS_COM_MANDATO = ["entrega"];
  function raMontarData(f) {
    var bloco = document.createElement("div");
    bloco.className = "ra-data";
    bloco.dataset.ra = f.chave;
    var mandatos = RA_DATAS_COM_MANDATO.indexOf(f.chave) !== -1;
    bloco.innerHTML = "<h3><span>" + esc(f.rotulo) + '</span><small></small>' +
      '<button type="button" class="ra-data-x" title="Limpar este filtro">×</button></h3>' +
      (mandatos ? '<div class="ra-mandatos">' + RA_MANDATOS.map(function (m, i) {
        return '<button type="button" data-i="' + i + '" title="' + m.de.split("-").reverse().join("/") +
          " a " + m.ate.split("-").reverse().join("/") + '">' + m.rotulo + "</button>";
      }).join("") + "</div>" : "") +
      '<div class="ra-data-campos"><label>de <input type="date" class="de"></label>' +
      '<label>até <input type="date" class="ate"></label></div>';
    bloco.querySelectorAll(".ra-mandatos button").forEach(function (b) {
      b.onclick = function (e) {
        e.stopPropagation();
        var m = RA_MANDATOS[+b.dataset.i], faixa = raFiltros["d_" + f.chave] || {};
        if (faixa.de === m.de && faixa.ate === m.ate) delete raFiltros["d_" + f.chave];
        else raFiltros["d_" + f.chave] = {de: m.de, ate: m.ate};
        raConsultar();
      };
    });
    bloco.querySelectorAll("input").forEach(function (i) {
      i.onchange = function (e) { e.stopPropagation(); raData(f.chave); };
    });
    bloco.querySelector(".ra-data-x").onclick = function () {
      delete raFiltros["d_" + f.chave]; raConsultar();
    };
    return bloco;
  }
  function raAtualizarData(bloco, f) {
    var faixa = raFiltros["d_" + f.chave] || {};
    var de = bloco.querySelector(".de"), ate = bloco.querySelector(".ate");
    de.value = faixa.de || ""; ate.value = faixa.ate || "";
    de.min = ate.min = f.min; de.max = ate.max = f.max;
    if (faixa.ate) de.max = faixa.ate;
    if (faixa.de) ate.min = faixa.de;
    bloco.classList.toggle("ativo", !!(faixa.de || faixa.ate));
    bloco.querySelector("small").textContent = f.n + " com data";
    bloco.querySelectorAll(".ra-mandatos button").forEach(function (b) {
      var m = RA_MANDATOS[+b.dataset.i];
      b.classList.toggle("ativo", faixa.de === m.de && faixa.ate === m.ate);
    });
  }
  function raDesenharFiltros(lista, datas) {
    var caixa = document.getElementById("raFiltros");
    datas = datas || [];
    if (!caixa.children.length) {
      caixa.innerHTML = '<div class="ra-filtros-topo"><b>FILTROS</b>' +
        '<label class="ra-top30" id="raTop30Rotulo"><input type="checkbox" id="raTop30">' +
        '<span class="ra-top30-trilho"><span></span></span>30 MAIORES MUNICÍPIOS ' +
        '<small>(população IBGE 2026)</small></label>' +
        '<label class="ra-top30" id="raEntregueRotulo" title="Obras Concluídas e Obras ' +
        'Inauguradas somadas num só status"><input type="checkbox" id="raEntregue">' +
        '<span class="ra-top30-trilho"><span></span></span>STATUS ENTREGUE ' +
        '<small>(concluídas + inauguradas)</small></label>' +
        '<button type="button" id="raLimpar">LIMPAR FILTROS</button></div>' +
        '<div class="ra-filtros-grade"></div>' +
        '<div class="ra-datas-topo">DATAS E PRAZOS <small>(com um intervalo ' +
        'preenchido, ações sem essa data ficam de fora)</small></div>' +
        '<div class="ra-datas-grade"></div>';
      caixa.querySelector("#raLimpar").onclick = raLimparFiltros;
      caixa.querySelector("#raTop30").onchange = function (e) {
        e.stopPropagation();
        if (e.target.checked) raFiltros.top30 = true; else delete raFiltros.top30;
        raConsultar();
      };
      caixa.querySelector("#raEntregue").onchange = function (e) {
        e.stopPropagation();
        if (e.target.checked) raFiltros.entregue = true; else delete raFiltros.entregue;
        raConsultar();
      };
      var grade = caixa.querySelector(".ra-filtros-grade");
      lista.forEach(function (f) { grade.appendChild(raMontarBloco(f)); });
      var gradeDatas = caixa.querySelector(".ra-datas-grade");
      datas.forEach(function (f) { gradeDatas.appendChild(raMontarData(f)); });
      if (!datas.length) caixa.querySelector(".ra-datas-topo").remove();
    }
    lista.forEach(function (f) {
      raOpcoes(caixa.querySelector('.ra-ms[data-ra="' + f.chave + '"]'), f);
    });
    datas.forEach(function (f) {
      var bloco = caixa.querySelector('.ra-data[data-ra="' + f.chave + '"]');
      if (bloco) raAtualizarData(bloco, f);
    });
    document.getElementById("raTop30").checked = !!raFiltros.top30;
    document.getElementById("raTop30Rotulo").classList.toggle("ativo", !!raFiltros.top30);
    document.getElementById("raEntregue").checked = !!raFiltros.entregue;
    document.getElementById("raEntregueRotulo").classList.toggle("ativo", !!raFiltros.entregue);
    document.getElementById("raLimpar").style.visibility =
      raChavesFiltro().length ? "visible" : "hidden";
  }

  // ---- tabela das ações do recorte ----------------------------------------
  var raBase = null, raBaseVersao = null, raBasePedido = null;
  var raIds = [];
  var raTab = {col: null, dir: 1, limite: 200};
  var RA_PASSO = 200;
  var raBRL = new Intl.NumberFormat("pt-BR", {minimumFractionDigits: 2, maximumFractionDigits: 2});
  function raTabelaAtualizar(d) {
    var alvo = document.getElementById("raTabela");
    raIds = d.vazio ? [] : d.ids;
    raTab.limite = RA_PASSO;
    if (!raIds.length) { alvo.innerHTML = ""; return; }
    if (raBase && raBaseVersao === d.versao) { raTabelaDesenhar(); return; }
    alvo.innerHTML = '<div class="ra-msg">Carregando a lista de ações…</div>';
    if (raBasePedido === d.versao) return;
    raBasePedido = d.versao;
    api("resumo_acoes_base").then(function (b) {
      raBasePedido = null;
      if (!b.ok) { alvo.innerHTML = '<div class="ra-msg">' + esc(b.mensagem) + "</div>"; return; }
      raBase = b; raBaseVersao = b.versao;
      raTabelaDesenhar();
    }).catch(function (e) {
      raBasePedido = null;
      alvo.innerHTML = '<div class="ra-msg">Erro ao carregar as ações: ' + esc(String(e)) + "</div>";
    });
  }
  function raTabCelula(c, v, id) {
    if (c.tipo === "valor") {
      return '<td class="dir">' + raBRL.format(v || 0) + (raBase.orc[id] ? "*" : "") + "</td>";
    }
    if (c.tipo === "data") {
      var p = v ? v.split("-") : null;
      return '<td class="centro">' + (p ? p[2] + "/" + p[1] + "/" + p[0] : "–") + "</td>";
    }
    return "<td>" + (v ? esc(v) : "–") + "</td>";
  }
  function raTabelaDesenhar() {
    var alvo = document.getElementById("raTabela");
    if (!raBase || !raIds.length) { alvo.innerHTML = ""; return; }
    var cols = raBase.colunas, L = raBase.linhas;
    var ids = raIds.slice();
    if (raTab.col !== null) {
      var i = raTab.col, tipo = cols[i].tipo, dir = raTab.dir;
      ids.sort(function (a, b) {
        var x = L[a][i], y = L[b][i];
        var vx = x === "" || x === null, vy = y === "" || y === null;
        if (vx || vy) return vx === vy ? 0 : (vx ? 1 : -1);
        if (tipo === "valor") return dir * (x - y);
        return dir * String(x).localeCompare(String(y), "pt-BR", {numeric: true});
      });
    }
    var cab = cols.map(function (c, i) {
      var classe = (c.tipo === "valor" ? "dir" : (c.tipo === "data" ? "centro" : "")) +
        ((c.filtro && raFiltros[c.filtro]) || (c.filtro === "municipio" && raFiltros.top30)
          ? " ra-col-filtro" : "");
      return '<th class="' + classe + '" data-col="' + i + '" title="Ordenar">' + esc(c.rotulo) +
        (raTab.col === i ? (raTab.dir > 0 ? " ▲" : " ▼") : "") + "</th>";
    }).join("");
    var vis = ids.slice(0, raTab.limite);
    var corpo = vis.map(function (id, n) {
      return '<tr><td class="centro">' + (n + 1) + "</td>" +
        cols.map(function (c, i) { return raTabCelula(c, L[id][i], id); }).join("") + "</tr>";
    }).join("");
    var resto = ids.length - vis.length;
    alvo.innerHTML = '<div class="ra-tabela-topo"><b>AÇÕES (' + ids.length.toLocaleString("pt-BR") +
      ")</b><span>Colunas destacadas: filtros em uso. Clique no cabeçalho para ordenar.</span>" +
      '<button type="button" id="raXlsx" title="Grava em Downloads a planilha das ações filtradas">EXPORTAR EXCEL</button></div>' +
      '<div class="ra-tabela-wrap"><table class="ra-acoes"><thead><tr><th class="centro">#</th>' + cab +
      "</tr></thead><tbody>" + corpo + "</tbody></table></div>" +
      (resto > 0 ? '<button type="button" class="ra-mais">Mostrar mais ' + Math.min(RA_PASSO, resto) +
        " (faltam " + resto.toLocaleString("pt-BR") + ")</button>" : "");
    alvo.querySelectorAll("th[data-col]").forEach(function (th) {
      th.onclick = function () {
        var i = +th.dataset.col;
        if (raTab.col === i) raTab.dir = -raTab.dir; else { raTab.col = i; raTab.dir = 1; }
        raTabelaDesenhar();
      };
    });
    var mais = alvo.querySelector(".ra-mais");
    if (mais) mais.onclick = function () { raTab.limite += RA_PASSO; raTabelaDesenhar(); };
    document.getElementById("raXlsx").onclick = raExportar;
  }
  function raExportar() {
    var msg = document.getElementById("raMsg");
    var botao = document.getElementById("raXlsx");
    if (botao) botao.disabled = true;
    msg.textContent = "Gerando a planilha…";
    api("resumo_acoes_xlsx", raFiltros)
      .then(function (r) { msg.textContent = r.mensagem; })
      .catch(function (e) { msg.textContent = "Erro ao gerar a planilha: " + e; })
      .finally(function () { var b = document.getElementById("raXlsx"); if (b) b.disabled = false; });
  }

  function raMontar(d) {
    raDesenharFiltros(d.filtros, d.datas);
    raTabelaAtualizar(d);
    if (d.vazio) {
      document.getElementById("raQuadro").innerHTML =
        '<div class="ra-vazio">Nenhuma ação com os filtros escolhidos.</div>';
      return;
    }
    var linha = function (l, classeExtra) {
      return "<td>" + esc(l.status) + "</td>" +
        l.celulas.map(function (c) { return raCelula(c, classeExtra); }).join("") +
        raCelula(l.total, "tot");
    };
    document.getElementById("raQuadro").innerHTML =
      (d.aviso_cards ? '<div class="ra-aviso-cards">' + esc(d.aviso_cards) + "</div>" : "") +
      '<div class="ra-cards" style="grid-template-columns:repeat(' + d.cards.length + ', minmax(0, 360px))">' +
      d.cards.map(function (c) {
        return '<div class="ra-card" style="--ra-cor:#' + esc(c.cor) + '"><span>' + esc(c.rotulo) +
          "</span><b>" + esc(c.txt) + "</b><span>" + esc(c.qtd) + "</span></div>";
      }).join("") + "</div>" +
      '<div class="ra-rolagem"><table class="ra-tab"><thead><tr><th>Status</th>' +
      d.eixos.map(function (e) { return "<th>" + esc(e) + "</th>"; }).join("") +
      '<th class="tot">Total</th></tr></thead><tbody>' +
      d.linhas.map(function (l) { return "<tr>" + linha(l) + "</tr>"; }).join("") +
      "</tbody><tfoot><tr>" + linha(d.total, "tot") + "</tr></tfoot></table></div>" +
      raQuadroObjeto(d.quadro_objeto) +
      '<div class="ra-nota">' + esc(d.nota) + "</div>";
  }
  // Quadro Status × Objeto: mesmo formato, com o eixo do objeto numa linha
  // de cabeçalho acima (mesma ordem de eixos do quadro Status × Eixo).
  function raQuadroObjeto(q) {
    if (!q || !q.grupos || !q.grupos.length) return "";
    var inicio = {}, n = 0;
    q.grupos.forEach(function (g) { inicio[n] = true; n += g.objetos.length; });
    var cel = function (c, i, classe) {
      var cl = (classe ? classe + " " : "") + (inicio[i] ? "ra-ini-grupo" : "");
      return raCelula(c, cl.trim());
    };
    var linha = function (l, classeExtra) {
      return "<td>" + esc(l.status) + "</td>" +
        l.celulas.map(function (c, i) { return cel(c, i, classeExtra); }).join("") +
        raCelula(l.total, "tot ra-ini-grupo");
    };
    return '<div class="ra-quadro-titulo">QUADRO RESUMO — STATUS × OBJETO ' +
      "<small>(objetos agrupados pelo eixo, na mesma ordem do quadro acima; role para o lado)</small></div>" +
      '<div class="ra-rolagem"><table class="ra-tab ra-tab-obj"><thead><tr>' +
      '<th rowspan="2">Status</th>' +
      q.grupos.map(function (g) {
        return '<th class="ra-grupo ra-ini-grupo" colspan="' + g.objetos.length + '">' + esc(g.eixo) + "</th>";
      }).join("") +
      '<th class="tot ra-ini-grupo" rowspan="2">Total</th></tr><tr>' +
      q.grupos.map(function (g) {
        return g.objetos.map(function (o, j) {
          return '<th class="ra-obj' + (j === 0 ? " ra-ini-grupo" : "") + '">' + esc(o) + "</th>";
        }).join("");
      }).join("") +
      "</tr></thead><tbody>" +
      q.linhas.map(function (l) { return "<tr>" + linha(l) + "</tr>"; }).join("") +
      "</tbody><tfoot><tr>" + linha(q.total, "tot") + "</tr></tfoot></table></div>";
  }
  function raConsultar() {
    var msg = document.getElementById("raMsg");
    var pedido = ++raPedido;
    msg.textContent = "Lendo a planilha Ações do Cronograma…";
    api("resumo_acoes", raFiltros).then(function (r) {
      if (pedido !== raPedido) return;
      if (!r.ok) { msg.textContent = r.mensagem; return; }
      msg.textContent = "";
      raMontar(r.dados);
    }).catch(function (e) { msg.textContent = "Erro ao ler a planilha: " + e; });
  }
  function raGerar() {
    var msg = document.getElementById("raMsg");
    var botao = document.getElementById("raGerar");
    botao.disabled = true;
    msg.textContent = "Gerando a apresentação…";
    var pedido = ++raPedido;
    api("resumo_acoes_pptx", raFiltros).then(function (r) {
      msg.textContent = r.mensagem;
      if (r.dados && pedido === raPedido) raMontar(r.dados);
    }).catch(function (e) { msg.textContent = "Erro ao gerar a apresentação: " + e; })
      .finally(function () { botao.disabled = false; });
  }
  function raHtml() {
    var msg = document.getElementById("raMsg");
    var botao = document.getElementById("raHtml");
    botao.disabled = true;
    msg.textContent = "Gerando o HTML…";
    // no navegador (servidor web) a aba é aberta já, no clique, para não
    // ser barrada como pop-up; o conteúdo chega depois
    var aba = window.PAC_MODO_WEB ? window.open("", "_blank") : null;
    api("resumo_acoes_html", raFiltros).then(function (r) {
      msg.textContent = r.mensagem || "";
      if (r.html && aba) { aba.document.open(); aba.document.write(r.html); aba.document.close(); }
      else if (aba) aba.close();
    }).catch(function (e) { msg.textContent = "Erro ao gerar o HTML: " + e; if (aba) aba.close(); })
      .finally(function () { botao.disabled = false; });
  }
  window.abrirResumoAcoes = function () {
    overlay.classList.add("aberto");
    overlay.setAttribute("aria-hidden", "false");
    raConsultar();
  };
  function fechar() {
    overlay.classList.remove("aberto");
    overlay.setAttribute("aria-hidden", "true");
  }
  document.getElementById("raFechar").onclick = fechar;
  document.getElementById("raGerar").onclick = raGerar;
  document.getElementById("raHtml").onclick = raHtml;
  overlay.addEventListener("click", function () {
    overlay.querySelectorAll(".ra-ms.aberto").forEach(function (b) { b.classList.remove("aberto"); });
  });
  document.addEventListener("keydown", function (e) {
    if (e.key !== "Escape" || !overlay.classList.contains("aberto")) return;
    var aberto = overlay.querySelector(".ra-ms.aberto");
    if (aberto) aberto.classList.remove("aberto"); else fechar();
    e.stopPropagation();
  }, true);
  document.querySelectorAll("[data-abrir-resumo-acoes]").forEach(function (b) {
    b.addEventListener("click", window.abrirResumoAcoes);
  });
})();
</script>
"""
