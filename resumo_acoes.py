# -*- coding: utf-8 -*-
"""Resumo das Ações Prioritárias e PAC — janela própria do dash.

Portado do Relatório SSP (pedido do usuário em 12/11/2026: "incluir esse
painel no dash do panorama pac conforme feito no relatório SSP"). Base
própria, a mesma do MONITORAMENTO ("Ações do Cronograma.xlsx", lida da
pasta do programa); independe da planilha e dos filtros do PAC.

Ligado ao módulo principal pelo mesmo padrão do gerar_pptx_pac:
configurar(sys.modules[__name__]).
"""
import base64
import gzip
import json
import math
import os
import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd

import identidade_visual as iv

pac = None


def configurar(modulo):
    global pac
    pac = modulo


RA_PLANILHA = "Ações do Cronograma.xlsx"
RA_COL_EIXO, RA_COL_STATUS = "Eixo", "Status"
RA_COL_VALOR, RA_COL_ORCAMENTO = "Valor de Investimento", "Valor do Orçamento"
RA_EIXOS = ["Mobilidade", "Educação", "Saneamento",
            "Prevenção a Desastres Naturais", "Turismo e Cultura", "Saúde"]
# Cores da fonte única (ver identidade_visual.py) — RA_AZUL era #1F5C85,
# RA_CLARO era #EEF3F8 e o segundo tom de RA_CORES_TOP3 era um dourado
# (#F2B632), todos divergentes da referência — corrigido a pedido do
# usuário em 14/11/2026.
RA_CORES_TOP3, RA_COR_TOTAL = [iv.VERDE, iv.LARANJA, iv.TEAL], iv.NAVY
RA_AZUL, RA_CLARO = iv.AZUL, iv.CINZA_CLARO
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

# População de cada um dos 30 acima — MESMA fonte/arquivo de RA_TOP30_FONTE
# (IBGE, "Estimativas da População Residente nos Municípios Brasileiros",
# data de referência 01/07/2026, arquivo oficial
# https://ftp.ibge.gov.br/Estimativas_de_Populacao/Estimativas_2026/estimativa_dou_2026.xlsx,
# aba "municípios", UF=BA). Conferido em 15/11/2026: a ordem de RA_TOP30
# bate exatamente com a ordem decrescente de população deste arquivo, e o
# 31º colocado é mesmo Bom Jesus da Lapa — usada para os indicadores da
# avaliação quantitativa (item 3.1 do pedido do usuário em 15/11/2026:
# per capita, gap de participação, Índice de Necessidade de Investimento).
RA_TOP30_POP = {
    "Salvador": 2_559_945, "Feira de Santana": 663_408,
    "Vitória da Conquista": 398_970, "Camaçari": 323_676,
    "Juazeiro": 257_616, "Lauro de Freitas": 221_023,
    "Itabuna": 196_041, "Ilhéus": 189_259, "Porto Seguro": 184_087,
    "Barreiras": 172_521, "Jequié": 169_623, "Alagoinhas": 161_682,
    "Teixeira de Freitas": 154_114, "Eunápolis": 121_577,
    "Simões Filho": 120_429, "Luís Eduardo Magalhães": 119_947,
    "Paulo Afonso": 119_686, "Santo Antônio de Jesus": 110_268,
    "Guanambi": 93_873, "Valença": 90_075, "Jacobina": 87_121,
    "Serrinha": 84_928, "Irecê": 79_105, "Senhor do Bonfim": 78_193,
    "Casa Nova": 76_416, "Dias d'Ávila": 75_516, "Campo Formoso": 75_311,
    "Candeias": 74_645, "Brumado": 74_617, "Conceição do Coité": 71_860,
}


def _chave_municipio(nome):
    """Só letras e números, sem acento e em maiúsculas: "Dias D'Ávila" e
    "DIAS D AVILA" viram a mesma chave."""
    s = unicodedata.normalize("NFKD", str(nome or ""))
    return "".join(c for c in s if c.isalnum()
                   and not unicodedata.combining(c)).upper()


RA_TOP30_CHAVES = {_chave_municipio(m) for m in RA_TOP30}
# Mesma chave normalizada de RA_TOP30_CHAVES, pra achar a população de um
# município pelo NOME COMO ELE VEM DA PLANILHA (pode diferir de RA_TOP30 em
# acento/maiúscula — ex.: "Dias D'Ávila" na planilha x "Dias d'Ávila" aqui),
# em vez de por igualdade exata de string — corrigido a pedido do usuário
# em 15/11/2026, depois de "Dias D'Ávila" sair com população 0 no teste.
RA_TOP30_POP_CHAVE = {_chave_municipio(nome): pop for nome, pop in RA_TOP30_POP.items()}

# ---- Carteira dos 30 maiores municípios: mapa oculto + gráfico/avaliação ---
# Pedido do usuário em 15/11/2026: o mapa de investimento por município sai
# de vista (mas as funções continuam existindo, religáveis por aqui), e no
# lugar entram um gráfico de colunas e uma avaliação quantitativa dos 30
# maiores municípios por população, antes da grade de cards.
RA_HTML_MOSTRAR_MAPA = True   # religado em 16/11/2026, com as cores do índice
RA_COR_SEM_ACAO_TOP30 = "#C3CAD5"   # 30 maiores sem ação no recorte
RA_COR_FORA_TOP30 = "#F1F3F6"       # demais municípios do estado

# Pesos do Índice de Necessidade de Investimento (0 a 100) — combinação
# ponderada e normalizada (min-max entre os 30) de: investimento per capita
# (invertido — quanto menor, maior a necessidade), gap de participação
# (população % menos investimento %), fração do investimento ainda em
# status iniciais da carteira (maturidade) e cobertura — nº de ações e de
# eixos atendidos (invertida — quanto menor a cobertura, maior a
# necessidade). Pedido do usuário em 15/11/2026; pesos documentados aqui
# porque a fórmula depende deles — mudar um sem mudar o outro descalibra o
# índice.
RA_PESOS_NECESSIDADE = {
    "per_capita": 0.40,
    "gap": 0.30,
    "maturidade": 0.20,
    "cobertura": 0.10,
}

# Classificação de maturidade da carteira por STATUS — mesmo código "NN -"
# no início do texto do Status já usado em todo o resto do módulo (ver
# re.sub em _acoes_do_municipio), aqui para decidir se o valor da ação
# ainda está em estágio inicial (estudo/projeto/licitação) ou já avançou
# para obra/conclusão. "14 - Obras Canceladas" fica de fora dos dois
# grupos — não é nem um bom nem um mau sinal de maturidade, é uma ação que
# saiu da carteira.
RA_CODIGOS_STATUS_INICIAL = {"03", "04", "05", "07", "08", "12", "13"}
RA_CODIGOS_STATUS_AVANCADO = {"01", "02", "09", "10"}

# Cores das 3 classes de prioridade (mesma paleta já usada nas faixas do
# mapa, RA_FAIXAS_MAPA — reaproveitadas aqui a pedido do usuário em
# 15/11/2026, sem inventar uma paleta nova): vermelho = Prioridade alta,
# laranja = Atenção, azul da marca (RA_AZUL) = Adequado.
RA_COR_PRIORIDADE_ALTA = "#D32F45"
RA_COR_ATENCAO = "#F0834E"
RA_COR_ADEQUADO = f"#{RA_AZUL}"


def _maturidade_status(status_bruto):
    """"inicial"/"avancado"/None (cancelada ou código não mapeado) a partir
    do código "NN - " no início do texto bruto do Status."""
    m = re.match(r"^\s*(\d+)\s*-", str(status_bruto or ""))
    codigo = m.group(1) if m else ""
    if codigo in RA_CODIGOS_STATUS_INICIAL:
        return "inicial"
    if codigo in RA_CODIGOS_STATUS_AVANCADO:
        return "avancado"
    return None
_ra_cache = None   # (caminho, data do arquivo, base)

# PPTX (mesmas cores/fontes do slide do SSP)
# Mesma fonte do PPTX de referência (gerar_pptx_pac.py: FONTE = "Roboto")
# — antes "Segoe UI"/"Segoe UI Black" — corrigido a pedido do usuário em
# 14/11/2026. Sem um peso "Black" equivalente de Roboto já estabelecido no
# projeto, PE_FONTE_BLACK vira a mesma Roboto (o negrito já vem do
# estilo "b"/"k" marcado em cada chamada de escrever()).
PE_FONTE = iv.FONTE_PPTX
PE_FONTE_BLACK = iv.FONTE_PPTX
PE_COR = {"navy": iv.NAVY, "cinza": iv.CINZA_TEXTO, "gov": iv.AZUL_CLARO_MARCA,
          "subcab": iv.AZUL_CLARO_TEXTO}


# ---- utilitários ------------------------------------------------------------
def caminho_planilha():
    caminho = pac.caminho_recurso(RA_PLANILHA) if pac else RA_PLANILHA
    return caminho if os.path.exists(caminho) else None


def data_atualizacao_base():
    """Data/hora de modificação de "Ações do Cronograma.xlsx" — base própria
    do MONITORAMENTO, independente da planilha do PAC (ver cabeçalho do
    módulo). Usada no cartão MONITORAMENTO da tela HOME — pedido do usuário
    em 14/11/2026. None se o arquivo não existir (tela HOME sem o cartão
    quebrar por causa disso)."""
    caminho = caminho_planilha()
    if caminho is None:
        return None
    return datetime.fromtimestamp(os.path.getmtime(caminho))


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


def _ra_brl(v):
    """1234.5 -> '1.234,50' (como a tabela da página mostra o valor)."""
    return f"{float(v):,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")


def _ra_texto_busca(v):
    """Célula como texto pesquisável: data em dd/mm/aaaa, número também no
    formato brasileiro, tudo sem acento e em minúsculas."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, (datetime, pd.Timestamp)):
        return f"{v:%d/%m/%Y}"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        txt = _ra_brl(v)
        return f"{v} {txt}" if float(v) != int(float(v)) else f"{int(float(v))} {txt}"
    return _normalizar(" ".join(str(v).split()))


def _ra_termos(termo):
    """Palavras do que foi digitado (sem acento, minúsculas)."""
    return [p for p in _normalizar(termo).split() if p]


def _ra_contem_todos(serie_txt, termos):
    """Linhas cujo texto (já normalizado) contém TODAS as palavras."""
    marca = pd.Series(True, index=serie_txt.index)
    for p in termos:
        marca &= serie_txt.str.contains(p, regex=False)
    return marca


def _ra_texto_coluna(tipo, serie):
    """Coluna da tabela como a página mostra, normalizada (busca por coluna)."""
    if tipo == "data":
        return serie.map(lambda d: f"{d:%d/%m/%Y}" if pd.notna(d) else "")
    if tipo == "valor":
        return serie.map(lambda v: _ra_brl(v or 0))
    return serie.map(_normalizar)


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
    # texto de TODAS as colunas da base, normalizado, para a busca geral
    # (pedido do usuário em 17/11/2026)
    originais = list(df.columns)
    df["_busca"] = [" | ".join(_ra_texto_busca(v) for v in linha)
                    for linha in df[originais].itertuples(index=False, name=None)]
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

    # Buscas digitadas (pedido do usuário em 17/11/2026): "busca" = geral,
    # em todas as colunas da base; "bcol" = {rótulo da coluna da tabela:
    # texto}, só naquela coluna. Todas as palavras precisam aparecer. Valem
    # como filtro: cards, quadros, tabela, Excel e HTML seguem o recorte.
    busca_geral = str(filtros.get("busca") or "").strip()
    busca_cols = {r: str(t).strip() for r, t in (filtros.get("bcol") or {}).items()
                  if str(t or "").strip()}
    marca_busca = pd.Series(True, index=base.index)
    if busca_geral:
        marca_busca &= _ra_contem_todos(base["_busca"], _ra_termos(busca_geral))
    if busca_cols:
        for rot, tipo, _c, serie in _ra_colunas_tabela(base):
            if rot in busca_cols:
                marca_busca &= _ra_contem_todos(_ra_texto_coluna(tipo, serie),
                                                _ra_termos(busca_cols[rot]))

    def marca_datas(com_top30=True):
        marca = marca_busca.copy()
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
    filtrado = bool(sel or sel_datas or top30 or busca_geral or busca_cols)
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
                   for c, _, rotulo in campos_data if c in sel_datas]
                + ([f'Busca: "{busca_geral}"'] if busca_geral else [])
                + [f'{rot} contém "{t}"' for rot, t in busca_cols.items()])
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
    # D8DCE3 = cinza de borda da identidade única (era D9DFE8) — corrigido
    # a pedido do usuário em 14/11/2026.
    cinza, linha_tab = PE_COR["cinza"], "D8DCE3"
    px = lambda v: Emu(int(v * 9525))   # noqa: E731 — slide de 1280 x 720 px

    def escrever(par, txt, tam, cor, estilo="n", alin=PP_ALIGN.LEFT):
        par.alignment = alin
        r = par.add_run()
        r.text = txt
        r.font.size = Pt(tam * 0.75)
        r.font.color.rgb = RGBColor.from_string(cor)
        r.font.name = PE_FONTE_BLACK if estilo == "k" else PE_FONTE
        # "k" (antes a fonte "Segoe UI Black" fazia o peso sozinha) agora
        # é a mesma Roboto do resto do PPTX — sem um peso Black dedicado,
        # o negrito do python-pptx é quem carrega esse destaque agora;
        # sem isso "k" saía sem nenhum peso extra — corrigido a pedido do
        # usuário em 14/11/2026.
        r.font.bold = estilo in ("b", "k")
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
    """({chave normalizada do geojson: {nome, valor, qtd, orc, eixos}} do
    recorte atual, a lista dos que não casaram com o geojson, dados dos
    cards de ação dos 30 maiores — ver _dados_cards_acoes_top30)."""
    import relatorio_monitoramento as mon
    base = _ra_base(caminho_base)
    if dados.get("entregue"):
        base, _ = _ra_juntar_entregues(base)
    df = base.iloc[dados["ids"]]
    if "Município" not in df.columns:
        return {}, [], {"campos": [], "acoes": []}
    chave = df["Município"].map(mon.normalizar).map(
        lambda k: mon.APELIDOS_MUNICIPIO.get(k, k))
    saida = {}
    for k, parte in df.groupby(chave):
        if not k:
            continue
        por_eixo = parte.groupby("_eixo")["_valor"].sum().sort_values(ascending=False)
        eixos_com_valor = [(e, float(v)) for e, v in por_eixo.items() if v > 0]
        # Maturidade da carteira (pedido do usuário em 15/11/2026, ver
        # RA_CODIGOS_STATUS_INICIAL/AVANCADO e _maturidade_status): soma do
        # _valor por grupo de status, à parte dos campos já formatados de
        # "acoes" — o índice de necessidade precisa do número bruto, não do
        # texto "R$ X Mi" que _acoes_do_municipio devolve.
        maturidade = parte["_f_status"].map(_maturidade_status)
        saida[k] = {"valor": float(parte["_valor"].sum()), "qtd": int(len(parte)),
                    "orc": bool(parte["_orc"].any()),
                    "nome": parte["Município"].astype(str).str.strip().mode().iloc[0],
                    "eixos": eixos_com_valor[:4],
                    "n_eixos": len(eixos_com_valor),
                    "valor_inicial": float(parte.loc[maturidade == "inicial", "_valor"].sum()),
                    "valor_avancado": float(parte.loc[maturidade == "avancado", "_valor"].sum()),
                    "acoes": _acoes_do_municipio(parte)}
    return saida, [], _dados_cards_acoes_top30(df)


def _dados_cards_acoes_top30(df):
    """Uma linha por ação dos 30 maiores municípios (RA_TOP30), pra janela de
    cards aberta ao clicar num objeto na grade Município > Status > Objeto
    (ver detalhe_rank/cartao_municipio) — pedido do usuário em 15/11/2026.
    Município/status/objeto na MESMA forma usada pelos data-* do clique (ver
    _acoes_do_municipio), pra o filtro em JS bater certo com a linha."""
    campos = ["mu", "st", "ob", "nome", "mun", "ex", "ds", "av", "pz", "val", "orc"]
    if "_top30" not in df.columns:
        return {"campos": campos, "acoes": []}
    parte = df[df["_top30"]]
    if parte.empty:
        return {"campos": campos, "acoes": []}
    objeto = (parte["_f_objeto"] if "_f_objeto" in parte.columns
              else pd.Series(RA_SEM_INFO, index=parte.index))
    status = parte["_f_status"].map(lambda s: re.sub(r"^\d+\s*-\s*", "", str(s)))
    nome_acao = (parte["Ação"] if "Ação" in parte.columns else objeto).map(
        lambda v: " ".join(str(v).split()) if pd.notna(v) else "")
    executor = (parte["Executor"] if "Executor" in parte.columns
                else pd.Series("", index=parte.index)).fillna("")
    if "Detalhamento Escopo" in parte.columns:
        descricao = parte["Detalhamento Escopo"]
    elif "Escopo" in parte.columns:
        descricao = parte["Escopo"]
    else:
        descricao = pd.Series("", index=parte.index)
    descricao = descricao.fillna("")
    avanco = (pd.to_numeric(parte["Avanço"], errors="coerce") if "Avanço" in parte.columns
              else pd.Series(float("nan"), index=parte.index))
    prazo = (parte["_d_entrega"] if "_d_entrega" in parte.columns
             else pd.Series(pd.NaT, index=parte.index))

    acoes = []
    for i in parte.index:
        av, pz, nome = avanco.loc[i], prazo.loc[i], str(nome_acao.loc[i]).strip()
        acoes.append([
            _chave_municipio(parte.loc[i, "Município"]), status.loc[i], objeto.loc[i],
            nome or str(objeto.loc[i]),
            str(parte.loc[i, "Município"]).strip() if pd.notna(parte.loc[i, "Município"]) else "",
            str(executor.loc[i]).strip(), " ".join(str(descricao.loc[i]).split()),
            round(float(av), 1) if pd.notna(av) else None,
            pz.strftime("%d/%m/%Y") if pd.notna(pz) else "A definir",
            _ra_rs(float(parte.loc[i, "_valor"])) + ("*" if bool(parte.loc[i, "_orc"]) else ""),
            bool(parte.loc[i, "_orc"]),
        ])
    return {"campos": campos, "acoes": acoes}


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


_ra_mapa_base_cache = None


def _svg_mapa_base(largura=640):
    """Bahia em SVG só com o desenho: cada município com data-k (chave de
    _chave_municipio) e data-n (nome), sem cor de dado — quem pinta é o JS
    (ligarMapa), com as cores do recorte visível. O desenho não muda, então
    fica guardado para as próximas páginas."""
    global _ra_mapa_base_cache
    if _ra_mapa_base_cache:
        return _ra_mapa_base_cache
    import relatorio_monitoramento as mon
    geo = pac.carregar_geojson_municipios() if pac else None
    if not geo:
        return '<div class="sem-mapa">Mapa indisponível: falta municipios_bahia.geojson</div>'
    feats = geo["features"]
    pts = [p for f in feats for pol in mon._poligonos(f) for anel in pol for p in anel]
    minlon, maxlon = min(p[0] for p in pts), max(p[0] for p in pts)
    minlat, maxlat = min(p[1] for p in pts), max(p[1] for p in pts)
    cx = math.cos(math.radians((minlat + maxlat) / 2))
    esc_x = largura / ((maxlon - minlon) * cx)
    altura = (maxlat - minlat) * esc_x
    def projetar(lon, lat):
        return (lon - minlon) * cx * esc_x, (maxlat - lat) * esc_x

    def centro(anel):
        # centroide do polígono (fórmula da área); cai pra média dos pontos
        # se a área for ~0
        pts = [projetar(lon, lat) for lon, lat in anel]
        a = sx = sy = 0.0
        for (x1, y1), (x2, y2) in zip(pts, pts[1:] + pts[:1]):
            c = x1 * y2 - x2 * y1
            a += c
            sx += (x1 + x2) * c
            sy += (y1 + y2) * c
        if abs(a) < 1e-9:
            return (sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts)), 0.0
        return (sx / (3 * a), sy / (3 * a)), abs(a) / 2

    caminhos = []
    for f in feats:
        nome = str(f["properties"].get("name", ""))
        aneis = [anel for pol in mon._poligonos(f) for anel in pol]
        d = "".join("M" + "L".join(
            "{:.1f},{:.1f}".format(*projetar(lon, lat)) for lon, lat in anel) + "Z" for anel in aneis)
        # centro do maior anel: onde vai o rótulo do município (ver
        # rotulosMapa no JS — só nos de Prioridade alta, pedido do usuário
        # em 16/11/2026)
        (px, py), _ = max((centro(anel) for anel in aneis if len(anel) > 2),
                          key=lambda c: c[1], default=((0.0, 0.0), 0.0))
        caminhos.append(f'<path d="{d}" fill="{RA_COR_FORA_TOP30}" '
                        f'data-k="{_chave_municipio(nome)}" data-n="{_esc(nome)}" '
                        f'data-cx="{px:.1f}" data-cy="{py:.1f}"/>')
    _ra_mapa_base_cache = (
        f'<svg class="mapa" viewBox="0 0 {largura:.0f} {altura:.0f}" role="img" '
        'aria-label="Mapa dos 30 maiores municípios pelo Índice de Necessidade">'
        + "".join(caminhos) + "</svg>")
    return _ra_mapa_base_cache


# ---- Avaliação da carteira dos 30 maiores municípios -----------------------
# Pedido do usuário em 15/11/2026 (substitui o mapa, oculto por
# RA_HTML_MOSTRAR_MAPA): indicadores quantitativos/comparativos e o Índice
# de Necessidade de Investimento de cada um dos 30 maiores municípios por
# população, usados no gráfico de colunas, na "Leitura da carteira" e no
# selo de cada cartão.
def _calcular_indicadores_carteira_top30(ranking):
    """[{pos_pop, nome, populacao, valor, qtd, orc, per_capita, pct_pop,
    pct_investimento, gap_participacao, fracao_inicial, n_eixos, indice,
    classe, cor_classe}], mediana_per_capita — um item por município de
    `ranking` (já na ordem de RA_TOP30, com os campos de
    _investimento_por_municipio ou os de reserva de município sem ação).
    A mediana usa só os municípios com ação no recorte atual, para o filtro
    não arrastar a mediana pra baixo com municípios zerados por ele."""
    investimento_total = sum(m["valor"] for m in ranking) or 1.0
    populacao_total = sum(RA_TOP30_POP.values()) or 1.0

    itens = []
    for i, m in enumerate(ranking):
        populacao = RA_TOP30_POP_CHAVE.get(_chave_municipio(m["nome"]), 0)
        per_capita = (m["valor"] / populacao) if populacao else 0.0
        pct_pop = populacao / populacao_total * 100
        pct_investimento = m["valor"] / investimento_total * 100
        base_maturidade = m.get("valor_inicial", 0.0) + m.get("valor_avancado", 0.0)
        fracao_inicial = (m.get("valor_inicial", 0.0) / base_maturidade
                          if base_maturidade else 0.0)
        itens.append({
            "pos_pop": i + 1, "nome": m["nome"], "populacao": populacao,
            "valor": m["valor"], "qtd": m["qtd"], "orc": m.get("orc", False),
            "per_capita": per_capita, "pct_pop": pct_pop,
            "pct_investimento": pct_investimento,
            "gap_participacao": pct_pop - pct_investimento,
            "fracao_inicial": fracao_inicial, "n_eixos": m.get("n_eixos", 0),
        })

    def minmax(valores):
        # 0.5 (meio da escala) quando todos os valores são iguais — não há
        # necessidade nem excesso relativo para distinguir.
        baixo, alto = min(valores), max(valores)
        if alto <= baixo:
            return [0.5] * len(valores)
        return [(v - baixo) / (alto - baixo) for v in valores]

    cobertura_qtd = minmax([it["qtd"] for it in itens])
    cobertura_eixo = minmax([it["n_eixos"] for it in itens])
    coberturas = [(a + b) / 2 for a, b in zip(cobertura_qtd, cobertura_eixo)]
    norm_per_capita = minmax([it["per_capita"] for it in itens])
    norm_gap = minmax([it["gap_participacao"] for it in itens])
    norm_maturidade = minmax([it["fracao_inicial"] for it in itens])
    pesos = RA_PESOS_NECESSIDADE

    for it, npc, ngap, nmat, cob in zip(
        itens, norm_per_capita, norm_gap, norm_maturidade, coberturas
    ):
        # per_capita e cobertura entram invertidos: quanto MENOR o valor
        # per capita/a cobertura, MAIOR a necessidade de investimento.
        it["indice"] = round(100 * (
            pesos["per_capita"] * (1 - npc)
            + pesos["gap"] * ngap
            + pesos["maturidade"] * nmat
            + pesos["cobertura"] * (1 - cob)
        ), 1)

    indices = pd.Series([it["indice"] for it in itens])
    q75, q50 = indices.quantile(0.75), indices.quantile(0.50)
    for it in itens:
        if it["indice"] >= q75:
            it["classe"], it["cor_classe"] = "Prioridade alta", RA_COR_PRIORIDADE_ALTA
        elif it["indice"] >= q50:
            it["classe"], it["cor_classe"] = "Atenção", RA_COR_ATENCAO
        else:
            it["classe"], it["cor_classe"] = "Adequado", RA_COR_ADEQUADO

    per_capitas_com_acao = [it["per_capita"] for it in itens if it["qtd"] > 0]
    mediana_per_capita = (float(pd.Series(per_capitas_com_acao).median())
                          if per_capitas_com_acao else 0.0)
    return itens, mediana_per_capita


def _gerar_leitura_carteira(itens, mediana_per_capita):
    """(html da lista dos 5 mais prioritários, html do parágrafo geral) —
    texto gerado a partir dos números de `itens` (ver
    _calcular_indicadores_carteira_top30), nunca um texto fixo."""
    por_investimento = sorted(itens, key=lambda it: -it["valor"])
    pos_investimento = {it["nome"]: i + 1 for i, it in enumerate(por_investimento)}
    prioritarios = sorted(itens, key=lambda it: -it["indice"])[:5]

    frases = []
    for it in prioritarios:
        delta = ((it["per_capita"] - mediana_per_capita) / mediana_per_capita * 100
                 if mediana_per_capita else 0.0)
        sentido = "acima" if delta >= 0 else "abaixo"
        frases.append(
            f'<li><b>{_esc(it["nome"])}</b>: {it["pos_pop"]}º mais populoso, '
            f'{pos_investimento[it["nome"]]}º em investimento; per capita de '
            f'{_esc(_rs_cheio(it["per_capita"]))}, {abs(delta):.0f}% {sentido} da '
            f'mediana dos 30; {it["fracao_inicial"] * 100:.0f}% do valor ainda em '
            'estudo/projeto/licitação.</li>'
        )

    total = sum(it["valor"] for it in itens) or 1.0
    valor_salvador = next(
        (it["valor"] for it in itens if _chave_municipio(it["nome"]) == _chave_municipio("Salvador")),
        0.0,
    )
    com_acao = sum(1 for it in itens if it["qtd"])
    abaixo_mediana = sum(1 for it in itens if it["qtd"] and it["per_capita"] < mediana_per_capita)
    paragrafo = (
        f'Salvador concentra {valor_salvador / total * 100:.0f}% do investimento '
        f'somado dos 30 maiores municípios por população. Entre os {com_acao} com '
        f'ação no recorte, {abaixo_mediana} têm investimento per capita abaixo da '
        'mediana do grupo — a desigualdade por habitante é maior do que a '
        'distribuição do investimento bruto, sozinha, sugere.'
    )
    return "".join(frases), paragrafo


# Geometria do gráfico de colunas (SVG) — ver _svg_grafico_investimento_top30.
# Margens esquerda/direita calibradas pelo texto mais largo que de fato
# aparece nelas (rótulo "R$ 1,04 BI" do eixo Y e o nome do último município,
# sempre "Conceição do Coité" — RA_TOP30 é fixa), medindo o getBBox() real
# no navegador: com as margens antigas (58/10) os rótulos mais largos do
# eixo Y e o nome da última coluna ficavam com as pontas cortadas pelo
# viewBox do SVG — corrigido a pedido do usuário em 15/11/2026.
RA_GRAFICO_LARGURA = 920
RA_GRAFICO_ALTURA = 260
RA_GRAFICO_MARGEM_ESQ = 68
RA_GRAFICO_MARGEM_DIR = 90
RA_GRAFICO_MARGEM_SUP = 34   # espaço para o rótulo das barras em destaque
RA_GRAFICO_MARGEM_INF = 108


def _svg_grafico_investimento_top30(itens, quebra, escala, vmax, mediana_per_capita):
    """Colunas (SVG puro, sem internet) do investimento dos 30 maiores
    municípios por população, na ordem de RA_TOP30 — substitui o mapa
    (RA_HTML_MOSTRAR_MAPA=False) a pedido do usuário em 15/11/2026. Cor por
    classe de prioridade (ver _calcular_indicadores_carteira_top30); traço
    pontilhado = investimento esperado pela população do município, pela
    mediana per capita dos que têm ação no recorte."""
    n = len(itens)
    largura_total = RA_GRAFICO_LARGURA
    altura_total = RA_GRAFICO_MARGEM_SUP + RA_GRAFICO_ALTURA + RA_GRAFICO_MARGEM_INF
    largura_util = largura_total - RA_GRAFICO_MARGEM_ESQ - RA_GRAFICO_MARGEM_DIR
    largura_col = largura_util / n if n else largura_util
    largura_barra = largura_col * 0.62
    base_y = RA_GRAFICO_MARGEM_SUP + RA_GRAFICO_ALTURA

    def y_de(valor):
        return base_y - (min(valor, escala) / escala * RA_GRAFICO_ALTURA if escala else 0)

    partes = []
    n_linhas_grade = 4
    for i in range(n_linhas_grade + 1):
        y = base_y - (RA_GRAFICO_ALTURA * i / n_linhas_grade)
        valor_linha = escala * i / n_linhas_grade
        partes.append(
            f'<line x1="{RA_GRAFICO_MARGEM_ESQ}" y1="{y:.1f}" x2="{largura_total - 6}" y2="{y:.1f}" '
            'class="grafico-grade"/>'
            f'<text x="{RA_GRAFICO_MARGEM_ESQ - 8}" y="{y + 4:.1f}" text-anchor="end" '
            f'class="grafico-eixo-y">{_esc(_ra_rs(valor_linha) if valor_linha else "R$ 0")}</text>'
        )

    for i, it in enumerate(itens):
        x = RA_GRAFICO_MARGEM_ESQ + i * largura_col + (largura_col - largura_barra) / 2
        tem_acao = it["qtd"] > 0
        y = y_de(it["valor"]) if tem_acao else base_y - 3
        altura_barra = max(0.0, base_y - y) if tem_acao else 3.0
        cor = it["cor_classe"] if tem_acao else "#D8DCE3"
        # investimento adequado esmaecido, pra destacar quem precisa de
        # atenção — pedido do usuário em 16/11/2026
        esmaecer = (' fill-opacity="0.35"'
                    if tem_acao and it["classe"] == "Adequado" else "")
        tip = (
            f'data-n="{_esc(it["nome"])}" data-pos="{it["pos_pop"]}" '
            f'data-v="{_esc(_ra_rs(it["valor"]) if tem_acao else "sem ação no recorte")}" '
            f'data-q="{_esc(_ra_qtd(it["qtd"]) if tem_acao else "")}" '
            f'data-pc="{_esc(_rs_cheio(it["per_capita"])) if it["populacao"] and tem_acao else ""}" '
            f'data-idx="{it["indice"]:.0f}" data-classe="{_esc(it["classe"])}"'
        )
        partes.append(
            f'<rect class="barra-top30" tabindex="0" data-alvo="mun-card-{i}" '
            f'x="{x:.1f}" y="{y:.1f}" width="{largura_barra:.1f}" height="{altura_barra:.1f}" '
            f'rx="3" fill="{cor}"{esmaecer} {tip}><title>{_esc(it["nome"])}</title></rect>'
        )
        if quebra and it["valor"] == vmax:
            partes.append(
                f'<text x="{x + largura_barra / 2:.1f}" y="{y - 6:.1f}" text-anchor="middle" '
                'class="grafico-quebra">≈</text>'
            )
        if it["populacao"] and mediana_per_capita:
            y_esperado = y_de(mediana_per_capita * it["populacao"])
            if y_esperado >= RA_GRAFICO_MARGEM_SUP:
                partes.append(
                    f'<line x1="{x - 2:.1f}" y1="{y_esperado:.1f}" x2="{x + largura_barra + 2:.1f}" '
                    f'y2="{y_esperado:.1f}" class="grafico-esperado"/>'
                )
        centro_x = x + largura_barra / 2
        # rótulo de dados só nas barras em destaque (Prioridade alta e
        # Atenção), em duas linhas (número / unidade) para caber na coluna,
        # logo acima da própria barra — pedido do usuário em 16/11/2026
        if tem_acao and it["classe"] != "Adequado":
            numero, _, unidade = _ra_rs(it["valor"]).replace("R$ ", "").partition(" ")
            y2 = max(24.0, y - 4)
            partes.append(
                f'<text x="{centro_x:.1f}" y="{y2 - 10:.1f}" text-anchor="middle" '
                f'class="grafico-rotulo" fill="{cor}">{_esc(numero)}</text>'
                + (f'<text x="{centro_x:.1f}" y="{y2:.1f}" text-anchor="middle" '
                   f'class="grafico-rotulo-un" fill="{cor}">{_esc(unidade)}</text>' if unidade else ""))
        rotulo = it["nome"] if tem_acao else f'{it["nome"]} (sem ação)'
        partes.append(
            f'<text x="{centro_x:.1f}" y="{base_y + 14:.1f}" '
            f'transform="rotate(45 {centro_x:.1f} {base_y + 14:.1f})" '
            f'class="grafico-eixo-x{"" if tem_acao else " sem-acao"}">{_esc(rotulo)}</text>'
        )

    partes.append(
        f'<line x1="{RA_GRAFICO_MARGEM_ESQ}" y1="{base_y:.1f}" x2="{largura_total - 6}" '
        f'y2="{base_y:.1f}" class="grafico-eixo"/>'
    )
    return (
        f'<svg class="grafico-top30" viewBox="0 0 {largura_total} {altura_total}" '
        'role="img" aria-label="Investimento dos 30 maiores municípios por população">'
        + "".join(partes) + "</svg>"
    )


def _escopar_css(css, escopo, nao_escopar=()):
    """Prefixa cada seletor de nível superior do `css` (inclusive dentro de
    @media) com `escopo `, convertendo ":root" e "*" nos equivalentes
    escopados — usado para isolar o CSS do Monitoramento (`.mon-ra`) de
    qualquer CSS global da Apresentação CGAPE quando os dois aparecem na
    mesma página (pedido do usuário em 15/11/2026: "nada de seletores
    globais como header, main, body, .card, .cards, :root" vazando de um
    pro outro, nos dois sentidos). Reaproveita o MESMO texto de CSS usado
    pela página autônoma (gerar_html_resumo_acoes) em vez de manter uma
    segunda cópia, que haveria o risco de divergir dela com o tempo.

    `nao_escopar`: prefixos de seletor que saem de "escopo" sem nenhum
    prefixo — caso do tooltip #mon-dica, que fica fora de qualquer
    .mon-ra de propósito (ver montar_partes_html_monitoramento)."""
    # Comentários removidos ANTES de tudo: um comentário logo antes de um
    # seletor (ex.: antes de ":root{...}") entrava junto do texto do
    # seletor e quebrava a comparação exata com ":root"/"*" abaixo — bug
    # encontrado na validação em 15/11/2026 (as variáveis de cor --navy/
    # --azul ficavam indefinidas dentro de .mon-ra).
    css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)

    def pertence_a(sel, p):
        # sel é o próprio p, ou p seguido de combinador/pseudo-classe/classe/
        # atributo (ex.: "p .x", "p:hover", "p.aberta", "p[data-x]") — NUNCA
        # um id/classe diferente que só começa igual (ex.: "p2" não é "p").
        # Faltava o caso "p.classe" (sem espaço): "#ra-oc-fundo.aberta" caía
        # no escopo .mon-ra por engano (a janela nunca aparecia, pois o
        # elemento de verdade fica FORA de .mon-ra) — bug encontrado na
        # validação em 15/11/2026.
        if sel == p:
            return True
        if not sel.startswith(p):
            return False
        seguinte = sel[len(p)]
        return not (seguinte.isalnum() or seguinte in "-_")

    def prefixar_seletor(sel):
        sel = sel.strip()
        if not sel:
            return sel
        if any(pertence_a(sel, p) for p in nao_escopar):
            return sel
        if sel == ":root":
            return escopo
        if sel == "*":
            return f"{escopo}, {escopo} *"
        if sel.startswith(escopo):
            return sel
        return f"{escopo} {sel}"

    def prefixar_lista(seletores):
        return ", ".join(prefixar_seletor(s) for s in seletores.split(","))

    def processar_bloco(texto):
        saida, i, n = [], 0, len(texto)
        while i < n:
            if texto[i].isspace():
                i += 1
                continue
            fim_sel = texto.index("{", i)
            seletor = texto[i:fim_sel]
            if seletor.strip().startswith("@media"):
                j, profundidade = fim_sel + 1, 1
                while profundidade:
                    if texto[j] == "{":
                        profundidade += 1
                    elif texto[j] == "}":
                        profundidade -= 1
                    j += 1
                conteudo = texto[fim_sel + 1:j - 1]
                saida.append(f"{seletor.strip()}{{{processar_bloco(conteudo)}}}")
                i = j
                continue
            fim_corpo = texto.index("}", fim_sel)
            corpo = texto[fim_sel + 1:fim_corpo]
            saida.append(f"{prefixar_lista(seletor)}{{{corpo}}}")
            i = fim_corpo + 1
        return "".join(saida)

    return processar_bloco(css)


def _sufixar_ids_monitoramento(html, sufixo="monra"):
    """Garante que os ids do bloco embutido na Apresentação CGAPE não
    colidam com os da própria apresentação (mesmo princípio do
    _sufixar_ids usado lá entre variantes de gestão/secretaria, ver
    gerar_pptx_pac.py) — pedido do usuário em 15/11/2026. Troca tanto a
    declaração ("id=") quanto toda referência a ela (aria-controls=,
    data-alvo=), já que o JS do Monitoramento lê esses atributos em vez de
    guardar o id como texto fixo. Também usada pelas variantes por eixo da
    página autônoma (sufixo = chave do eixo)."""
    for id_html in set(re.findall(r'\bid="([^"]+)"', html)):
        novo = f"{id_html}--{sufixo}"
        html = html.replace(f'"{id_html}"', f'"{novo}"').replace(f"'{id_html}'", f"'{novo}'")
    return html


def _montar_blocos_monitoramento(caminho_base, dados):
    """Monta as peças reaproveitáveis do HTML "Ações em Monitoramento":
    cards, gráfico/avaliação dos 30 maiores municípios, grade de cards por
    município (com o detalhamento Município > Status > Objeto) e os dois
    quadros Status × Eixo/Objeto — junto do CSS e do JS que os fazem
    funcionar, tudo em texto bruto (sem nenhum isolamento/prefixo ainda).
    Usada tanto pela página autônoma (gerar_html_resumo_acoes, que monta
    essas peças exatamente como antes) quanto pelas páginas MON_* da
    Apresentação CGAPE (montar_partes_html_monitoramento, que embrulha e
    isola as mesmas peças) — pedido do usuário em 15/11/2026, pra
    incorporar o Monitoramento na apresentação do PAC sem duplicar a
    lógica de geração nem criar um terceiro gerador."""
    por_mun, _, dados_cards_acoes = _investimento_por_municipio(caminho_base, dados)

    # Lista = os 30 maiores municípios por POPULAÇÃO (RA_TOP30, já em ordem
    # de população), não mais os 15 maiores investimentos — pedido do usuário
    # em 13/11/2026. Município sem ação no recorte aparece zerado, sem "+".
    por_chave = {_chave_municipio(i["nome"]): i for i in por_mun.values()}
    ranking = [por_chave.get(_chave_municipio(nome))
               or {"nome": nome, "valor": 0.0, "qtd": 0, "orc": False, "eixos": [],
                   "n_eixos": 0, "valor_inicial": 0.0, "valor_avancado": 0.0, "acoes": []}
               for nome in RA_TOP30]
    valores_rank = sorted((m["valor"] for m in ranking), reverse=True)
    vmax = valores_rank[0] if valores_rank and valores_rank[0] > 0 else 1
    # um município muito acima dos demais (Salvador) achataria todas as
    # outras barras: a escala passa a ser a do 2º maior e a barra do maior
    # leva uma marca de quebra
    quebra = len(valores_rank) > 1 and valores_rank[1] > 0 and vmax > 3 * valores_rank[1]
    escala = valores_rank[1] * 1.15 if quebra else vmax

    indicadores, mediana_per_capita = _calcular_indicadores_carteira_top30(ranking)
    frases_leitura, paragrafo_leitura = _gerar_leitura_carteira(indicadores, mediana_per_capita)
    grafico_top30 = _svg_grafico_investimento_top30(
        indicadores, quebra, escala, vmax, mediana_per_capita)

    # detalhamento de cada município (Município > Status > Objeto) abre pelo
    # "+" (mesmo botão das seções do HTML do PAC), sem caixa de diálogo —
    # reaproveitado tanto pelos cards dos 30 maiores quanto, se o mapa for
    # religado (RA_HTML_MOSTRAR_MAPA), pela lista de ranking ao lado dele.
    # Cada status é um subtítulo com o seu próprio "+"; tudo fica oculto
    # enquanto o município estiver recolhido. Clicar no NOME do objeto abre
    # a janela de cards com cada ação daquele Município > Status > Objeto
    # (ver raObjClic no JS e _dados_cards_acoes_top30) — pedido do usuário
    # em 15/11/2026.
    def detalhe_rank(prefixo, m):
        chave_mun = _chave_municipio(m["nome"])
        blocos = []
        for j, g in enumerate(m["acoes"]):
            tabela = ('<table><thead><tr><th>Objeto</th><th class="r">Investimento</th>'
                      '</tr></thead><tbody>'
                      + "".join(
                          f'<tr><td><span class="ra-obj-clic" data-mun="{_esc(chave_mun)}" '
                          f'data-status="{_esc(g["status"])}" data-objeto="{_esc(o)}" '
                          f'title="Ver os cards das ações">{_esc(o)}</span>'
                          + (f"<small>{_esc(q)}</small>" if q else "") + "</td>"
                          f'<td class="r">{_esc(v)}</td></tr>' for o, q, v in g["objetos"])
                      + "</tbody></table>")
            blocos.append(
                f'<div class="rk-st"><div class="rk-st-linha"><button type="button" class="det det-st" '
                f'aria-expanded="false" aria-controls="rk-det-{prefixo}-{j}" '
                f'title="Ver objetos — {_esc(g["status"])}">+</button>'
                f'<span class="rk-st-nome">{_esc(g["status"])}<small>{_esc(g["qtd"])}</small></span>'
                f'<b>{_esc(g["valor"])}</b></div>'
                f'<div class="rk-st-obj oculto" id="rk-det-{prefixo}-{j}">{tabela}</div></div>')
        return "".join(blocos)

    # ---- Grade dos 30 maiores municípios — cards acima dos quadros, no
    # lugar da antiga lista ao lado do mapa (pedido do usuário em
    # 15/11/2026). Cada card abre, logo abaixo (ocupando a largura toda da
    # grade — ver .mun-det{{grid-column:1/-1}} no CSS), o detalhamento por
    # status (reaproveita detalhe_rank acima); município sem ação fica
    # esmaecido e não clicável.
    def cartao_municipio(i, m, it):
        selo = f'<span class="mun-selo" style="background:{it["cor_classe"]}">{_esc(it["classe"])}</span>'
        if not m["qtd"]:
            return (
                '<div class="mun-card sem-acao" aria-disabled="true">'
                f'<div class="mun-card-topo"><span class="mun-pos">{it["pos_pop"]}º em população</span></div>'
                f'<div class="mun-nome">{_esc(m["nome"])}</div>'
                '<div class="mun-vazio">sem ação no recorte</div></div>'
            )
        corpo = (
            f'<div class="mun-card-topo"><span class="mun-pos">{it["pos_pop"]}º em população</span>{selo}</div>'
            f'<div class="mun-nome">{_esc(m["nome"])}</div>'
            '<div class="mun-metricas">'
            f'<div><span class="mun-rot">INVESTIMENTO</span><b>{_esc(_ra_rs(m["valor"]))}'
            f'{"*" if m["orc"] else ""}</b></div>'
            f'<div><span class="mun-rot">AÇÕES</span><b>{m["qtd"]}</b></div>'
            '</div>'
        )
        return (
            f'<div class="mun-card" id="mun-card-{i}" role="button" tabindex="0" '
            f'aria-expanded="false" aria-controls="mun-det-{i}" '
            f'title="Ver detalhamento por status — {_esc(m["nome"])}">{corpo}</div>'
            f'<div class="rk-det mun-det oculto" id="mun-det-{i}">{detalhe_rank(f"m{i}", m)}</div>'
        )

    cartoes_mun = "".join(
        cartao_municipio(i, m, it) for i, (m, it) in enumerate(zip(ranking, indicadores)))

    # Metodologia dos indicadores, recolhida num <details> — pedido do
    # usuário em 15/11/2026 (nota pequena, não precisa estar sempre à vista).
    pesos_pct = {k: f"{v * 100:.0f}%" for k, v in RA_PESOS_NECESSIDADE.items()}
    nota_metodologia = (
        '<details class="leitura-metodologia"><summary>Metodologia do Índice de Necessidade de Investimento</summary>'
        '<p>Combinação ponderada e normalizada (mín-máx entre os 30) de: investimento per capita — '
        f'invertido, peso {pesos_pct["per_capita"]} (quanto menor o valor por habitante, maior a '
        'necessidade); gap de participação — % da população menos % do investimento dos 30, peso '
        f'{pesos_pct["gap"]}; fração do investimento ainda em status iniciais da carteira (em '
        'estudo, verificação de viabilidade, desenvolvimento de projetos, autorizo ou licitação), '
        f'peso {pesos_pct["maturidade"]}; e cobertura — nº de ações e de eixos atendidos, invertida, '
        f'peso {pesos_pct["cobertura"]}. Classificação por quartil do índice entre os 30: Prioridade '
        'alta (25% mais altos), Atenção (entre o 2º e o 3º quartil) e Adequado (os demais 50%). '
        f'População: {_esc(RA_TOP30_FONTE)}. A mediana per capita (linha pontilhada do gráfico) usa '
        'só os municípios com ação no recorte atual.</p></details>'
    )

    # Separadas em duas variáveis (gráfico/avaliação x grade de cards) —
    # pedido do usuário em 15/11/2026: na Apresentação CGAPE elas viram
    # páginas MON_* diferentes (MON_TOP30 e MON_MUNICIPIOS); concatenadas
    # as duas, dá exatamente o antigo "secao_top30" da página autônoma.
    secao_top30_grafico = (
        '<section class="sec-top30"><h2>INVESTIMENTO DOS 30 MAIORES MUNICÍPIOS (POPULAÇÃO)'
        '<small>passe o mouse (ou clique) numa coluna para ver o detalhe</small></h2>'
        '<div class="legenda-top30">'
        f'<span><i style="background:{RA_COR_PRIORIDADE_ALTA}"></i>Prioridade alta</span>'
        f'<span><i style="background:{RA_COR_ATENCAO}"></i>Atenção</span>'
        f'<span><i style="background:{RA_COR_ADEQUADO};opacity:.35"></i>Adequado</span>'
        '<span><i class="legenda-esperado"></i>investimento esperado pela população '
        '(mediana per capita dos 30)</span></div>'
        f'<div class="rolagem-grafico">{grafico_top30}</div>'
        '<div class="nota-mapa">Em ordem de população (IBGE, estimativa 2026).'
        + (' Barra do maior investimento interrompida (≈) para não achatar as demais.'
           if quebra else '') + '</div>'
        '<h3>LEITURA DA CARTEIRA</h3>'
        f'<ol class="leitura-lista">{frases_leitura}</ol>'
        f'<p class="leitura-paragrafo">{_esc(paragrafo_leitura)}</p>'
        f'{nota_metodologia}</section>'
    )
    secao_cartoes = (
        '<section class="sec-cartoes"><h2>OS 30 MAIORES MUNICÍPIOS (POPULAÇÃO)</h2>'
        f'<div class="mun-grade">{cartoes_mun}</div></section>'
    )

    # Mapa dos 30 maiores municípios colorido pela classe do Índice de
    # Necessidade (mesmas cores do gráfico; "Adequado" esmaecido) — religado
    # a pedido do usuário em 16/11/2026. O desenho (SVG) é um só para a página
    # toda (p["mapa_svg"], guardado em .ra-mapa-guarda): cada recorte só
    # traz as cores/dicas em JSON (.ra-mapa-dados) e o JS (ligarMapa) move o
    # SVG para o .mapa-slot do recorte visível e o pinta.
    secao_mapa, mapa_svg = "", ""
    if RA_HTML_MOSTRAR_MAPA:
        mapa_svg = _svg_mapa_base()
        cores_mapa, grupos = {}, {"Prioridade alta": [], "Atenção": [], "Adequado": []}
        sem_acao = []
        # último item = id do cartão do município (o mesmo data-alvo da
        # coluna do gráfico): clicar no mapa faz o mesmo que clicar na
        # coluna — pedido do usuário em 16/11/2026
        for i, it in enumerate(indicadores):
            k = _chave_municipio(it["nome"])
            if not it["qtd"]:
                cores_mapa[k] = [RA_COR_SEM_ACAO_TOP30, "sem ação no recorte", "", "", "", "", ""]
                sem_acao.append(it["nome"])
                continue
            cores_mapa[k] = [it["cor_classe"], it["classe"],
                             _ra_rs(it["valor"]) + ("*" if it["orc"] else ""),
                             _ra_qtd(it["qtd"]), f'{it["indice"]:.0f}',
                             "0.35" if it["classe"] == "Adequado" else "",
                             f"mun-card-{i}"]
            grupos[it["classe"]].append(it["nome"])
        cor_de = {"Prioridade alta": RA_COR_PRIORIDADE_ALTA, "Atenção": RA_COR_ATENCAO,
                  "Adequado": RA_COR_ADEQUADO}

        def item_legenda(rotulo, cor, nomes, opacidade=""):
            estilo = f"background:{cor}" + (f";opacity:{opacidade}" if opacidade else "")
            n = len(nomes)
            if not n:
                return ""
            return (f'<li><div class="leg-linha"><i style="{estilo}"></i>{_esc(rotulo)}'
                    f'<b>{n} município{"s" if n != 1 else ""}</b></div>'
                    + (f'<small>{_esc(", ".join(nomes))}</small>' if nomes else "") + "</li>")

        legenda = (item_legenda("Prioridade alta", cor_de["Prioridade alta"], grupos["Prioridade alta"])
                   + item_legenda("Atenção", cor_de["Atenção"], grupos["Atenção"])
                   + item_legenda("Adequado", cor_de["Adequado"], grupos["Adequado"], ".35")
                   + (item_legenda("Sem ação no recorte", RA_COR_SEM_ACAO_TOP30, sem_acao)
                      if sem_acao else "")
                   + f'<li><div class="leg-linha"><i style="background:{RA_COR_FORA_TOP30};'
                     'border:1px solid #D8DCE3"></i>Demais municípios (fora dos 30 maiores)</div></li>')
        json_mapa = json.dumps(cores_mapa, ensure_ascii=False,
                               separators=(",", ":")).replace("</", "<\\/")
        secao_mapa = (
            '<section class="sec-mapa"><h2>MAPA DOS 30 MAIORES MUNICÍPIOS — ÍNDICE DE NECESSIDADE'
            '<small>passe o mouse num município para ver o detalhe</small></h2>'
            '<div class="mapa-grade"><div><div class="mapa-slot"></div>'
            '<div class="nota-mapa">Cores pela classe do Índice de Necessidade de Investimento '
            '(ver metodologia abaixo do gráfico), calculado entre os 30 maiores municípios por população.</div></div>'
            f'<div><h3>CLASSIFICAÇÃO</h3><ul class="legenda legenda-prioridade">{legenda}</ul></div></div>'
            f'<script type="application/json" class="ra-mapa-dados">{json_mapa}</script></section>'
        )

    cards_html = "".join(
        f'<div><div class="card" style="--c:#{c["cor"]}"><span>{_esc(c["rotulo"])}</span>'
        f'<b>{_esc(c["txt"])}</b><small>{_esc(c["qtd"])}</small></div>'
        + (f'<div class="card-nota">{_esc(c["nota"])}</div>' if c.get("nota") else "")
        + "</div>"
        for c in dados["cards"])
    filtros_html = (f'<b>Recorte filtrado.</b> Filtros: {_esc(dados.get("filtros_txt") or "")}. '
                    if dados["filtrado"] else "<b>Base completa.</b> ")
    aviso_cards_html = (f'<div class="aviso-cards">{_esc(dados.get("aviso_cards"))}</div>'
                        if dados.get("aviso_cards") else "")
    # Recorte em texto puro (sem HTML), pro selo "AÇÕES EM MONITORAMENTO ·
    # base Ações do Cronograma · <recorte>" da Apresentação CGAPE — pedido
    # do usuário em 15/11/2026.
    recorte_txt = (f"Filtros: {dados.get('filtros_txt') or ''}"
                   if dados["filtrado"] else "Base completa")

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

    quadro_status_eixo_html = (
        '<section><h2>QUADRO RESUMO — STATUS × EIXO</h2>'
        f'<div class="rolagem">{quadro}</div>'
        f'<div class="nota">{_esc(dados["nota"])}</div></section>'
    )
    quadro_status_objeto_html = (
        '<section><h2>QUADRO RESUMO — STATUS × OBJETO '
        '<small>objetos agrupados pelo eixo, na mesma ordem do quadro acima</small></h2>'
        f'<div class="rolagem rolagem-obj">{quadro_obj}</div>'
        f'<div class="nota">{_esc(dados["nota"])}</div></section>'
    ) if quadro_obj else ""

    css_bruto = f"""/* Tokens alinhados à identidade visual única do projeto (ver
   identidade_visual.py) — --azul era #1F5C85 (divergente da referência) e
   os cinzas eram quase, mas não exatamente, os da referência — corrigido
   a pedido do usuário em 14/11/2026. */
:root{{--navy:#0B1F3F;--azul:#0B5CAD;--cinza:#5B6576;--linha:#D8DCE3;--claro:#F2F4F7;--fundo:#EDEFF3}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--fundo);color:#1A2433;font-family:'Roboto','Segoe UI',Arial,sans-serif}}
header{{background:var(--navy);color:#fff;padding:18px 36px;display:flex;justify-content:space-between;align-items:center;gap:20px}}
header h1{{margin:0;font-size:24px;font-weight:900;letter-spacing:.3px}}
header p{{margin:4px 0 0;color:#C9D6EA;font-size:12px;font-weight:600;letter-spacing:.4px}}
header .gov{{text-align:right;color:#4FA3E0;font-weight:700;font-size:13px;line-height:1.4}}
main{{max-width:1500px;margin:0 auto;padding:20px 36px 36px}}
.filtros{{font-size:13px;color:var(--cinza);margin-bottom:14px}}
.filtros b{{color:var(--navy)}}
.card-nota{{font-size:12px;font-style:italic;color:var(--cinza);margin-top:6px;padding-left:4px}}
.aviso-cards{{display:inline-block;background:#FDECD2;border:1px solid #E68A00;color:#7A4A00;border-radius:8px;padding:6px 12px;font-size:13px;font-weight:600;margin-bottom:10px}}
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
.mapa path.top30{{cursor:pointer}}
.mapa-rotulos text{{font-size:11px;font-weight:800;fill:#0B1F3F;paint-order:stroke;stroke:#fff;stroke-width:3px;stroke-linejoin:round;pointer-events:none;font-family:inherit}}
.legenda{{list-style:none;padding:0;margin:0 0 18px;font-size:13px}}
.legenda li{{display:flex;align-items:center;gap:8px;margin:5px 0}}
.legenda i{{width:16px;height:16px;border-radius:4px;flex:none}}
.legenda b{{margin-left:auto;color:var(--cinza);font-weight:600}}
.legenda-prioridade li{{display:block;margin:10px 0}}
.leg-linha{{display:flex;align-items:center;gap:8px;font-weight:600;color:var(--navy)}}
.legenda-prioridade small{{display:block;margin:3px 0 0 24px;color:var(--cinza);font-size:12.5px;line-height:1.45}}
.ra-mapa-guarda{{display:none}}
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
/* ---- Gráfico + avaliação dos 30 maiores municípios (substitui o mapa
   oculto — RA_HTML_MOSTRAR_MAPA — pedido do usuário em 15/11/2026) ---- */
.legenda-top30{{display:flex;flex-wrap:wrap;gap:16px;font-size:12.5px;color:var(--cinza);margin-bottom:10px}}
.legenda-top30 i{{width:14px;height:14px;border-radius:3px;display:inline-block;margin-right:5px;vertical-align:-2px}}
.legenda-top30 i.legenda-esperado{{background:none;border-top:2px dashed var(--navy);width:16px;height:0;vertical-align:4px}}
.rolagem-grafico{{overflow-x:auto}}
/* ocupa toda a largura da seção (até o limite do <main>) em telas largas —
   só cai pro mínimo de 920px, com rolagem horizontal, se a tela for estreita
   demais pra isso (pedido do usuário em 15/11/2026, aproveitar o espaço
   lateral que sobrava com a largura fixa) */
.grafico-top30{{width:100%;min-width:920px;height:auto;display:block}}
.grafico-grade{{stroke:#E3E8EF;stroke-width:1}}
.grafico-eixo{{stroke:var(--navy);stroke-width:1.2}}
.grafico-eixo-y,.grafico-eixo-x{{font-size:10.5px;font-family:inherit}}
.grafico-eixo-y{{fill:var(--cinza)}}
.grafico-eixo-x{{fill:var(--navy);text-anchor:start}}
.grafico-eixo-x.sem-acao{{fill:var(--cinza);font-style:italic}}
.grafico-rotulo{{font-size:9.5px;font-weight:800;font-family:inherit;paint-order:stroke;stroke:#fff;stroke-width:3px;stroke-linejoin:round}}
.grafico-rotulo-un{{font-size:8px;font-weight:700;font-family:inherit;paint-order:stroke;stroke:#fff;stroke-width:3px;stroke-linejoin:round}}
.grafico-quebra{{font-size:13px;font-weight:700;fill:var(--navy);font-family:inherit}}
.grafico-esperado{{stroke:var(--navy);stroke-width:1.6;stroke-dasharray:3 2;opacity:.55}}
.barra-top30{{cursor:pointer;transition:opacity .1s}}
.barra-top30:hover,.barra-top30:focus-visible{{opacity:.78;outline:none}}
.leitura-lista{{margin:10px 0 14px;padding-left:20px;font-size:13.5px;line-height:1.55}}
.leitura-lista li{{margin-bottom:6px}}
.leitura-lista b{{color:var(--navy)}}
.leitura-paragrafo{{font-size:13.5px;color:var(--cinza);line-height:1.55;margin:0 0 10px}}
.leitura-metodologia{{font-size:12px;color:var(--cinza)}}
.leitura-metodologia summary{{cursor:pointer;color:var(--navy);font-weight:600}}
.leitura-metodologia p{{margin:8px 0 0}}
.mun-grade{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}}
@media (max-width:900px){{.mun-grade{{grid-template-columns:repeat(2,1fr)}}}}
@media (max-width:560px){{.mun-grade{{grid-template-columns:1fr}}}}
.mun-card{{background:#fff;border:1px solid var(--linha);border-top:4px solid var(--azul);border-radius:10px;padding:14px 16px;box-shadow:0 2px 10px rgba(11,31,63,.06);cursor:pointer;transition:box-shadow .15s}}
.mun-card:hover,.mun-card:focus-visible{{box-shadow:0 6px 18px rgba(11,31,63,.16);outline:none}}
.mun-card.aberto{{border-top-color:var(--navy)}}
.mun-card.sem-acao{{opacity:.55;cursor:default;border-top-color:#D8DCE3}}
.mun-card-topo{{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:6px}}
.mun-pos{{font-size:11px;color:var(--cinza);font-weight:600}}
.mun-selo{{font-size:10.5px;font-weight:700;color:#fff;padding:2px 9px;border-radius:10px;white-space:nowrap}}
.mun-nome{{font-size:15.5px;font-weight:700;color:var(--navy);margin-bottom:8px}}
.mun-metricas{{display:flex;gap:20px}}
.mun-rot{{display:block;font-size:10px;color:var(--cinza);letter-spacing:.4px;margin-bottom:2px}}
.mun-metricas b{{font-size:18px;color:var(--azul);font-weight:800}}
.mun-vazio{{font-size:12.5px;color:var(--cinza);font-style:italic;margin-top:4px}}
.mun-det{{grid-column:1/-1}}
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
/* Cor fixa (não var(--navy)): na Apresentação CGAPE o tooltip (#mon-dica)
   fica DE PROPÓSITO fora de .mon-ra, então não herdaria a variável — ficava
   com fundo transparente, texto branco ilegível sobre o gráfico. Bug
   encontrado pelo usuário em 15/11/2026. */
#dica{{position:fixed;pointer-events:none;background:#0B1F3F;color:#fff;border-radius:8px;padding:9px 12px;font-size:12.5px;box-shadow:0 6px 20px rgba(0,0,0,.25);display:none;max-width:280px;z-index:5}}
#dica b{{display:block;font-size:13.5px;margin-bottom:2px}}
#dica .v{{color:#9FD3FF;font-weight:700}}
#dica ul{{list-style:none;padding:6px 0 0;margin:6px 0 0;border-top:1px solid rgba(255,255,255,.2)}}
#dica li{{display:flex;justify-content:space-between;gap:12px;color:#C9D6EA}}
@media (max-width:900px){{.mapa-grade{{grid-template-columns:1fr}}.cards{{grid-template-columns:1fr 1fr}}main,header{{padding-left:16px;padding-right:16px}}}}
@media print{{body{{background:#fff}}section{{box-shadow:none;break-inside:avoid}}#dica{{display:none!important}}}}
.btn-pdf{{position:fixed;top:18px;right:24px;z-index:20;padding:11px 22px;font:700 14px "Segoe UI",Roboto,system-ui,sans-serif;color:#fff;background:#0B5CAD;border:none;border-radius:8px;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.3)}}
.btn-pdf:hover{{background:#0A4E92}}
@media print{{.btn-pdf,.ranking .det,.mun-det .det{{display:none!important}}.rk-det{{max-height:none;overflow:visible}}section.sec-mapa{{break-inside:auto}}.rk-det tr{{break-inside:avoid}}
.mapa-grade{{grid-template-columns:minmax(0,1fr) minmax(0,1.15fr);gap:18px}}.mapa{{max-width:330px;margin:0}}.mapa-grade>div:first-child{{break-inside:avoid}}
.ranking .rk-linha{{grid-template-columns:18px minmax(0,1.9fr) minmax(0,.7fr) 96px;gap:6px}}.ranking .nm small{{margin-left:4px}}
/* Gráfico/avaliação numa página só, cards com o status (nível 1) já
   aberto e os objetos (nível 2) escondidos — pedido do usuário em
   15/11/2026. Força a visibilidade por CSS, independente da classe
   "oculto" que o JS controla na tela. */
.sec-top30,.sec-cartoes{{break-inside:avoid}}
.rolagem-grafico{{overflow:visible}}
.mun-card{{break-inside:avoid;cursor:default}}
.mun-det.oculto{{display:block!important}}
.mun-det .rk-st-obj{{display:none!important}}}}
/* Janela de cards ao clicar num objeto da grade de município (ver
   detalhe_rank) — pedido do usuário em 15/11/2026. Tudo dentro dela
   prefixado por "#ra-oc-fundo " (como o tooltip #dica) pra ficar de fora
   do escopo .mon-ra quando embutida na Apresentação CGAPE: a janela
   precisa ficar FORA de qualquer slide (ver oc_html em
   montar_partes_html_monitoramento), senão o overflow:hidden do slide a
   recortaria. ".ra-obj-clic" fica de fora dessa regra de propósito: o
   objeto clicável é parte do conteúdo normal (dentro de .mon-ra).*/
.ra-obj-clic{{cursor:pointer;color:var(--azul);text-decoration:underline dotted;text-underline-offset:2px}}
.ra-obj-clic:hover{{color:var(--navy);text-decoration-style:solid}}
#ra-oc-fundo{{display:none;position:fixed;inset:0;z-index:1150;background:rgba(11,31,63,.55);align-items:center;justify-content:center;padding:24px}}
#ra-oc-fundo.aberta{{display:flex}}
#ra-oc-fundo .ra-oc-janela{{background:#fff;border-radius:12px;width:min(1200px,100%);max-height:100%;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.35)}}
#ra-oc-fundo .ra-oc-topo{{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;padding:20px 24px 14px;border-bottom:1px solid var(--linha)}}
#ra-oc-fundo .ra-oc-titulo{{font-size:19px;font-weight:800;color:var(--navy)}}
#ra-oc-fundo .ra-oc-sub{{font-size:13px;color:var(--cinza);margin-top:4px}}
#ra-oc-fundo .ra-oc-fechar{{background:none;border:none;font-size:26px;line-height:1;color:var(--cinza);cursor:pointer;padding:0 4px}}
#ra-oc-fundo .ra-oc-fechar:hover{{color:var(--navy)}}
/* Cards no MESMO desenho dos cards de ação do HTML do PAC (.entrega-card
   em gerar_pptx_pac.py): velocímetro de 76px que nunca encolhe (a descrição
   não interfere mais no tamanho dele) e descrição em até 3 linhas — pedido
   do usuário em 16/11/2026. */
#ra-oc-fundo .ra-oc-corpo{{padding:16px 24px 24px;overflow-y:auto;display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:16px}}
#ra-oc-fundo .ra-oc-card{{background:#fff;border:1px solid #D8DCE3;border-top:4px solid #0B5CAD;border-radius:10px;padding:16px 18px;display:flex;flex-direction:column;height:100%}}
#ra-oc-fundo .ra-oc-card-topo{{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}}
#ra-oc-fundo .ra-oc-card-id{{min-width:0}}
#ra-oc-fundo .ra-oc-card-objeto{{font-size:15px;font-weight:700;color:#0B1F3F;line-height:1.25}}
#ra-oc-fundo .ra-oc-card-mun{{font-size:13px;color:#5B6576;margin-top:2px}}
#ra-oc-fundo .ra-oc-card-ex{{font-size:12px;color:#8C8C8C;margin-top:3px}}
#ra-oc-fundo .ra-oc-card-gauge{{flex-shrink:0;width:76px;height:76px}}
#ra-oc-fundo .ra-oc-card-desc{{font-size:13px;color:#0B1F3F;line-height:1.35;margin-top:12px;margin-bottom:12px;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}}
#ra-oc-fundo .ra-oc-card-rodape{{display:flex;justify-content:space-between;gap:10px;margin-top:auto;border-top:1px solid #F0F2F6;padding-top:10px}}
#ra-oc-fundo .ra-oc-card-rotulo{{font-size:10.5px;font-weight:700;color:#5B6576;text-transform:uppercase}}
#ra-oc-fundo .ra-oc-card-valor{{font-size:14.5px;font-weight:700;color:#0B1F3F;margin-top:2px}}
#ra-oc-fundo .ra-oc-card-inv{{font-size:18px;font-weight:900;color:#0B5CAD;white-space:nowrap}}
#ra-oc-fundo .ra-oc-sem-gauge{{width:76px;height:76px;display:flex;align-items:center;justify-content:center;text-align:center;font-size:10.5px;font-weight:700;color:#8C8C8C;line-height:1.3}}
#ra-oc-fundo .ra-oc-vazio{{font-size:13px;color:var(--cinza);padding:20px 4px}}
@media print{{#ra-oc-fundo{{display:none!important}}.ra-obj-clic{{color:inherit;text-decoration:none}}}}
"""

    agora = datetime.now()
    titulo = "AÇÕES EM MONITORAMENTO"
    js_bruto = f"""(function(){{
  // Tudo que liga eventos fica em ligar(raiz): a página autônoma troca o
  // recorte (eixo/executor) trocando o HTML e chama window.raLigar de novo
  // só no conteúdo novo — pedido do usuário em 16/11/2026.
  function todos(raiz,sel){{return Array.prototype.slice.call(raiz.querySelectorAll(sel));}}
  // "+" / "–" de cada status/objeto: abre/fecha o detalhamento no lugar.
  function alternar(b,abrir){{
    var alvo=document.getElementById(b.getAttribute("aria-controls"));
    if(abrir===undefined)abrir=alvo.classList.contains("oculto");
    alvo.classList.toggle("oculto",!abrir);
    b.textContent=abrir?"–":"+";
    b.setAttribute("aria-expanded",abrir?"true":"false");
    b.title=abrir?b.title.replace("Ver","Ocultar"):b.title.replace("Ocultar","Ver");
  }}

  // Dica (tooltip) ao passar o mouse — mapa e colunas do gráfico.
  var dica=document.getElementById("dica");
  function mostrarDica(html,ev,elemento){{
    dica.innerHTML=html;
    dica.style.display="block";
    var r=dica.getBoundingClientRect();
    var x,y;
    if(ev){{
      x=ev.clientX+14;y=ev.clientY+14;
    }} else {{
      var re=elemento.getBoundingClientRect();
      x=re.left;y=re.bottom+8;
    }}
    if(x+r.width>innerWidth-8)x=innerWidth-r.width-8;
    if(y+r.height>innerHeight-8)y=innerHeight-r.height-8;
    dica.style.left=Math.max(4,x)+"px";dica.style.top=Math.max(4,y)+"px";
  }}
  function linhaDica(rotulo,valor){{
    return valor?"<li><span>"+rotulo+"</span><span>"+valor+"</span></li>":"";
  }}

  // Mapa: um SVG só para a página (guardado em .ra-mapa-guarda); vai para o
  // .mapa-slot do recorte visível e é pintado com as cores dele
  // (.ra-mapa-dados: chave -> [cor, classe, investimento, ações, índice, opacidade]).
  var mapaM={{}};
  function ligarMapa(raiz){{
    var slot=raiz.querySelector(".mapa-slot");
    var svg=document.querySelector(".ra-mapa-guarda svg")||document.querySelector(".mapa-slot svg");
    if(!slot||!svg)return;
    if(svg.parentNode!==slot)slot.appendChild(svg);
    var el=raiz.querySelector(".ra-mapa-dados");
    mapaM=el?JSON.parse(el.textContent):{{}};
    todos(svg,"path").forEach(function(p){{
      var d=mapaM[p.dataset.k];
      p.setAttribute("fill",d?d[0]:"{RA_COR_FORA_TOP30}");
      if(d&&d[5])p.setAttribute("fill-opacity",d[5]);else p.removeAttribute("fill-opacity");
      p.classList.toggle("top30",!!d);
    }});
    // nome do município no mapa só nos de Prioridade alta do recorte
    // visível — pedido do usuário em 16/11/2026
    var antigo=svg.querySelector(".mapa-rotulos");
    if(antigo)antigo.parentNode.removeChild(antigo);
    var NS="http://www.w3.org/2000/svg",g=document.createElementNS(NS,"g");
    g.setAttribute("class","mapa-rotulos");
    // caixas estimadas (largura ~6,2 por letra, 13 de altura): rótulo que
    // bateria em outro já posto desce uma linha (municípios vizinhos, ex.:
    // Candeias e Dias d'Ávila)
    var postos=[];
    todos(svg,"path").forEach(function(p){{
      var d=mapaM[p.dataset.k];
      if(!d||d[1]!=="Prioridade alta"||!p.dataset.cx)return;
      var x=+p.dataset.cx,y=+p.dataset.cy,w=p.dataset.n.length*6.2,tent=0;
      function bate(){{return postos.some(function(b){{return Math.abs(b.x-x)<(b.w+w)/2&&Math.abs(b.y-y)<13;}});}}
      while(bate()&&tent<6){{y+=13;tent++;}}
      postos.push({{x:x,y:y,w:w}});
      var t=document.createElementNS(NS,"text");
      t.setAttribute("x",x);t.setAttribute("y",y);
      t.setAttribute("text-anchor","middle");t.setAttribute("dy","0.35em");
      t.textContent=p.dataset.n;
      g.appendChild(t);
    }});
    svg.appendChild(g);
    if(svg.dataset.ligado)return;
    svg.dataset.ligado="1";
    svg.addEventListener("mousemove",function(ev){{
      var p=ev.target.closest("path");
      if(!p){{dica.style.display="none";return;}}
      var d=mapaM[p.dataset.k];
      mostrarDica("<b>"+p.dataset.n+"</b>"+(d?(d[2]
          ?"<ul>"+linhaDica("Classe",d[1])+linhaDica("Investimento",d[2])+linhaDica("Ações",d[3])+linhaDica("Índice de necessidade",d[4])+"</ul>"
          :"sem ação no recorte"):"fora dos 30 maiores municípios"),ev);
    }});
    svg.addEventListener("mouseleave",function(){{dica.style.display="none";}});
    // clique num dos 30 maiores = clique na coluna dele no gráfico (rola até
    // o cartão do município e o abre)
    svg.addEventListener("click",function(ev){{
      var p=ev.target.closest("path"),d=p&&mapaM[p.dataset.k];
      if(!d||!d[6])return;
      var barra=document.querySelector('.barra-top30[data-alvo="'+d[6]+'"]');
      if(barra){{dica.style.display="none";barra.dispatchEvent(new MouseEvent("click",{{bubbles:true}}));}}
    }});
  }}

  function ligar(raiz){{
    todos(raiz,".det").forEach(function(b){{
      b.addEventListener("click",function(){{alternar(b);}});
    }});

    // Cartão de cada um dos 30 maiores municípios: abre logo abaixo o
    // detalhamento por status; só um cartão aberto por vez.
    var cartoes=todos(raiz,".mun-card:not(.sem-acao)");
    function alternarCartao(cartao){{
      var det=document.getElementById(cartao.getAttribute("aria-controls"));
      var vaiAbrir=det.classList.contains("oculto");
      cartoes.forEach(function(outro){{
        document.getElementById(outro.getAttribute("aria-controls")).classList.add("oculto");
        outro.classList.remove("aberto");
        outro.setAttribute("aria-expanded","false");
      }});
      det.classList.toggle("oculto",!vaiAbrir);
      cartao.classList.toggle("aberto",vaiAbrir);
      cartao.setAttribute("aria-expanded",vaiAbrir?"true":"false");
    }}
    cartoes.forEach(function(cartao){{
      cartao.addEventListener("click",function(){{alternarCartao(cartao);}});
      cartao.addEventListener("keydown",function(ev){{
        if(ev.key==="Enter"||ev.key===" "){{ev.preventDefault();alternarCartao(cartao);}}
      }});
    }});

    // Clicar numa coluna do gráfico rola até o cartão do município e o abre.
    todos(raiz,".barra-top30[data-alvo]").forEach(function(barra){{
      function irParaCartao(){{
        var cartao=document.getElementById(barra.dataset.alvo);
        if(!cartao||cartao.classList.contains("sem-acao"))return;
        cartao.scrollIntoView({{behavior:"smooth",block:"center"}});
        if(!cartao.classList.contains("aberto"))alternarCartao(cartao);
      }}
      barra.addEventListener("click",irParaCartao);
      barra.addEventListener("keydown",function(ev){{
        if(ev.key==="Enter"||ev.key===" "){{ev.preventDefault();irParaCartao();}}
      }});
    }});

    todos(raiz,".barra-top30").forEach(function(b){{
      function html(){{
        return "<b>"+b.dataset.n+"</b><ul>"
          +linhaDica("Posição (população)",b.dataset.pos+"º")
          +linhaDica("Investimento",b.dataset.v)
          +linhaDica("Ações",b.dataset.q)
          +linhaDica("Per capita",b.dataset.pc)
          +linhaDica("Índice de necessidade",b.dataset.idx+" ("+b.dataset.classe+")")
          +"</ul>";
      }}
      b.addEventListener("mousemove",function(ev){{mostrarDica(html(),ev);}});
      b.addEventListener("mouseleave",function(){{dica.style.display="none";}});
      b.addEventListener("focus",function(){{mostrarDica(html(),null,b);}});
      b.addEventListener("blur",function(){{dica.style.display="none";}});
    }});
    ligarMapa(raiz);
  }}
  window.raLigar=ligar;
  ligar(document);

  // Imprimir/salvar em PDF: o CSS de @media print já abre o que precisa.
  window.gerarPdf=function(){{window.print();}};
}})();
// Janela de cards ao clicar num objeto da grade de município (ver
// detalhe_rank/_dados_cards_acoes_top30) — pedido do usuário em 15/11/2026.
(function(){{
  // Dados das ações: o bloco JSON "ra-oc-dados" do recorte visível (ou o
  // indicado em data-oc pelo contêiner).
  var fundo=document.getElementById("ra-oc-fundo");
  if(!fundo)return;
  var D=null,C={{}};
  function carregar(id){{
    var el=document.getElementById(id);
    D=el?JSON.parse(el.textContent):{{campos:[],acoes:[]}};C={{}};
    (D.campos||[]).forEach(function(k,i){{C[k]=i;}});
  }}
  function esc(t){{var d=document.createElement("div");d.textContent=t==null?"":String(t);return d.innerHTML;}}
  // velocímetro igual ao do PAC (arco de 250°, 76px; ver #pac-gauge em
  // gerar_pptx_pac.py), desenhado aqui mesmo (sem <use>)
  var ARCO='M 11.74 56.39 A 32.06 32.06 0 1 1 64.26 56.39';
  function gauge(v){{
    if(v==null)return '<div class="ra-oc-sem-gauge">Avanço<br>não<br>informado</div>';
    v=Math.max(0,Math.min(100,v));
    return '<svg width="76" height="76" viewBox="0 0 76 76" class="ra-oc-card-gauge">'
      +'<path d="'+ARCO+'" fill="none" stroke="#E9ECF2" stroke-width="9.88" stroke-linecap="round" pathLength="100"></path>'
      +(v>0.5?'<path d="'+ARCO+'" fill="none" stroke="#2E7D32" stroke-width="9.88" stroke-linecap="round" pathLength="100" stroke-dasharray="'+v.toFixed(1)+' 200"></path>':'')
      +'<text x="38" y="44.92" text-anchor="middle" font-size="19.8" font-weight="900" fill="#0B1F3F">'+Math.round(v)+'%</text></svg>';
  }}
  function card(a){{
    return '<div class="ra-oc-card"><div class="ra-oc-card-topo"><div class="ra-oc-card-id">'
      +'<div class="ra-oc-card-objeto">'+esc(a[C.nome])+'</div>'
      +'<div class="ra-oc-card-mun">'+esc(a[C.mun]||'Município não informado')+'</div>'
      +(a[C.ex]?'<div class="ra-oc-card-ex">'+esc(a[C.ex])+'</div>':'')
      +'</div>'+gauge(a[C.av])+'</div>'
      +(a[C.ds]?'<div class="ra-oc-card-desc" title="'+esc(a[C.ds])+'">'+esc(a[C.ds])+'</div>':'')
      +'<div class="ra-oc-card-rodape">'
      +'<div><div class="ra-oc-card-rotulo">Conclusão prevista</div><div class="ra-oc-card-valor">'+esc(a[C.pz]||'A definir')+'</div></div>'
      +'<div><div class="ra-oc-card-rotulo">Investimento</div><div class="ra-oc-card-inv">'+esc(a[C.val])+'</div></div>'
      +'</div></div>';
  }}
  window.raOcFechar=function(){{if(fundo)fundo.classList.remove("aberta");}};
  document.addEventListener("click",function(ev){{
    var alvo=ev.target.closest(".ra-obj-clic");
    if(!alvo)return;
    var var_el=alvo.closest("[data-oc]");
    carregar(var_el?var_el.dataset.oc:"ra-oc-dados");
    var mun=alvo.dataset.mun,status=alvo.dataset.status,objeto=alvo.dataset.objeto;
    var lista=D.acoes.filter(function(a){{
      return a[C.mu]===mun&&a[C.st]===status&&a[C.ob]===objeto;
    }});
    document.getElementById("ra-oc-titulo").textContent=objeto;
    document.getElementById("ra-oc-sub").textContent=[status,lista.length+(lista.length===1?" ação":" ações")].filter(Boolean).join(" · ");
    document.getElementById("ra-oc-corpo").innerHTML=lista.map(card).join("")
      ||'<div class="ra-oc-vazio">Nenhuma ação encontrada.</div>';
    fundo.classList.add("aberta");
  }});
  if(fundo)fundo.addEventListener("click",function(ev){{if(ev.target===fundo)window.raOcFechar();}});
  document.addEventListener("keydown",function(e){{
    if(!fundo||!fundo.classList.contains("aberta"))return;
    if(e.key==="Escape")window.raOcFechar();
  }});
}})();
"""

    # Janela de cards ao clicar num objeto (ver detalhe_rank/.ra-obj-clic):
    # fica FORA de <main>, igual ao tooltip #dica, pelo mesmo motivo (ver
    # oc_html em montar_partes_html_monitoramento) — pedido do usuário em
    # 15/11/2026.
    oc_dados_json = json.dumps(dados_cards_acoes, ensure_ascii=False,
                               separators=(",", ":")).replace("</", "<\\/")
    oc_html = (
        f'<script type="application/json" id="ra-oc-dados">{oc_dados_json}</script>'
        '<div id="ra-oc-fundo">'
        '<div class="ra-oc-janela" role="dialog" aria-modal="true">'
        '<div class="ra-oc-topo"><div><div id="ra-oc-titulo" class="ra-oc-titulo"></div>'
        '<div id="ra-oc-sub" class="ra-oc-sub"></div></div>'
        '<button type="button" class="ra-oc-fechar" onclick="raOcFechar()" title="Fechar (Esc)">&times;</button></div>'
        '<div id="ra-oc-corpo" class="ra-oc-corpo"></div></div></div>'
    )

    return {
        "cards_html": cards_html,
        "aviso_cards_html": aviso_cards_html,
        "filtros_html": filtros_html,
        "secao_mapa": secao_mapa,
        "mapa_svg": mapa_svg,
        "secao_top30_grafico": secao_top30_grafico,
        "secao_cartoes": secao_cartoes,
        "quadro_status_eixo_html": quadro_status_eixo_html,
        "quadro_status_objeto_html": quadro_status_objeto_html,
        "css_bruto": css_bruto,
        "js_bruto": js_bruto,
        "oc_html": oc_html,
        "recorte_txt": recorte_txt,
        "agora": agora,
        "titulo": titulo,
    }


CSS_SELETOR_RECORTE = """
.sel-recorte{position:fixed;left:50%;transform:translateX(-50%);bottom:10px;z-index:1000;display:flex;flex-direction:column;align-items:center;gap:6px;width:max-content;max-width:94vw;opacity:0;pointer-events:none;transition:opacity .25s}
body.mostra-rodape .sel-recorte{opacity:1;pointer-events:auto}
@media (hover:none){.sel-recorte{opacity:1;pointer-events:auto}}
.sel-barra{display:flex;gap:4px;flex-wrap:wrap;justify-content:center;background:rgba(11,31,63,.92);padding:5px;border-radius:10px;box-shadow:0 4px 16px rgba(0,0,0,.3)}
.sel-barra button{background:transparent;color:#C9D6EA;border:none;border-radius:7px;padding:7px 12px;font:700 12px 'Roboto','Segoe UI',Arial,sans-serif;letter-spacing:.3px;cursor:pointer}
.sel-barra button:hover{background:rgba(255,255,255,.12);color:#fff}
.sel-barra button.oculto{display:none}
.sel-barra-eixo button{padding:8px 14px;font-size:13px}
.sel-barra-eixo button.ativo{background:#4FA3E0;color:#0B1F3F}
.sel-barra-sec button.ativo{background:#E68A00;color:#fff}
@media print{.sel-recorte{display:none!important}}
"""

JS_SELETOR_RECORTE = """
(function(){
  // Seletor EIXO x SECRETARIA/EXECUTOR (pedido do usuário em 16/11/2026),
  // no mesmo desenho dos seletores de Gestão/Secretaria do PAC: barras
  // flutuantes no rodapé, que aparecem com o mouse perto da borda de baixo
  // (body.mostra-rodape). Cada combinação com ação é o conteúdo inteiro já
  // calculado, guardado compactado (gzip + base64) num <script class="ra-v"
  // data-k="E|X">; ao escolher, é descompactado, posto no lugar e os
  // eventos ligados de novo (window.raLigar).
  function bytes(s){var b=atob(s),u=new Uint8Array(b.length);for(var i=0;i<b.length;i++)u[i]=b.charCodeAt(i);return u;}
  function esconderDica(){["dica","mon-dica"].forEach(function(id){var d=document.getElementById(id);if(d)d.style.display="none";});}
  var raiz=document.querySelector(".ra-sel-raiz"),barra=document.querySelector(".sel-recorte");
  if(!raiz||!barra)return;
  var lista=raiz.querySelector(".ra-v-lista");
  if(!lista)return;
  var V=JSON.parse(lista.textContent),cache={},atual="0|0",e=0,x=0;
  var cont=raiz.querySelector(".ra-recorte");
  var bEixo=Array.prototype.slice.call(barra.querySelectorAll(".sel-eixo-btn"));
  var bSec=Array.prototype.slice.call(barra.querySelectorAll(".sel-sec-btn"));
  // página autônoma: mesma régua de proximidade do rodapé da apresentação
  // (lá o próprio carrossel já liga/desliga body.mostra-rodape)
  if(!document.getElementById("nav-carrossel")){
    document.addEventListener("mousemove",function(ev){document.body.classList.toggle("mostra-rodape",ev.clientY>=window.innerHeight-150);});
    document.addEventListener("mouseleave",function(){document.body.classList.remove("mostra-rodape");});
  }
  function abrir(k){
    if(cache[k])return Promise.resolve(cache[k]);
    var el=raiz.querySelector('.ra-v[data-k="'+k+'"]');
    var fluxo=new Blob([bytes(el.textContent.trim())]).stream().pipeThrough(new DecompressionStream("gzip"));
    return new Response(fluxo).text().then(function(h){cache[k]=h;return h;});
  }
  function controles(){
    bEixo.forEach(function(b){b.classList.toggle("ativo",+b.dataset.e===e);});
    // secretaria sem ação no eixo escolhido some da barra (igual ao PAC)
    bSec.forEach(function(b){
      var i=+b.dataset.x;
      b.classList.toggle("ativo",i===x);
      b.classList.toggle("oculto",i!==0&&V.indexOf(e+"|"+i)<0);
    });
  }
  function mostrar(ne,nx){
    if(V.indexOf(ne+"|"+nx)<0)nx=0;
    var k=ne+"|"+nx;
    e=ne;x=nx;controles();
    if(k===atual)return;
    if(!window.DecompressionStream){alert("Este navegador não consegue abrir os recortes. Use o Edge ou o Chrome atualizados.");return;}
    document.body.style.cursor="progress";
    abrir(k).then(function(h){
      var svg=cont.querySelector(".mapa-slot svg"),guarda=document.querySelector(".ra-mapa-guarda");
      if(svg&&guarda)guarda.appendChild(svg);
      cont.innerHTML=h;atual=k;
      window.raLigar(cont);
      esconderDica();
      document.body.style.cursor="";
    }).catch(function(err){document.body.style.cursor="";alert("Não consegui abrir o recorte: "+err);});
  }
  bEixo.forEach(function(b){b.addEventListener("click",function(){mostrar(+b.dataset.e,x);});});
  bSec.forEach(function(b){b.addEventListener("click",function(){mostrar(e,+b.dataset.x);});});
  controles();
})();
"""


def _variantes_recorte(caminho_base, dados, filtros):
    """Combinações EIXO x SECRETARIA/EXECUTOR do recorte para o seletor —
    pedido do usuário em 16/11/2026. Devolve (eixos, executores,
    combinações) com eixos = ["TODOS OS EIXOS", ...], executores =
    ["TODAS AS SECRETARIAS / EXECUTORES", ...] e combinações = [(i_eixo,
    i_executor, filtros)] só das que têm ação; o índice 0 é sempre "todos"."""
    eixos = ["TODOS OS EIXOS"]
    executores = ["TODOS"]
    if filtros is None:
        return eixos, executores, [(0, 0, None)]
    base = _ra_base(caminho_base)
    linhas = base.iloc[dados["ids"]]
    if "_f_executor" not in linhas.columns:
        linhas = linhas.assign(_f_executor="")
    eixos += [e for e, c in zip(dados["eixos"], dados["total"]["celulas"]) if c]
    executores += sorted(set(linhas["_f_executor"]) - {""}, key=_ra_ordem)
    pares = set(zip(linhas["_eixo"], linhas["_f_executor"]))
    eixos_com = {e for e, _ in pares}
    execs_com = {x for _, x in pares}
    combos = []
    for ie, e in enumerate(eixos):
        for ix, x in enumerate(executores):
            if (ie and e not in eixos_com) or (ix and x not in execs_com):
                continue
            if ie and ix and (e, x) not in pares:
                continue
            f = dict(filtros)
            if ie:
                f["eixo"] = [e]
            if ix:
                f["executor"] = [x]
            combos.append((ie, ix, f))
    if len(eixos) < 3 and len(executores) < 3:
        return eixos, executores, [(0, 0, None)]
    return eixos, executores, combos


def _montar_recortes(caminho_base, dados, filtros, sufixo=None):
    """(html, seletor, p) — conteúdo do recorte inicial + as demais
    combinações compactadas (dentro de .ra-sel-raiz), as barras flutuantes
    EIXO x SECRETARIA/EXECUTOR (.sel-recorte, fora do conteúdo: ficam
    fixas no rodapé — ver JS_SELETOR_RECORTE) e p = peças do recorte
    inicial (CSS/JS/janela/mapa). `sufixo`: ids de cada recorte sufixados (página
    embutida na Apresentação CGAPE — ver _sufixar_ids_monitoramento)."""
    eixos, executores, combos = _variantes_recorte(caminho_base, dados, filtros)
    p, inicial, guardados, chaves = None, "", [], []
    for ie, ix, f in combos:
        dv = dados if (ie, ix) == (0, 0) else resumo_acoes_dados(caminho_base, f)
        if dv["vazio"]:
            continue
        pv = _montar_blocos_monitoramento(caminho_base, dv)
        if p is None:
            p = pv
        json_oc, _, _ = pv["oc_html"].partition('<div id="ra-oc-fundo">')
        corpo = (
            f'<div class="filtros">{pv["filtros_html"]}'
            + (f'Gerado em {pv["agora"]:%d/%m/%Y às %Hh%M}.' if not sufixo else "") + '</div>'
            f'{pv["aviso_cards_html"]}'
            f'<div class="cards">{pv["cards_html"]}</div>'
            f'{pv["secao_mapa"]}{pv["secao_top30_grafico"]}{pv["secao_cartoes"]}'
            f'{pv["quadro_status_eixo_html"]}{pv["quadro_status_objeto_html"]}{json_oc}')
        if sufixo:
            corpo = _sufixar_ids_monitoramento(corpo, sufixo)
        if (ie, ix) == (0, 0):
            inicial = corpo
        if len(combos) > 1:
            chaves.append(f"{ie}|{ix}")
            dado = base64.b64encode(gzip.compress(corpo.encode("utf-8"), 9)).decode("ascii")
            guardados.append(f'<script type="application/octet-stream" class="ra-v" data-k="{ie}|{ix}">{dado}</script>')
    data_oc = "ra-oc-dados" + (f"--{sufixo}" if sufixo else "")
    seletor = ""
    if len(combos) > 1:
        b_eixo = "".join(
            f'<button type="button" class="sel-eixo-btn{" ativo" if i == 0 else ""}" data-e="{i}">'
            f'{_esc(e)}</button>' for i, e in enumerate(eixos))
        b_sec = "".join(
            f'<button type="button" class="sel-sec-btn{" ativo" if i == 0 else ""}" data-x="{i}">'
            f'{_esc("Todas as secretarias" if i == 0 else x)}</button>'
            for i, x in enumerate(executores))
        seletor = ('<div class="sel-recorte" role="group" aria-label="Eixo e secretaria/executor">'
                   f'<div class="sel-barra sel-barra-sec">{b_sec}</div>'
                   f'<div class="sel-barra sel-barra-eixo">{b_eixo}</div></div>')
        guardados.append('<script type="application/json" class="ra-v-lista">'
                         + json.dumps(chaves) + "</script>")
    html = (f'<div class="ra-sel-raiz">'
            f'<div class="ra-recorte" data-oc="{data_oc}">{inicial}</div>'
            f'{"".join(guardados)}</div>')
    return html, seletor, p


def gerar_html_resumo_acoes(caminho_base, dados, filtros=None):
    """Página única (abre sem internet) com cards, mapa e gráfico/avaliação
    dos 30 maiores municípios por população, cartões por município e os
    quadros Status × Eixo/Objeto do recorte atual. Com `filtros`, ganha no
    topo o seletor EIXO (botões) x SECRETARIA/EXECUTOR (lista)."""
    html_recortes, seletor, p = _montar_recortes(caminho_base, dados, filtros)
    fundo_oc = '<div id="ra-oc-fundo">' + p["oc_html"].partition('<div id="ra-oc-fundo">')[2]
    guarda_mapa = f'<div class="ra-mapa-guarda">{p["mapa_svg"]}</div>' if p.get("mapa_svg") else ""
    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ações em Monitoramento</title>
<style>
{p["css_bruto"]}{CSS_SELETOR_RECORTE}
</style></head><body>
<button type="button" class="btn-pdf" onclick="gerarPdf()" title="Imprime/salva em PDF o recorte selecionado, com o mapa, o gráfico, a avaliação e o detalhamento por status de cada município já abertos">Gerar PDF</button>
<header><div><h1>{p["titulo"]}</h1><p>COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS ESTRATÉGICAS</p></div>
<div class="gov">GOVERNO DO ESTADO DA BAHIA<br>CASA CIVIL</div></header>
<main>
{html_recortes}
</main>
<footer>CGAPE · Casa Civil · Base Ações do Cronograma</footer>
{seletor}
<div id="dica"></div>
{guarda_mapa}
{fundo_oc}
<script>
{p["js_bruto"]}{JS_SELETOR_RECORTE}
</script>
</body></html>"""


def montar_partes_html_monitoramento(caminho_base, dados, filtros=None):
    """Pedaços do HTML "Ações em Monitoramento", já isolados (CSS sob
    .mon-ra, ids sufixados, tooltip em #mon-dica) prontos para entrar como a
    página MON_ACOES da Apresentação CGAPE (ver gerar_pptx_pac.py) — pedido
    do usuário em 15/11/2026. Mesma estrutura da página autônoma
    (gerar_html_resumo_acoes), inclusive o seletor EIXO x
    SECRETARIA/EXECUTOR (pedido do usuário em 16/11/2026: os seletores de
    Gestão/Secretaria da apresentação, que são do PAC, somem nesta página).
    Base e filtros do Monitoramento são independentes dos do PAC; a página
    MON_ACOES é renderizada uma única vez, fora das variantes de
    gestão/secretaria do PAC."""
    html_recortes, seletor, p = _montar_recortes(caminho_base, dados,
                                                 {} if filtros is None else filtros, sufixo="monra")
    guarda = f'<div class="ra-mapa-guarda">{p["mapa_svg"]}</div>' if p.get("mapa_svg") else ""
    secoes = {"MON_ACOES": f'<div class="mon-ra">{html_recortes}{guarda}</div>'}

    # O tooltip (#dica) da página autônoma vira, no embutido, #mon-dica —
    # fora de qualquer .mon-ra, pra não sumir se o slide que o contém não
    # for o ativo — e window.gerarPdf vira window.monRaGerarPdf, sem botão
    # próprio: a apresentação já tem o seu, que imprime a página inteira.
    css_renomeado = (p["css_bruto"] + CSS_SELETOR_RECORTE).replace("#dica", "#mon-dica")
    css_isolado = _escopar_css(css_renomeado, ".mon-ra",
                               nao_escopar=("#mon-dica", "#ra-oc-fundo", ".sel-recorte",
                                            ".sel-barra", ".sel-barra-eixo", ".sel-barra-sec",
                                            "body.mostra-rodape"))
    js_isolado = (
        p["js_bruto"]
        .replace('getElementById("dica")', 'getElementById("mon-dica")')
        .replace("window.gerarPdf=", "window.monRaGerarPdf=")
        + JS_SELETOR_RECORTE
    )
    fundo_oc = '<div id="ra-oc-fundo">' + p["oc_html"].partition('<div id="ra-oc-fundo">')[2]
    return {
        "css": css_isolado,
        "secoes": secoes,
        "js": js_isolado,
        # barras EIXO x SECRETARIA/EXECUTOR: fora de qualquer slide (fixas no
        # rodapé), visíveis só na página do Monitoramento — ver
        # atualizarSeletoresMonitoramento() em gerar_pptx_pac.py
        "dica_html": '<div id="mon-dica"></div>' + seletor,
        # Janela de cards ao clicar num objeto (ver .ra-obj-clic): fica fora
        # de qualquer .mon-ra/slide; os dados de cada recorte vêm do
        # "ra-oc-dados--monra" do recorte visível (data-oc).
        "oc_html": fundo_oc,
        "recorte": p["recorte_txt"],
    }


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


def api_atualizar_conder():
    """Botão "ATUALIZAR DA CONDER" (pedido do usuário em 17/11/2026): baixa
    a planilha do serviço público da CONDER (importar_cronograma.py, mesmo
    conteúdo do botão "Exportar para planilha" do site) e grava por cima da
    da pasta do programa — só troca se o conteúdo mudou."""
    global _ra_cache
    try:
        import importar_cronograma as ic
        caminho = pac.caminho_recurso(RA_PLANILHA) if pac else RA_PLANILHA
        r = ic.atualizar_se_mudou(caminho)
    except PermissionError:
        return {"ok": False, "mensagem": f"Não consegui gravar '{RA_PLANILHA}': ela está aberta "
                                         "em outro programa (Excel?). Feche e tente de novo."}
    except Exception as exc:
        return {"ok": False, "mensagem": f"Não consegui baixar do serviço da CONDER: {exc}"}
    _ra_cache = None
    # o MONITORAMENTO (relatorio_monitoramento.py) guarda a base lida uma vez
    rm = sys.modules.get("relatorio_monitoramento")
    if rm is not None and hasattr(rm, "_BASE"):
        with rm._TRAVA_BASE:
            rm._BASE = None
    qtd = f"{r['acoes']:,}".replace(",", ".")
    return {"ok": True, "mudou": r["mudou"],
            "mensagem": (f"Base atualizada pelo serviço da CONDER: {qtd} ações."
                         if r["mudou"] else
                         f"A base já estava igual à da CONDER ({qtd} ações).")}


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
        html = gerar_html_resumo_acoes(caminho_base, dados, filtros or {})
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
        <button type="button" class="ra-inicio" id="raInicio" title="Voltar à tela inicial">← Início</button>
        <button type="button" id="raConder" title="Baixa agora a planilha Ações do Cronograma do serviço da CONDER (o mesmo conteúdo do botão Exportar para planilha do site app.conder.ba.gov.br/cronogramaacoes)">ATUALIZAR DA CONDER</button>
        <button type="button" id="raHtml" title="Gera uma página HTML com o gráfico de investimento dos 30 maiores municípios por população, a avaliação da carteira e o quadro Status × Eixo do recorte atual">GERAR HTML</button>
        <button type="button" id="raIncluirApresentacao" title="Abre a Apresentação CGAPE (HTML) já com as páginas do Monitoramento marcadas — você escolhe quais páginas do PAC entram junto">INCLUIR NA APRESENTAÇÃO</button>
        <button type="button" id="raGerar" title="Lê de novo a planilha Ações do Cronograma e gera o slide em Downloads">GERAR APRESENTAÇÃO</button>
        <button type="button" class="ra-fechar" id="raFechar">FECHAR</button>
      </div>
    </div>
    <div class="ra-corpo">
      <div class="ra-busca-geral">
        <input type="search" id="raBuscaGeral" autocomplete="off"
          placeholder="Buscar em todas as colunas da base (código, ação, município, executor, objeto, observações…)"
          title="Mostra só as ações que têm TODAS as palavras digitadas em qualquer coluna da planilha Ações do Cronograma">
        <button type="button" id="raBuscaGeralLimpar" title="Limpar a busca">×</button>
      </div>
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
  /* Texto branco — var(--cor-acento-teal) virou um azul escuro saturado na
     identidade única do projeto; #10251F (quase preto) não tinha mais
     contraste nenhum em cima dele (fazia sentido só com o teal claro
     antigo) — corrigido a pedido do usuário em 14/11/2026. */
  #raGerar, #raXlsx { background:var(--cor-acento-teal); color:#fff;
    border-color:var(--cor-acento-teal); font-weight:700; }
  #raGerar:hover, #raXlsx:hover { background:var(--cor-acento-teal-hover);
    color:#fff; }
  #raGerar:disabled, #raXlsx:disabled, #raHtml:disabled, #raConder:disabled { opacity:.6; cursor:default; }
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
  /* Bolinha branca — #10251F (quase preto) ficava escondida em cima do
     trilho, que virou um azul escuro saturado na identidade única do
     projeto (antes era um teal claro) — corrigido a pedido do usuário em
     14/11/2026. */
  .ra-top30.ativo .ra-top30-trilho span { transform:translateX(14px); background:#fff; }
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
    color:#fff; }
  .ra-data-campos { display:flex; flex-direction:column; gap:4px; }
  .ra-data-campos label { display:flex; align-items:center; gap:6px; font-size:11px;
    color:var(--cor-texto-secundario); }
  .ra-data-campos input { flex:1; min-width:0; background:var(--cor-fundo);
    color:var(--cor-texto-primario); border:1px solid var(--cor-card-elevado);
    border-radius:6px; padding:3px 6px; font-size:12px; font-family:inherit; }
  html[data-tema="escuro"] .ra-data-campos input { color-scheme:dark; }
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
  /* Buscas digitadas (pedido do usuário em 17/11/2026) */
  .ra-busca-geral { position:relative; margin-bottom:10px; }
  #raBuscaGeral { width:100%; box-sizing:border-box; background:var(--cor-card);
    color:var(--cor-texto-primario); border:1px solid var(--cor-card-elevado);
    border-radius:8px; padding:10px 36px 10px 38px; font-size:13px; font-family:inherit;
    box-shadow:var(--sombra-card); outline:none;
    background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%23889' stroke-width='2.4' stroke-linecap='round'%3E%3Ccircle cx='11' cy='11' r='7'/%3E%3Cpath d='M20 20l-4-4'/%3E%3C/svg%3E");
    background-repeat:no-repeat; background-position:12px center; }
  #raBuscaGeral::-webkit-search-cancel-button { display:none; }
  #raBuscaGeral:focus, .ra-busca-geral.ativo #raBuscaGeral { border-color:var(--cor-acento-teal); }
  #raBuscaGeralLimpar { position:absolute; right:8px; top:50%; transform:translateY(-50%);
    background:none; border:none; color:var(--cor-texto-secundario); font-size:18px;
    line-height:1; cursor:pointer; padding:2px 6px; display:none; }
  .ra-busca-geral.ativo #raBuscaGeralLimpar { display:block; }
  .ra-acoes th .ra-col-busca { display:block; width:100%; min-width:70px; box-sizing:border-box;
    margin-top:5px; background:var(--cor-fundo); color:var(--cor-texto-primario);
    border:1px solid var(--cor-texto-terciario); border-radius:5px; padding:3px 6px;
    font-size:11.5px; font-weight:400; font-family:inherit; cursor:text; outline:none; }
  .ra-acoes th .ra-col-busca:focus, .ra-acoes th .ra-col-busca.ativo {
    border-color:var(--cor-acento-teal); }
  .ra-acoes td.ra-sem-linhas { text-align:center; padding:18px; white-space:normal;
    color:var(--cor-texto-secundario); }
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

  // Chave de localStorage pros filtros deste overlay sobreviverem a uma
  // troca de tela (HOME/PAC/MONITORAMENTO recarrega a página inteira, ver
  // _carregar_tela em "CGAPE - BALANÇO PAC.py") — pedido do usuário em
  // 15/11/2026: o botão "VER APRESENTAÇÃO (HTML)" do painel do PAC precisa
  // enxergar o último filtro escolhido aqui, mesmo se a pessoa tiver saído
  // desta tela antes de clicar nele.
  var MON_RA_CHAVE_STORAGE = "monRaFiltros";
  var raFiltros = {};
  try {
    var raFiltrosSalvos = window.localStorage && localStorage.getItem(MON_RA_CHAVE_STORAGE);
    if (raFiltrosSalvos) raFiltros = JSON.parse(raFiltrosSalvos) || {};
  } catch (e) { /* localStorage indisponível (ex.: aba anônima) — segue com base completa */ }
  function raSalvarFiltros() {
    try { if (window.localStorage) localStorage.setItem(MON_RA_CHAVE_STORAGE, JSON.stringify(raFiltros)); }
    catch (e) { /* ignora — não é crítico, só perde a persistência entre telas */ }
  }
  // Exposto pro script do painel (montar_html_painel), que injeta este
  // overlay na mesma página — usado pelo botão "VER APRESENTAÇÃO (HTML)" /
  // "INCLUIR NA APRESENTAÇÃO" pra levar o filtro do Monitoramento junto do
  // filtro do PAC, sem cruzar as duas bases — pedido do usuário em
  // 15/11/2026.
  window.monRaFiltrosAtuais = function () { return raFiltros; };
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
  // 13/11/2026): um clique preenche de/até; clicar no ativo limpa. "Macro
  // Jerônimo" (os dois mandatos somados) adicionada a pedido do usuário em
  // 15/11/2026, mantendo M01/M02 com os mesmos períodos de cada mandato.
  var RA_MANDATOS = [
    {rotulo: "1º M", de: "2023-01-01", ate: "2026-12-31"},
    {rotulo: "2º M", de: "2027-01-01", ate: "2030-12-31"},
    {rotulo: "Jerônimo", de: "2023-01-01", ate: "2030-12-31"}];
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
    raSincronizarBuscaGeral();
  }

  // ---- tabela das ações do recorte ----------------------------------------
  var raBase = null, raBaseVersao = null, raBasePedido = null;
  var raIds = [];
  var raTab = {col: null, dir: 1, limite: 200};
  var RA_PASSO = 200;
  var raBRL = new Intl.NumberFormat("pt-BR", {minimumFractionDigits: 2, maximumFractionDigits: 2});
  // Buscas digitadas (pedido do usuário em 17/11/2026): geral (todas as
  // colunas da base, raFiltros.busca) e por coluna da tabela (raFiltros.bcol
  // = {rótulo: texto}). O filtro em si é feito no Python (resumo_acoes_dados),
  // então cards, quadros, tabela, Excel e HTML seguem o mesmo recorte.
  var raBuscaTimer = null;
  function raConsultarDepois() {
    clearTimeout(raBuscaTimer);
    raBuscaTimer = setTimeout(raConsultar, 350);
  }
  function raTemBuscaColuna() {
    return !!(raFiltros.bcol && Object.keys(raFiltros.bcol).length);
  }
  function raDefinirBuscaColuna(rot, texto) {
    var b = raFiltros.bcol || {};
    if (texto.trim()) b[rot] = texto; else delete b[rot];
    if (Object.keys(b).length) raFiltros.bcol = b; else delete raFiltros.bcol;
    raConsultarDepois();
  }
  (function () {
    var caixa = document.getElementById("raBuscaGeral");
    var limpar = document.getElementById("raBuscaGeralLimpar");
    if (!caixa) return;
    function aplicar(texto) {
      if (texto.trim()) raFiltros.busca = texto; else delete raFiltros.busca;
      caixa.parentNode.classList.toggle("ativo", !!texto.trim());
    }
    caixa.oninput = function () { aplicar(caixa.value); raConsultarDepois(); };
    caixa.onkeydown = function (e) {
      if (e.key === "Enter") { clearTimeout(raBuscaTimer); raConsultar(); }
      if (e.key === "Escape" && caixa.value) { e.stopPropagation(); caixa.value = ""; aplicar(""); raConsultar(); }
    };
    limpar.onclick = function () { caixa.value = ""; aplicar(""); raConsultar(); caixa.focus(); };
  })();
  function raSincronizarBuscaGeral() {
    var caixa = document.getElementById("raBuscaGeral");
    if (!caixa) return;
    if (document.activeElement !== caixa) caixa.value = raFiltros.busca || "";
    caixa.parentNode.classList.toggle("ativo", !!(raFiltros.busca || "").trim());
  }

  function raTabelaAtualizar(d) {
    var alvo = document.getElementById("raTabela");
    raIds = d.vazio ? [] : d.ids;
    raTab.limite = RA_PASSO;
    // com busca por coluna, a tabela fica mesmo vazia (senão some o campo
    // onde se digitou)
    if (!raIds.length && !raTemBuscaColuna()) { alvo.innerHTML = ""; return; }
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
    if (!raBase || (!raIds.length && !raTemBuscaColuna())) { alvo.innerHTML = ""; return; }
    // campo de busca da coluna em uso: volta com o foco e o cursor no lugar
    var focado = document.activeElement, focoRot = null, focoIni = 0, focoFim = 0;
    if (focado && focado.classList && focado.classList.contains("ra-col-busca") && alvo.contains(focado)) {
      focoRot = focado.dataset.rot; focoIni = focado.selectionStart; focoFim = focado.selectionEnd;
    }
    var bcol = raFiltros.bcol || {};
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
        ((c.filtro && raFiltros[c.filtro]) || (c.filtro === "municipio" && raFiltros.top30) ||
         bcol[c.rotulo] ? " ra-col-filtro" : "");
      var termo = bcol[c.rotulo] || "";
      return '<th class="' + classe + '" data-col="' + i + '" title="Ordenar">' + esc(c.rotulo) +
        (raTab.col === i ? (raTab.dir > 0 ? " ▲" : " ▼") : "") +
        '<input type="text" class="ra-col-busca' + (termo ? " ativo" : "") + '" data-rot="' + esc(c.rotulo) +
        '" value="' + esc(termo) + '" placeholder="Buscar…" autocomplete="off" title="Buscar só nesta coluna">' +
        "</th>";
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
      "</tr></thead><tbody>" + (corpo || '<tr><td class="ra-sem-linhas" colspan="' + (cols.length + 1) +
        '">Nenhuma ação com a busca e os filtros escolhidos.</td></tr>') + "</tbody></table></div>" +
      (resto > 0 ? '<button type="button" class="ra-mais">Mostrar mais ' + Math.min(RA_PASSO, resto) +
        " (faltam " + resto.toLocaleString("pt-BR") + ")</button>" : "");
    alvo.querySelectorAll(".ra-col-busca").forEach(function (inp) {
      inp.onclick = function (e) { e.stopPropagation(); };
      inp.onkeydown = function (e) {
        e.stopPropagation();
        if (e.key === "Enter") { clearTimeout(raBuscaTimer); raConsultar(); }
        if (e.key === "Escape" && inp.value) { inp.value = ""; raDefinirBuscaColuna(inp.dataset.rot, ""); }
      };
      inp.oninput = function () {
        inp.classList.toggle("ativo", !!inp.value.trim());
        raDefinirBuscaColuna(inp.dataset.rot, inp.value);
      };
      if (focoRot !== null && inp.dataset.rot === focoRot) {
        inp.focus();
        try { inp.setSelectionRange(focoIni, focoFim); } catch (e) { /* ignora */ }
      }
    });
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
    // Chamada depois de QUALQUER mutação de raFiltros (ver raAlternar,
    // raLimparFiltros etc.) — ponto único pra persistir em localStorage,
    // sem precisar repetir isso em cada lugar que mexe no filtro.
    raSalvarFiltros();
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
  // "ATUALIZAR DA CONDER" (pedido do usuário em 17/11/2026): baixa a base
  // do serviço da CONDER e relê tudo (a tabela recarrega pela versão nova)
  document.getElementById("raConder").onclick = function () {
    var msg = document.getElementById("raMsg");
    var botao = document.getElementById("raConder");
    botao.disabled = true;
    msg.textContent = "Baixando a planilha do serviço da CONDER (leva uns 30 segundos)…";
    api("resumo_acoes_atualizar_conder").then(function (r) {
      if (!r.ok) { msg.textContent = r.mensagem; return; }
      raBase = null; raBaseVersao = null;
      raConsultar();
      setTimeout(function () { msg.textContent = r.mensagem; }, 1500);
    }).catch(function (e) {
      msg.textContent = "Erro ao atualizar pela CONDER: " + e;
    }).finally(function () { botao.disabled = false; });
  };
  // "INCLUIR NA APRESENTAÇÃO" — abre a mesma janela de seleção de slides
  // da Apresentação CGAPE (definida no script do painel, injetado nesta
  // mesma página — ver window.abrirApresentacaoComMonitoramento em "CGAPE
  // - BALANÇO PAC.py"), já com as páginas MON_* marcadas. Os filtros do
  // Monitoramento usados são os deste overlay (raFiltros, via
  // window.monRaFiltrosAtuais); os do PAC são os últimos do painel (ou
  // base completa, se o painel nunca foi aberto) — pedido do usuário em
  // 15/11/2026.
  var botaoIncluirApresentacao = document.getElementById("raIncluirApresentacao");
  if (botaoIncluirApresentacao) {
    if (window.abrirApresentacaoComMonitoramento) {
      botaoIncluirApresentacao.onclick = window.abrirApresentacaoComMonitoramento;
    } else {
      // Nunca deveria acontecer (este overlay só existe injetado no painel
      // do PAC, que sempre define essa função) — esconde em vez de deixar
      // um botão quebrado.
      botaoIncluirApresentacao.style.display = "none";
    }
  }
  // irParaHome vem do próprio documento do painel (ver montar_html_painel)
  // — este overlay vive dentro dele, nunca sozinho — pedido do usuário em
  // 14/11/2026.
  var btnInicio = document.getElementById("raInicio");
  if (btnInicio && window.irParaHome) btnInicio.onclick = window.irParaHome;
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
