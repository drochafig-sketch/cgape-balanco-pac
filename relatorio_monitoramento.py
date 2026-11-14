# -*- coding: utf-8 -*-
"""Seção MONITORAMENTO do relatório PDF do CGAPE - BALANÇO PAC.

Base: "Ações do Cronograma.xlsx" (aba "Dados"), lida da mesma pasta do
.exe/script, igual à planilha do PAC. O resumo segue o modelo do Relatório
SSP (Painel Geral, Resumo Financeiro, Detalhamento por status em dois
capítulos e controle de qualidade de prazos), agrupando por EIXO onde o SSP
agrupa por OBJETO.

Nenhuma linha é descartada na leitura: qualquer recorte vem só dos filtros
do painel (seletor MONITORAMENTO).

Mesmo padrão do gerar_pptx_pac.py: o script principal chama
configurar(sys.modules[__name__]) e este módulo usa dele página, margens,
fontes, cores, geojson e formatação de valores, sem duplicar nada disso.
"""

import math
import os
import re
import threading
import unicodedata
from datetime import datetime

import pandas as pd
from reportlab.graphics.shapes import (Circle, Drawing, Line, Path, Rect,
                                       String, Wedge)
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import (Flowable, PageBreak, Paragraph, Spacer, Table,
                                TableStyle)

import identidade_visual as iv

pac = None  # módulo principal, injetado por configurar()


def configurar(modulo):
    global pac
    pac = modulo


ARQUIVO_PLANILHA = "Ações do Cronograma.xlsx"
ABA_DADOS = "Dados"
ROTULO_SECAO = "MONITORAMENTO"   # rótulo da numeração ("MONITORAMENTO - Página 01 de 40")

# As medidas dos desenhos foram portadas do Relatório SSP, cuja página útil
# tem 892 x 460 pt (960 x 540 menos as margens). Aqui tudo é desenhado nas
# medidas do SSP e ampliado pela razão entre as larguras úteis, para manter
# as mesmas proporções na página maior do BALANÇO PAC.
LARGURA_UTIL_SSP = 960 - 24 * 72 / 25.4


def _largura():
    """Largura do frame do PDF (margens e recuo interno do ReportLab)."""
    return pac.LARGURA_UTIL - 2 * pac.PADDING_FRAME_DOCUMENTO


def _altura():
    return pac.ALTURA_UTIL - 2 * pac.PADDING_FRAME_DOCUMENTO


def escala():
    return _largura() / LARGURA_UTIL_SSP


# Um grupo por código de status, sem agregação. Ordem do fim do processo
# para o começo. Cores alinhadas à identidade visual única do projeto (ver
# identidade_visual.CORES_CATEGORICAS — escala de 14 tons derivada da
# paleta da referência, não cores soltas) — antes eram uma paleta própria
# deste relatório, sem vermelho (reservado ao prazo vencido e à faixa alta
# do mapa, continua reservado) — corrigido a pedido do usuário em
# 14/11/2026.
_CC = iv.CORES_CATEGORICAS
GRUPOS_STATUS = [
    ("01", "OBRAS CONCLUÍDAS", f"#{_CC[0]}"),
    ("10", "OBRAS INAUGURADAS", f"#{_CC[1]}"),
    ("02", "OBRAS/SERVIÇOS EM ANDAMENTO", f"#{_CC[2]}"),
    ("09", "EMISSÃO ORDEM DE SERVIÇO", f"#{_CC[3]}"),
    ("03", "EM LICITAÇÃO/CONTRATAÇÃO", f"#{_CC[4]}"),
    ("04", "EM PROCESSAMENTO DE LICITAÇÃO", f"#{_CC[5]}"),
    ("12", "AUTORIZADO", f"#{_CC[6]}"),
    ("13", "AGUARDANDO AUTORIZO", f"#{_CC[7]}"),
    ("05", "EM DESENVOLVIMENTO DE PROJETOS", f"#{_CC[8]}"),
    ("07", "EM ESTUDO", f"#{_CC[9]}"),
    ("08", "VERIFICAÇÃO DE VIABILIDADE", f"#{_CC[10]}"),
    ("14", "OBRAS CANCELADAS", f"#{_CC[12]}"),
]
NOME_POR_CODIGO = {c: n for c, n, _h in GRUPOS_STATUS}
COR_POR_GRUPO = {n: colors.HexColor(h) for _c, n, h in GRUPOS_STATUS}
ORDEM_GRUPOS = [n for _c, n, _h in GRUPOS_STATUS]
CODIGOS_LICITACAO = {"03", "04"}   # layout de seção em licitação
GRUPO_ANDAMENTO = "OBRAS/SERVIÇOS EM ANDAMENTO"
GRUPOS_CONCLUIDOS = {"OBRAS CONCLUÍDAS", "OBRAS INAUGURADAS"}

SEM_INFO = "NÃO INFORMADO"
SEM_PRAZO = "SEM PRAZO DEFINIDO"
MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Filtros do painel: (chave, coluna derivada, título do bloco)
CAMPOS_FILTRO = [
    ("EIXO", "EIXO", "Eixo"),
    ("SITUACAO", "GRUPO_STATUS", "Situação"),
    ("RESPONSAVEL", "RESPONSAVEL", "Responsável Informação"),
    ("EXECUTOR", "EXECUTOR", "Executor"),
    ("PROGRAMA", "PROGRAMA", "Programa de Transferência de Recursos"),
    ("OBJETO", "OBJETO", "Objeto"),
    ("ESCOPO", "ESCOPO", "Escopo"),
    ("TERRITORIO", "TERRITORIO", "Território de Identidade"),
    ("MUNICIPIO", "MUNICIPIO", "Município"),
]

# Colunas disponíveis no detalhamento: (rótulo, peso de largura, alinhamento)
COLUNAS_DETALHAMENTO = [
    ("CÓDIGO", 5, "centro"),
    ("CONTRATO", 7, "centro"),
    ("AÇÃO", 24, "esq"),
    ("ESCOPO", 8.5, "centro"),
    ("DETALHAMENTO ESCOPO", 12, "esq"),
    ("OBJETO", 11, "centro"),
    ("EIXO", 8, "centro"),
    ("TERRITÓRIO", 9, "centro"),
    ("MUNICÍPIO", 8.5, "centro"),
    ("ORÇAMENTO", 7.5, "dir"),
    ("INVESTIMENTO (R$)", 7.5, "dir"),
    ("STATUS", 11, "centro"),
    ("AVANÇO", 4.5, "centro"),
    ("CONTRATADA", 14, "centro"),
    ("ENTREGA OBRA", 6, "centro"),
]
COLUNAS_PADRAO = ["CONTRATO", "AÇÃO", "ESCOPO", "OBJETO", "MUNICÍPIO",
                  "INVESTIMENTO (R$)", "CONTRATADA", "AVANÇO", "ENTREGA OBRA"]
LIMITE_COLUNAS = 12
# Em licitação a obra ainda não começou: saem avanço, contrato e contratada
FORA_NA_LICITACAO = {"AVANÇO", "CONTRATO", "CONTRATADA"}

OPCOES_ORDEM = [
    ("MUNICIPIO", "Município (A a Z)"),
    ("ACAO", "Ação (A a Z)"),
    ("VALOR_DESC", "Investimento (maior para menor)"),
    ("AVANCO_DESC", "Avanço (maior para menor)"),
    ("PRAZO_ASC", "Entrega da obra (mais próxima primeiro)"),
    ("CONTRATO", "Contrato (ano, depois número)"),
]

ACOES_POR_PAGINA_MAX = 50      # teto por página do detalhamento
# Corpo das células do detalhamento, em medida do SSP (ampliado pela escala:
# 6,4 dá 10,2 pt nesta página). O SSP usa 7,4; o usuário escolheu 6,4 em
# 11/11/2026 para chegar perto de 30 ações por página (média 34).
FONTE_DETALHAMENTO_SSP = 6.4

# Páginas da seção na janela de seleção do relatório
SECOES = [
    {"chave": "MON_PAINEL_GERAL", "titulo": "Monitoramento: Painel Geral",
     "layout": "graficos"},
    {"chave": "MON_RESUMO_FINANCEIRO",
     "titulo": "Monitoramento: Resumo Financeiro", "layout": "barras"},
    {"chave": "MON_DETALHAMENTO", "titulo": "Monitoramento: Detalhamento",
     "layout": "detalhamento"},
]
CHAVES_SECOES = {s["chave"] for s in SECOES}

# (limite inferior, limite superior, cor), maior para menor. Mesmas 3 faixas
# e cores do SSP, com limites para a escala desta base (R$ 3 Mi / R$ 10 Mi
# do SSP deixavam quase todo o estado na faixa de cima): R$ 30 Mi e R$ 60 Mi
# dividem os municípios da base completa em terços.
CLASSES_MAPA = [
    (60_000_000, None, colors.HexColor("#D32F45")),
    (30_000_000, 60_000_000, colors.HexColor("#F0834E")),
    (0, 30_000_000, colors.HexColor("#2E75B6")),
]
COR_SEM_ACAO = colors.HexColor("#DCDCDC")
COR_VENCIDO = colors.HexColor("#C0182B")
COR_TEXTO = colors.HexColor("#1A1A1A")
COR_CINZA_LINHA = colors.HexColor("#E5E8EF")
COR_LINHA_ALTERNADA = colors.HexColor("#F4F6FA")


# ----------------------------------------------------------------------------
# UTILITÁRIOS
# ----------------------------------------------------------------------------

def normalizar(texto):
    """Sem acento e em caixa baixa, para agrupar e casar com o geojson."""
    if texto is None or (isinstance(texto, float) and math.isnan(texto)):
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in s if not unicodedata.combining(c)).strip().lower()


def mi_bi(valor):
    """formatar_mi_bi do PAC sem as entidades HTML (texto de desenho)."""
    return pac.formatar_mi_bi(float(valor or 0)).replace("&nbsp;", " ")


def moeda(valor):
    v = float(valor or 0)
    if v <= 0:
        return "-"
    txt = f"R$ {v:,.2f}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


def _numero_moeda(v):
    """'R$\xa011.206.520,78' -> 11206520.78"""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return float("nan")
    s = (str(v).replace("\xa0", " ").replace("R$", "").strip()
         .replace(".", "").replace(",", "."))
    try:
        return float(s)
    except ValueError:
        return float("nan")


def _escapar(texto):
    return (str(texto).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def _fonte():
    return pac.FONTE_PADRAO


def _fonte_negrito():
    return pac.FONTE_PADRAO_NEGRITO


def _largura_texto(texto, fonte, tamanho):
    return pdfmetrics.stringWidth(texto, fonte, tamanho)


def _desenho_ampliado(largura_ssp, altura_ssp):
    """Drawing desenhado em medidas do SSP e ampliado pela escala."""
    s = escala()
    d = Drawing(largura_ssp * s, altura_ssp * s)
    d.scale(s, s)
    return d


# ----------------------------------------------------------------------------
# LEITURA DA BASE (sem descartar nenhuma linha)
# ----------------------------------------------------------------------------

_BASE = None   # DataFrame preparado; lido uma vez, na primeira utilização
_DATA_BASE = None
_TRAVA_BASE = threading.Lock()


def caminho_planilha():
    return pac.caminho_recurso(ARQUIVO_PLANILHA)


def base_disponivel():
    return os.path.exists(caminho_planilha())


def _texto(serie):
    return serie.fillna("").astype(str).str.strip()


def _com_padrao(serie):
    s = _texto(serie)
    return s.where(s != "", SEM_INFO)


def _grafia_unificada(serie):
    """Unifica grafias que só diferem em caixa/acento pela mais frequente
    ('construção' e 'Construção' viram 'Construção')."""
    s = _com_padrao(serie)
    chave = s.map(normalizar)
    preferida = (pd.DataFrame({"k": chave, "v": s})
                 .groupby("k")["v"].agg(lambda x: x.value_counts().index[0]))
    return chave.map(preferida)


# Grafias da planilha que não casam com o nome oficial do geojson (chaves
# já normalizadas). Só as inequívocas; o que não estiver aqui e não casar
# fica fora do mapa, mas continua em todas as contas.
APELIDOS_MUNICIPIO = {
    "ibirapoa": "ibirapua",
    "lafayette coutinho": "lafaiete coutinho",
    "ltiuba": "itiuba",
    "muquem do sao francisco": "muquem de sao francisco",
    "pau-brasil": "pau brasil",
    "salinas das margaridas": "salinas da margarida",
    "santa terezinha": "santa teresinha",
}


def _nomes_municipios(serie):
    """Nome oficial do geojson quando casa; senão a grafia mais frequente."""
    s = _com_padrao(serie)
    chave = s.map(normalizar).map(lambda k: APELIDOS_MUNICIPIO.get(k, k))
    oficiais = {}
    geo = pac.carregar_geojson_municipios()
    if geo:
        for f in geo["features"]:
            nome = str(f["properties"].get("name", ""))
            oficiais[normalizar(nome)] = nome.upper()
    frequente = (pd.DataFrame({"k": chave, "v": s.str.upper()})
                 .groupby("k")["v"].agg(lambda x: x.value_counts().index[0]))
    return chave.map(lambda k: oficiais.get(k) or frequente[k]), chave


def ler_base(caminho=None):
    caminho = caminho or caminho_planilha()
    df = pd.read_excel(caminho, sheet_name=ABA_DADOS, dtype=str)

    out = pd.DataFrame(index=df.index)
    out["CODIGO"] = _texto(df["Código"])
    out["ACAO"] = _texto(df["Ação"])
    out["TERRITORIO"] = _com_padrao(df["Território Identidade"])
    out["MUNICIPIO"], out["MUN_NORM"] = _nomes_municipios(df["Município"])
    out["EXECUTOR"] = _com_padrao(df["Executor"])
    out["NUMERO_CONTRATO"] = _texto(df["Número do Contrato"])
    out["PROGRAMA"] = _com_padrao(df["Programa de Transferência de Recursos"])
    out["CONTRATADA"] = _texto(df["Contratada"])
    out["ESCOPO"] = _grafia_unificada(df["Escopo"])
    out["DETALHAMENTO_ESCOPO"] = _texto(df["Detalhamento Escopo"])
    out["OBJETO"] = _com_padrao(df["Objeto"])
    out["STATUS"] = _com_padrao(df["Status"])
    out["EIXO"] = _com_padrao(df["Eixo"])
    out["RESPONSAVEL"] = _com_padrao(df["Responsável Informação"])

    codigo = out["STATUS"].str[:2]
    out["CODIGO_STATUS"] = codigo
    out["GRUPO_STATUS"] = codigo.map(NOME_POR_CODIGO).fillna(
        out["STATUS"].str.upper())

    out["ORCAMENTO"] = df["Valor do Orçamento"].map(_numero_moeda)
    out["INVESTIMENTO"] = df["Valor de Investimento"].map(_numero_moeda)
    # Valor das contas: investimento e, na falta dele, orçamento (como no SSP)
    out["VALOR"] = out["INVESTIMENTO"].where(
        out["INVESTIMENTO"].fillna(0) > 0, out["ORCAMENTO"]).fillna(0.0)
    # Nas seções em licitação o valor exibido é o orçamento (com o
    # investimento quando o orçamento está vazio)
    out["VALOR_LICITACAO"] = out["ORCAMENTO"].where(
        out["ORCAMENTO"].fillna(0) > 0, out["INVESTIMENTO"]).fillna(0.0)

    out["AVANCO"] = pd.to_numeric(
        _texto(df["Avanço"]).str.replace(",", "."), errors="coerce")

    out["ENTREGA"] = pd.to_datetime(df["Entrega Obra"], format="%d/%m/%Y",
                                    errors="coerce")
    out["AVISO_LICITACAO"] = pd.to_datetime(
        df["Aviso de Licitação"], format="%d/%m/%Y", errors="coerce")
    out["PRAZO_ANO"] = out["ENTREGA"].dt.year.map(
        lambda a: f"{int(a)}" if pd.notna(a) else SEM_PRAZO)
    out["PRAZO_MES"] = out["ENTREGA"].dt.strftime("%Y-%m").fillna("")

    # Atraso já calculado na base ("42 dias"); vazio = sem atraso
    out["DIAS_ATRASO"] = pd.to_numeric(
        _texto(df["Dias em Atraso (De Marco)"]).str.extract(r"(-?\d+)")[0],
        errors="coerce").fillna(0).astype(int)
    return out.reset_index(drop=True)


def base():
    """Base preparada (lida na primeira chamada). None se não houver a
    planilha nesta máquina."""
    global _BASE, _DATA_BASE
    with _TRAVA_BASE:   # a pré-carga em segundo plano e o painel podem chegar juntos
        if _BASE is None:
            if not base_disponivel():
                return None
            _BASE = ler_base()
            _DATA_BASE = datetime.fromtimestamp(
                os.path.getmtime(caminho_planilha()))
    return _BASE


def texto_atualizacao():
    data = _DATA_BASE or datetime.now()
    return data.strftime("Fonte: Ações do Cronograma - Atualizado em: "
                         "%d/%m/%Y às %Hhs%Mmin")


# ----------------------------------------------------------------------------
# FILTROS DO PAINEL
# ----------------------------------------------------------------------------

def _ordem_opcoes(chave, valores):
    if chave == "SITUACAO":
        return ([g for g in ORDEM_GRUPOS if g in valores]
                + sorted(v for v in valores if v not in ORDEM_GRUPOS))
    return sorted(valores, key=lambda v: (v == SEM_INFO, normalizar(v)))


def _arvore_conclusao(df):
    """[{chave, rotulo, meses:[{chave, rotulo}]}]: anos crescentes e
    'Sem prazo definido' ao final (menu CONCLUSÃO, sobre Entrega Obra)."""
    itens = []
    anos = sorted(a for a in df["PRAZO_ANO"].unique() if a.isdigit())
    for ano in anos:
        meses = sorted(m for m in df.loc[df["PRAZO_ANO"] == ano,
                                         "PRAZO_MES"].unique() if m)
        itens.append({"chave": ano, "rotulo": ano,
                      "meses": [{"chave": m,
                                 "rotulo": MESES_PT[int(m[5:]) - 1]}
                                for m in meses]})
    if (df["PRAZO_ANO"] == SEM_PRAZO).any():
        itens.append({"chave": SEM_PRAZO, "rotulo": "Sem prazo definido",
                      "meses": []})
    return itens


def dados_painel():
    """Tudo o que a grade de filtros do MONITORAMENTO precisa, de uma vez:
    blocos, linhas para a filtragem cruzada no navegador, menu CONCLUSÃO,
    ordenação e colunas do detalhamento."""
    df = base()
    if df is None:
        return None
    blocos = [{"chave": chave, "titulo": titulo,
               "opcoes": _ordem_opcoes(chave, df[coluna].unique().tolist())}
              for chave, coluna, titulo in CAMPOS_FILTRO]
    colunas_linha = [coluna for _c, coluna, _t in CAMPOS_FILTRO]
    linhas = df[colunas_linha + ["PRAZO_ANO", "PRAZO_MES"]].copy()
    linhas.columns = [c for c, _col, _t in CAMPOS_FILTRO] + [
        "PRAZO_ANO", "PRAZO_MES"]
    return {
        "blocos": blocos,
        "linhas": linhas.to_dict("records"),
        "conclusao": _arvore_conclusao(df),
        "ordens": [{"chave": c, "rotulo": r} for c, r in OPCOES_ORDEM],
        "colunas": [{"chave": r, "padrao": r in COLUNAS_PADRAO}
                    for r, _p, _a in COLUNAS_DETALHAMENTO],
        "limiteColunas": LIMITE_COLUNAS,
        "atualizacao": texto_atualizacao(),
        "total": int(len(df)),
    }


def filtrar(filtros):
    """Recorte do MONITORAMENTO pelos filtros do painel ({} = base toda)."""
    df = base()
    if df is None:
        return None
    filtros = filtros or {}
    for chave, coluna, _t in CAMPOS_FILTRO:
        valores = filtros.get(chave) or []
        if valores:
            df = df[df[coluna].isin(valores)]
    conclusao = filtros.get("CONCLUSAO") or []
    if conclusao:
        anos = {v for v in conclusao if "-" not in v}
        meses = {v for v in conclusao if "-" in v}
        df = df[df["PRAZO_ANO"].isin(anos) | df["PRAZO_MES"].isin(meses)]
    return df


# ----------------------------------------------------------------------------
# CONTROLE DE QUALIDADE (coluna "Dias em Atraso (De Marco)" da base)
# ----------------------------------------------------------------------------

def acoes_em_atraso(df):
    """Ações com atraso de marco informado na base, maior atraso primeiro.
    Nas obras em andamento coincide com o critério do SSP (Entrega Obra
    anterior a hoje)."""
    atrasadas = df[df["DIAS_ATRASO"] > 0]
    return atrasadas.sort_values("DIAS_ATRASO", ascending=False)


def aviso_qualidade(df):
    """{total, grupos:[{status, cor, itens:[{titulo, detalhe}]}]} ou None."""
    if df is None or df.empty:
        return None
    atrasadas = acoes_em_atraso(df)
    if atrasadas.empty:
        return None
    grupos = []
    for nome in ORDEM_GRUPOS + sorted(set(atrasadas["GRUPO_STATUS"])
                                      - set(ORDEM_GRUPOS)):
        parte = atrasadas[atrasadas["GRUPO_STATUS"] == nome]
        if parte.empty:
            continue
        itens = []
        for _i, r in parte.iterrows():
            contrato = r["NUMERO_CONTRATO"] or "s/ contrato"
            data = _data_da_linha(r, r["CODIGO_STATUS"] in CODIGOS_LICITACAO)
            avanco = _texto_avanco(r["AVANCO"])
            itens.append({
                "titulo": f"{r['CODIGO']} | {contrato} | {r['ACAO']} "
                          f"({r['MUNICIPIO']})",
                "detalhe": f"marco: {data} | {r['DIAS_ATRASO']} dia(s) em "
                           f"atraso | avanço: {avanco}",
            })
        grupos.append({"status": nome, "cor": COR_POR_GRUPO.get(
            nome, colors.grey).hexval().replace("0x", "#"), "itens": itens})
    return {"total": int(len(atrasadas)), "grupos": grupos}


# ----------------------------------------------------------------------------
# COMPONENTES GRÁFICOS (medidas do SSP, ampliadas)
# ----------------------------------------------------------------------------

def _estilos():
    s = escala()
    f, fb = _fonte(), _fonte_negrito()
    corpo = FONTE_DETALHAMENTO_SSP * s
    entrelinha = corpo * 1.12
    return {
        "titulo_barra": ParagraphStyle(
            "mon_titulo_barra", fontName=fb, fontSize=12 * s,
            leading=14 * s, textColor=pac.COR_MARCA_TEAL_ESCURO,
            alignment=TA_LEFT),
        "nota": ParagraphStyle(
            "mon_nota", fontName=f, fontSize=7.5 * s, leading=9.5 * s,
            textColor=colors.HexColor("#555555")),
        "esq": ParagraphStyle(
            "mon_det_esq", fontName=f, fontSize=corpo, leading=entrelinha,
            textColor=COR_TEXTO),
        "dir": ParagraphStyle(
            "mon_det_dir", fontName=f, fontSize=corpo, leading=entrelinha,
            textColor=COR_TEXTO, alignment=TA_RIGHT),
        "centro": ParagraphStyle(
            "mon_det_centro", fontName=f, fontSize=corpo,
            leading=entrelinha, textColor=COR_TEXTO,
            alignment=TA_CENTER),
        "cab": ParagraphStyle(
            "mon_det_cab", fontName=fb, fontSize=corpo + 0.2 * s,
            leading=corpo + 2 * s, textColor=colors.white,
            alignment=TA_CENTER),
    }


def titulo_secao(texto, estilos):
    s = escala()
    tab = Table([[Paragraph(texto, estilos["titulo_barra"])]],
                colWidths=[_largura()])
    tab.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pac.COR_MARCA_TEAL_CLARA),
        ("LEFTPADDING", (0, 0), (-1, -1), 8 * s),
        ("TOPPADDING", (0, 0), (-1, -1), 4 * s),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4 * s),
        ("LINEBELOW", (0, 0), (-1, -1), 1.4 * s, pac.COR_MARCA_GOLD),
    ]))
    return tab


def card_kpi(titulo, valor, largura, altura):
    """largura/altura em medidas do SSP."""
    d = _desenho_ampliado(largura, altura)
    fb = _fonte_negrito()
    f_titulo = max(7.5, min(9.5, altura * 0.115))
    f_valor = max(13.0, min(21.0, altura * 0.26))
    d.add(Rect(0, 0, largura, altura, rx=8, ry=8,
               fillColor=pac.COR_MARCA_TEAL_ESCURO, strokeColor=None))
    d.add(Rect(0, altura - 3, largura, 3, fillColor=pac.COR_MARCA_GOLD,
               strokeColor=None))
    d.add(String(largura / 2, altura - f_titulo - 5, titulo,
                 textAnchor="middle", fontName=fb, fontSize=f_titulo,
                 fillColor=pac.COR_MARCA_TEAL_CLARA))
    d.add(String(largura / 2, altura * 0.42 - f_valor / 2, valor,
                 textAnchor="middle", fontName=fb, fontSize=f_valor,
                 fillColor=colors.white))
    return d


def resumir_por(df, coluna):
    g = (df.groupby(coluna)
           .agg(qtd=("VALOR", "size"), valor=("VALOR", "sum"))
           .reset_index())
    return [(r[coluna], int(r["qtd"]), float(r["valor"]))
            for _i, r in g.sort_values("valor", ascending=False).iterrows()]


def donut_status(resumo, largura, altura):
    """resumo: [(grupo, qtd, valor)] na ordem dos grupos. Investimento em
    destaque na legenda e quantidade na linha de apoio (como no SSP)."""
    d = _desenho_ampliado(largura, altura)
    f, fb = _fonte(), _fonte_negrito()
    compacto = altura < 170
    raio = min(altura / 2 - 4, 72)
    cx, cy = raio + 6, altura / 2
    total = sum(v for _g, _q, v in resumo) or 1.0
    ang = 90.0
    for grupo, _qtd, valor in resumo:
        fatia = 360.0 * (valor / total)
        if fatia <= 0:
            continue
        d.add(Wedge(cx, cy, raio, ang - fatia, ang, yradius=raio,
                    radius1=raio * 0.56,
                    fillColor=COR_POR_GRUPO.get(grupo, colors.grey),
                    strokeColor=colors.white, strokeWidth=1.5))
        pct = 100.0 * valor / total
        if pct >= 4:
            meio = math.radians(ang - fatia / 2)
            d.add(String(cx + raio * 0.78 * math.cos(meio),
                         cy + raio * 0.78 * math.sin(meio) - 4,
                         f"{pct:.0f}%", textAnchor="middle", fontName=fb,
                         fontSize=10 if compacto else 12,
                         fillColor=colors.white))
        ang -= fatia

    f_centro = 13 if compacto else 15
    d.add(String(cx, cy + 1, mi_bi(total), textAnchor="middle",
                 fontName=fb, fontSize=f_centro,
                 fillColor=pac.COR_MARCA_TEAL_ESCURO))
    d.add(String(cx, cy - f_centro + 0.5, "investimento",
                 textAnchor="middle", fontName=f, fontSize=8.5,
                 fillColor=COR_TEXTO))

    # Legenda à direita, em duas colunas (são 12 status, contra 9 no SSP):
    # investimento em destaque na cor do status, quantidade ao lado e o
    # nome do status na linha de apoio.
    lx0 = cx + raio + 14
    colunas = 2 if len(resumo) > 6 else 1
    por_coluna = math.ceil(len(resumo) / colunas)
    larg_coluna = (largura - lx0) / colunas
    margem_v = 4
    passo = min(34.0, (altura - 2 * margem_v) / max(por_coluna, 1))
    f_valor = min(12.0, passo * 0.42)
    f_nome = min(8.0, passo * 0.3)
    f_qtd = f_nome
    quad = min(10.0, passo * 0.34)
    topo = cy + passo * por_coluna / 2
    for i, (grupo, qtd, valor) in enumerate(resumo):
        coluna, linha = divmod(i, por_coluna)
        lx = lx0 + coluna * larg_coluna
        ly = topo - linha * passo - f_valor
        cor = COR_POR_GRUPO.get(grupo, colors.grey)
        d.add(Rect(lx, ly - 0.5, quad, quad, fillColor=cor, strokeColor=None))
        rotulo_valor = mi_bi(valor)
        d.add(String(lx + quad + 4, ly, rotulo_valor, fontName=fb,
                     fontSize=f_valor, fillColor=cor))
        d.add(String(lx + quad + 8 + _largura_texto(rotulo_valor, fb, f_valor),
                     ly, f"({qtd})", fontName=f, fontSize=f_qtd,
                     fillColor=colors.HexColor("#555555")))
        nome = grupo
        while _largura_texto(nome, fb, f_nome) > larg_coluna - quad - 8:
            nome = nome[:-2] + "…"
        d.add(String(lx + quad + 4, ly - f_nome - 1.5, nome, fontName=fb,
                     fontSize=f_nome, fillColor=COR_TEXTO))
    return d


def barras_eixo(resumo, largura, altura):
    """resumo: [(eixo, qtd, valor)] ordenado por valor."""
    d = _desenho_ampliado(largura, altura)
    f, fb = _fonte(), _fonte_negrito()
    if not resumo:
        d.add(String(largura / 2, altura / 2, "sem dados",
                     textAnchor="middle", fontName=f, fontSize=10))
        return d
    max_v = max(v for _o, _q, v in resumo) or 1.0
    total = sum(v for _o, _q, v in resumo) or 1.0
    margem_esq = 150
    faixa = min(40.0, (altura - 8) / len(resumo))
    alt_barra = min(15.0, faixa * 0.62)
    larg_util = largura - margem_esq - 150
    y = altura - faixa / 2 - 4
    f_nome = min(9.5, max(7.5, faixa * 0.55))
    for eixo, qtd, valor in resumo:
        w = max(larg_util * (valor / max_v), 2)
        d.add(Rect(margem_esq, y - alt_barra / 2, w, alt_barra,
                   fillColor=pac.COR_MARCA_TEAL_ESCURO, strokeColor=None,
                   rx=2, ry=2))
        d.add(Rect(margem_esq, y + alt_barra / 2 - 2.5, w, 2.5,
                   fillColor=pac.COR_MARCA_GOLD, strokeColor=None))
        nome = eixo if len(eixo) <= 30 else eixo[:29] + "…"
        d.add(String(margem_esq - 8, y - 3.2, nome, textAnchor="end",
                     fontName=fb, fontSize=f_nome, fillColor=COR_TEXTO))
        rotulo = f"{mi_bi(valor)} ({qtd})"
        d.add(String(margem_esq + w + 8, y - 3, rotulo, fontName=fb,
                     fontSize=f_nome - 0.5,
                     fillColor=pac.COR_MARCA_TEAL_ESCURO))
        d.add(String(margem_esq + w + 14
                     + _largura_texto(rotulo, fb, f_nome - 0.5), y - 3,
                     f"{100.0 * valor / total:.1f}%".replace(".", ","),
                     fontName=f, fontSize=f_nome - 1.5,
                     fillColor=colors.HexColor("#555555")))
        y -= faixa
    d.add(Line(margem_esq, y + faixa - alt_barra / 2 - 4, margem_esq,
               altura - 4, strokeColor=COR_CINZA_LINHA, strokeWidth=1))
    return d


def barras_empilhadas_eixo(df, largura, altura):
    """Uma barra por EIXO, segmentada pelas situações (EIXO x SITUAÇÃO)."""
    d = _desenho_ampliado(largura, altura)
    f, fb = _fonte(), _fonte_negrito()
    eixos = [o for o, _q, v in resumir_por(df, "EIXO") if v > 0]
    grupos = [g for g in ORDEM_GRUPOS if g in set(df["GRUPO_STATUS"])]
    if not eixos:
        d.add(String(largura / 2, altura / 2, "sem dados",
                     textAnchor="middle", fontName=f, fontSize=10))
        return d
    total_geral = float(df["VALOR"].sum()) or 1.0
    maior = max(float(df.loc[df["EIXO"] == o, "VALOR"].sum())
                for o in eixos) or 1.0

    margem_esq = 150
    margem_dir = 150
    alt_legenda = 44
    larg_util = largura - margem_esq - margem_dir
    faixa = (altura - alt_legenda) / len(eixos)
    alt_barra = min(46.0, faixa * 0.58)
    y = altura - faixa / 2 - alt_barra / 2
    f_nome = min(11.0, max(8.0, faixa * 0.4))
    for eixo in eixos:
        sub = df[df["EIXO"] == eixo]
        total_eixo = float(sub["VALOR"].sum())
        largura_barra = larg_util * (total_eixo / maior)
        d.add(String(margem_esq - 8, y + alt_barra / 2 - 3.5,
                     eixo if len(eixo) <= 30 else eixo[:29] + "…",
                     textAnchor="end", fontName=fb, fontSize=f_nome,
                     fillColor=COR_TEXTO))
        x = margem_esq
        for grupo in grupos:
            valor = float(sub.loc[sub["GRUPO_STATUS"] == grupo, "VALOR"].sum())
            if valor <= 0:
                continue
            w = largura_barra * (valor / total_eixo)
            d.add(Rect(x, y, w, alt_barra, fillColor=COR_POR_GRUPO[grupo],
                       strokeColor=colors.white, strokeWidth=0.8))
            pct = 100.0 * valor / total_eixo
            if w > 34 and pct >= 7:
                d.add(String(x + w / 2, y + alt_barra / 2 - 3.5,
                             f"{pct:.0f}%", textAnchor="middle", fontName=fb,
                             fontSize=9, fillColor=colors.white))
            x += w
        d.add(String(margem_esq + largura_barra + 10, y + alt_barra / 2 + 1.5,
                     mi_bi(total_eixo), fontName=fb, fontSize=11.5,
                     fillColor=pac.COR_MARCA_TEAL_ESCURO))
        d.add(String(margem_esq + largura_barra + 10, y + alt_barra / 2 - 9,
                     f"{len(sub)} ação(ões) · "
                     f"{100.0 * total_eixo / total_geral:.0f}% do total",
                     fontName=f, fontSize=8, fillColor=colors.HexColor("#555555")))
        y -= faixa

    # legenda das situações, embaixo
    lx, ly = margem_esq, 26
    for grupo in grupos:
        larg_item = _largura_texto(grupo, f, 8) + 26
        if lx + larg_item > largura - 4:
            lx, ly = margem_esq, ly - 14
        d.add(Rect(lx, ly, 9, 9, fillColor=COR_POR_GRUPO[grupo],
                   strokeColor=None))
        d.add(String(lx + 13, ly + 1, grupo, fontName=f, fontSize=8,
                     fillColor=COR_TEXTO))
        lx += larg_item
    return d


# ----------------------------------------------------------------------------
# MAPA COROPLÉTICO (3 faixas, legenda dinâmica) E PINS
# ----------------------------------------------------------------------------

def _poligonos(feat):
    geom = feat["geometry"]
    if geom["type"] == "Polygon":
        return [geom["coordinates"]]
    if geom["type"] == "MultiPolygon":
        return geom["coordinates"]
    # geojson do PAC: lista de anéis direto em "coordinates"
    return [geom["coordinates"]]


def legenda_classes_mapa(vmax):
    itens = []
    for minimo, maximo, cor in CLASSES_MAPA:
        if vmax < minimo:
            continue
        if minimo == 0:
            teto = vmax if vmax < maximo else maximo
            rotulo = f"até {mi_bi(teto)}"
        else:
            teto = vmax if (maximo is None or vmax < maximo) else maximo
            rotulo = f"{mi_bi(minimo)} a {mi_bi(teto)}"
        itens.append((cor, rotulo))
    return itens


def cor_da_classe(valor):
    for minimo, maximo, cor in CLASSES_MAPA:
        if valor >= minimo and (maximo is None or valor < maximo):
            return cor
    return COR_SEM_ACAO


_CENTROIDES = None


def centroides_municipios():
    """{município normalizado: (lon, lat)}: centro do maior anel."""
    global _CENTROIDES
    if _CENTROIDES is None:
        _CENTROIDES = {}
        geo = pac.carregar_geojson_municipios() or {"features": []}
        for feat in geo["features"]:
            maior = max(_poligonos(feat), key=lambda p: len(p[0]))[0]
            _CENTROIDES[normalizar(feat["properties"].get("name", ""))] = (
                sum(p[0] for p in maior) / len(maior),
                sum(p[1] for p in maior) / len(maior))
    return _CENTROIDES


def pontos_das_obras(df):
    """[(lon, lat, rótulo)]: um pin por município (a base não tem
    coordenadas)."""
    centros = centroides_municipios()
    pontos = []
    for mun_norm, nome in (df[["MUN_NORM", "MUNICIPIO"]]
                           .drop_duplicates("MUN_NORM").itertuples(index=False)):
        centro = centros.get(mun_norm)
        if centro:
            pontos.append((centro[0], centro[1], nome))
    return pontos


def mapa_coropletico(df, largura, altura, pontos=None, compacto=False,
                     coluna_valor="VALOR"):
    """Bahia colorida pelo valor somado por município (medidas do SSP)."""
    d = _desenho_ampliado(largura, altura)
    f, fb = _fonte(), _fonte_negrito()
    geo = pac.carregar_geojson_municipios()
    if not geo:
        d.add(String(largura / 2, altura / 2,
                     "Mapa indisponível: falta municipios_bahia.geojson",
                     textAnchor="middle", fontName=f, fontSize=8))
        return d
    feats = geo["features"]
    valores = df.groupby("MUN_NORM")[coluna_valor].sum()

    lon_min = lat_min = float("inf")
    lon_max = lat_max = float("-inf")
    for feat in feats:
        for poly in _poligonos(feat):
            for x, y in poly[0]:
                lon_min, lon_max = min(lon_min, x), max(lon_max, x)
                lat_min, lat_max = min(lat_min, y), max(lat_max, y)

    larg_legenda = 0 if compacto else 158
    area_l = largura - larg_legenda
    esc = min(area_l / (lon_max - lon_min), altura / (lat_max - lat_min))
    off_x = (area_l - (lon_max - lon_min) * esc) / 2
    off_y = (altura - (lat_max - lat_min) * esc) / 2

    def px(lon, lat):
        return off_x + (lon - lon_min) * esc, off_y + (lat - lat_min) * esc

    com_acao = 0
    for feat in feats:
        nome = normalizar(feat["properties"].get("name", ""))
        valor = float(valores.get(nome, 0.0))
        tem = nome in valores.index and valor > 0
        com_acao += 1 if tem else 0
        cor = cor_da_classe(valor) if tem else COR_SEM_ACAO
        for poly in _poligonos(feat):
            anel = poly[0]
            p = Path(fillColor=cor, strokeColor=colors.white, strokeWidth=0.3)
            x0, y0 = px(*anel[0])
            p.moveTo(x0, y0)
            for lon, lat in anel[1:]:
                p.lineTo(*px(lon, lat))
            p.closePath()
            d.add(p)

    escuro = pac.COR_MARCA_TEAL_ESCURO
    if pontos:
        raio_pin = 3.2 if compacto else 3.8
        for lon, lat, rotulo in pontos:
            x, y = px(lon, lat)
            d.add(Path([x - raio_pin * 0.72, y, x, y - raio_pin * 1.9,
                        x + raio_pin * 0.72, y], [0, 1, 1],
                       fillColor=escuro, strokeColor=None))
            d.add(Circle(x, y, raio_pin, fillColor=escuro,
                         strokeColor=colors.white, strokeWidth=0.8))
            d.add(Circle(x, y, raio_pin * 0.42, fillColor=colors.white,
                         strokeColor=None))
            if compacto and len(pontos) <= 20:
                d.add(String(x + raio_pin + 3, y - 3, rotulo, fontName=fb,
                             fontSize=8.5, fillColor=escuro))

    vmax = float(valores.max()) if len(valores) else 0.0
    itens_legenda = legenda_classes_mapa(vmax)
    f_leg = 8 if compacto else 9.5
    quad = 10 if compacto else 12
    passo = 16 if compacto else 22
    if compacto:
        lx = 4
        ly = passo * (len(itens_legenda) + (2 if pontos else 1)) + 26
    else:
        lx = area_l + 8
        ly = altura - 24
    for cor, rotulo in itens_legenda:
        d.add(Rect(lx, ly, quad, quad, fillColor=cor, strokeColor=None))
        d.add(String(lx + quad + 5, ly + 1.5, rotulo, fontName=fb,
                     fontSize=f_leg, fillColor=COR_TEXTO))
        ly -= passo
    ly -= (4 if compacto else 8)
    d.add(Rect(lx, ly, quad, quad, fillColor=COR_SEM_ACAO, strokeColor=None))
    d.add(String(lx + quad + 5, ly + 1.5, "Sem ação no filtro", fontName=f,
                 fontSize=f_leg, fillColor=COR_TEXTO))
    if pontos:
        ly -= passo
        d.add(Circle(lx + quad / 2, ly + quad / 2, quad * 0.42,
                     fillColor=escuro, strokeColor=colors.white,
                     strokeWidth=0.8))
        d.add(String(lx + quad + 5, ly + 1.5, "Município com ação",
                     fontName=f, fontSize=f_leg, fillColor=COR_TEXTO))
    ly -= (18 if compacto else 28)
    d.add(String(lx, ly, f"{com_acao} de {len(feats)} municípios",
                 fontName=fb, fontSize=f_leg, fillColor=escuro))
    d.add(String(lx, ly - (9 if compacto else 12), "com ação no filtro",
                 fontName=f, fontSize=f_leg - 1, fillColor=COR_TEXTO))
    return d


# ----------------------------------------------------------------------------
# PÁGINAS
# ----------------------------------------------------------------------------

def _tabela_sem_borda(linhas, larguras, estilo_extra=()):
    t = Table(linhas, colWidths=larguras)
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ] + list(estilo_extra)))
    return t


def pagina_painel_geral(historia, df, estilos):
    s = escala()
    historia.append(titulo_secao("PAINEL GERAL", estilos))
    historia.append(Spacer(1, 3 * 2.835 * s))

    util_ssp = LARGURA_UTIL_SSP
    largura_card = (util_ssp - 3 * 17) / 4
    concl = df[df["GRUPO_STATUS"].isin(GRUPOS_CONCLUIDOS)]
    andamento = df[df["GRUPO_STATUS"] == GRUPO_ANDAMENTO]
    cards = _tabela_sem_borda([[
        card_kpi("TOTAL DE AÇÕES", f"{len(df)}", largura_card, 54),
        card_kpi("INVESTIMENTO TOTAL", mi_bi(df["VALOR"].sum()),
                 largura_card, 54),
        card_kpi("CONCLUÍDAS / INAUGURADAS", f"{len(concl)}", largura_card, 54),
        card_kpi("EM ANDAMENTO", f"{len(andamento)}", largura_card, 54),
    ]], [(largura_card + 17) * s] * 3 + [largura_card * s])
    historia.append(cards)
    historia.append(Spacer(1, 4 * s))
    historia.append(Paragraph(
        "* O tamanho de cada fatia ou barra corresponde ao valor financeiro "
        "(R$), e os valores entre parênteses ( ) indicam a quantidade de "
        "ações. O valor é o investimento e, na falta dele, o orçamento.",
        estilos["nota"]))
    historia.append(Spacer(1, 6 * s))

    ordem = {g: i for i, g in enumerate(ORDEM_GRUPOS)}
    resumo_status = sorted(
        [(g, q, v) for g, q, v in resumir_por(df, "GRUPO_STATUS") if q > 0],
        key=lambda t: ordem.get(t[0], 99))
    resumo_eixo = [(o, q, v) for o, q, v in resumir_por(df, "EIXO") if v > 0]

    larg_esq = util_ssp * 0.46
    larg_dir = util_ssp - larg_esq
    alt_donut = 156
    alt_barras = 124
    esquerda = Table(
        [[Paragraph("PANORAMA POR STATUS *", estilos["titulo_barra"])],
         [donut_status(resumo_status, larg_esq - 6, alt_donut)],
         [Spacer(1, 6 * s)],
         [Paragraph("INVESTIMENTO POR EIXO *", estilos["titulo_barra"])],
         [barras_eixo(resumo_eixo, larg_esq - 6, alt_barras)]],
        colWidths=[larg_esq * s])
    esquerda.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), pac.COR_MARCA_TEAL_CLARA),
        ("LINEBELOW", (0, 0), (0, 0), 1.2 * s, pac.COR_MARCA_GOLD),
        ("BACKGROUND", (0, 3), (0, 3), pac.COR_MARCA_TEAL_CLARA),
        ("LINEBELOW", (0, 3), (0, 3), 1.2 * s, pac.COR_MARCA_GOLD),
        ("LEFTPADDING", (0, 0), (-1, -1), 6 * s),
        ("TOPPADDING", (0, 0), (-1, -1), 3 * s),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * s),
    ]))
    alt_mapa = alt_donut + alt_barras + 24
    direita = Table(
        [[Paragraph("DISTRIBUIÇÃO POR MUNICÍPIO: INVESTIMENTO E ORÇAMENTO",
                    estilos["titulo_barra"])],
         [mapa_coropletico(df, larg_dir - 10, alt_mapa)]],
        colWidths=[larg_dir * s])
    direita.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), pac.COR_MARCA_TEAL_CLARA),
        ("LINEBELOW", (0, 0), (0, 0), 1.2 * s, pac.COR_MARCA_GOLD),
        ("LEFTPADDING", (0, 0), (-1, -1), 6 * s),
        ("TOPPADDING", (0, 0), (-1, -1), 3 * s),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * s),
    ]))
    historia.append(_tabela_sem_borda([[esquerda, direita]],
                                      [larg_esq * s, larg_dir * s]))
    historia.append(PageBreak())


def pagina_resumo_financeiro(historia, df, estilos):
    s = escala()
    historia.append(titulo_secao(
        "RESUMO FINANCEIRO: INVESTIMENTO POR EIXO E SITUAÇÃO", estilos))
    historia.append(Spacer(1, 6 * s))
    historia.append(Paragraph(
        "* O comprimento de cada barra corresponde ao investimento total do "
        "eixo; as faixas coloridas mostram quanto está em cada situação.",
        estilos["nota"]))
    historia.append(Spacer(1, 6 * s))
    altura_ssp = _altura() / s - 62
    historia.append(barras_empilhadas_eixo(df, LARGURA_UTIL_SSP, altura_ssp))
    historia.append(PageBreak())


def _texto_avanco(v):
    return "-" if pd.isna(v) else f"{float(v):.1f}%".replace(".", ",")


def _data_da_linha(r, eh_licitacao):
    data = r["AVISO_LICITACAO"] if eh_licitacao else r["ENTREGA"]
    return data.strftime("%d/%m/%Y") if pd.notna(data) else "-"


def chave_contrato(valor):
    texto = str(valor or "").strip()
    achado = re.search(r"(\d{1,4})\s*/\s*(\d{2,4})", texto)
    if not achado:
        return (1, 9999, 9999, normalizar(texto))
    numero, ano = int(achado.group(1)), int(achado.group(2))
    if ano < 100:
        ano += 2000
    return (0, ano, numero, normalizar(texto))


def ordenar(sub, ordem):
    if ordem == "VALOR_DESC":
        return sub.sort_values(["VALOR", "MUNICIPIO"], ascending=[False, True])
    if ordem == "AVANCO_DESC":
        return sub.sort_values(["AVANCO", "MUNICIPIO"], ascending=[False, True],
                               na_position="last")
    if ordem == "PRAZO_ASC":
        return sub.sort_values(["ENTREGA", "MUNICIPIO"], na_position="last")
    if ordem == "ACAO":
        return sub.sort_values(["ACAO", "MUNICIPIO"])
    if ordem == "CONTRATO":
        chaves = sub["NUMERO_CONTRATO"].map(chave_contrato)
        tmp = sub.assign(_k=chaves)
        return tmp.sort_values(["_k", "MUNICIPIO"]).drop(columns=["_k"])
    return sub.sort_values(["MUNICIPIO", "ACAO"])


def rotulo_ordem(ordem):
    return dict(OPCOES_ORDEM).get(ordem, OPCOES_ORDEM[0][1])


def _celulas(r, colunas, estilos, eh_licitacao):
    data = _data_da_linha(r, eh_licitacao)
    if r["DIAS_ATRASO"] > 0 and data != "-":   # atraso: data em vermelho
        data = (f'<font color="#{COR_VENCIDO.hexval()[2:]}"><b>{data}</b>'
                f"</font>")
    valor = r["VALOR_LICITACAO"] if eh_licitacao else r["VALOR"]
    por_rotulo = {
        "CÓDIGO": r["CODIGO"],
        "CONTRATO": r["NUMERO_CONTRATO"] or "-",
        "AÇÃO": r["ACAO"],
        "ESCOPO": r["ESCOPO"],
        "DETALHAMENTO ESCOPO": r["DETALHAMENTO_ESCOPO"] or "-",
        "OBJETO": r["OBJETO"],
        "EIXO": r["EIXO"],
        "TERRITÓRIO": r["TERRITORIO"],
        "MUNICÍPIO": r["MUNICIPIO"],
        "ORÇAMENTO": moeda(r["ORCAMENTO"]),
        "INVESTIMENTO (R$)": moeda(valor),
        "STATUS": r["STATUS"],
        "AVANÇO": _texto_avanco(r["AVANCO"]),
        "CONTRATADA": r["CONTRATADA"] or "-",
        "ENTREGA OBRA": data,
    }
    linha = []
    for rotulo, _p, alinh in colunas:
        texto = por_rotulo[rotulo]
        if rotulo != "ENTREGA OBRA":
            texto = _escapar(texto)
        linha.append(Paragraph(texto, estilos[alinh]))
    return linha


def _paginar_linhas(alturas, altura_disponivel):
    """Fatias [inicio, fim) por página: no máximo ACOES_POR_PAGINA_MAX
    linhas e nunca além da altura (a página vira antes, sem cortar linha)."""
    fatias, inicio = [], 0
    while inicio < len(alturas):
        usado, fim = 0.0, inicio
        while (fim < len(alturas) and fim - inicio < ACOES_POR_PAGINA_MAX
               and usado + alturas[fim] <= altura_disponivel):
            usado += alturas[fim]
            fim += 1
        if fim == inicio:          # linha mais alta que a página: sozinha
            fim = inicio + 1
        fatias.append((inicio, fim))
        inicio = fim
    return fatias


# Conferência da paginação (lida pela verificação): uma entrada por página
# do detalhamento com seção, quantidade de linhas e intervalo.
REGISTRO_PAGINAS = []


def paginas_detalhamento(historia, df, estilos, colunas_ativas=None,
                         ordem="MUNICIPIO"):
    s = escala()
    ativas = [c for c in (colunas_ativas or COLUNAS_PADRAO)
              if c in {r for r, _p, _a in COLUNAS_DETALHAMENTO}][:LIMITE_COLUNAS]
    if not ativas:
        ativas = list(COLUNAS_PADRAO)
    colunas_base = [c for c in COLUNAS_DETALHAMENTO if c[0] in ativas]
    colunas_lic = [(r, p * 1.5 if r == "INVESTIMENTO (R$)" else p, a)
                   for r, p, a in colunas_base if r not in FORA_NA_LICITACAO]

    e_salvador = df["MUN_NORM"] == "salvador"
    secoes = []
    for capitulo, mascara in (("SALVADOR", e_salvador),
                              ("INTERIOR", ~e_salvador)):
        for grupo in ORDEM_GRUPOS + sorted(set(df["GRUPO_STATUS"])
                                           - set(ORDEM_GRUPOS)):
            parte = df[(df["GRUPO_STATUS"] == grupo) & mascara]
            if not parte.empty:
                secoes.append((grupo, capitulo, parte))

    altura_titulo = 30 * s          # faixa de título + respiro
    altura_tabela = _altura() - altura_titulo - 6
    del REGISTRO_PAGINAS[:]

    for grupo, capitulo, sub in secoes:
        sub = ordenar(sub, ordem)
        eh_licitacao = sub["CODIGO_STATUS"].iloc[0] in CODIGOS_LICITACAO
        colunas = colunas_lic if eh_licitacao else colunas_base
        util = _largura()
        larg_tabela = util * 0.66 if eh_licitacao else util
        soma = sum(c[1] for c in colunas)
        larguras = [c[1] / soma * larg_tabela for c in colunas]
        cor = COR_POR_GRUPO.get(grupo, colors.grey)

        def rotulo_cab(nome):
            if eh_licitacao and nome == "INVESTIMENTO (R$)":
                return "ORÇAMENTO (R$)"
            if eh_licitacao and nome == "ENTREGA OBRA":
                return "AVISO DE LICITAÇÃO"
            return nome

        cab = [Paragraph(rotulo_cab(r), estilos["cab"]) for r, _p, _a in colunas]
        corpo = [_celulas(r, colunas, estilos, eh_licitacao)
                 for _i, r in sub.iterrows()]

        # alturas reais das linhas, medidas uma vez na tabela inteira
        medida = Table([cab] + corpo, colWidths=larguras)
        medida.setStyle(_estilo_tabela(cor))
        medida.wrap(larg_tabela, 10 ** 7)
        alturas = list(medida._rowHeights)
        fatias = _paginar_linhas(alturas[1:], altura_tabela - alturas[0])

        total_valor = float(sub["VALOR_LICITACAO" if eh_licitacao
                                else "VALOR"].sum())
        sufixo = " · valores de orçamento" if eh_licitacao else ""
        titulo = (f"{grupo} · {capitulo} ({len(sub)} | "
                  f"{mi_bi(total_valor)}){sufixo}")
        mapa = None
        if eh_licitacao:
            larg_mapa = (util - larg_tabela) / s
            mapa = mapa_coropletico(
                sub, larg_mapa - 6, altura_tabela / s - 6,
                pontos=pontos_das_obras(sub), compacto=True,
                coluna_valor="VALOR_LICITACAO")

        for indice, (ini, fim) in enumerate(fatias):
            texto = titulo + (" (cont.)" if indice else "")
            if len(fatias) > 1:
                texto += f" · ações {ini + 1} a {fim} de {len(sub)}"
            historia.append(titulo_secao(texto, estilos))
            historia.append(Spacer(1, 4 * s))
            tabela = Table([cab] + corpo[ini:fim], colWidths=larguras)
            tabela.setStyle(_estilo_tabela(cor))
            if mapa is not None:
                lado = Table([[tabela, mapa]],
                             colWidths=[larg_tabela, util - larg_tabela])
                lado.setStyle(TableStyle([
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (0, 0), 6),
                    ("RIGHTPADDING", (1, 0), (1, 0), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]))
                lado.hAlign = "LEFT"
                historia.append(lado)
            else:
                historia.append(tabela)
            historia.append(PageBreak())
            REGISTRO_PAGINAS.append({
                "secao": f"{grupo} · {capitulo}", "linhas": fim - ini,
                "inicio": ini + 1, "fim": fim, "total": len(sub),
                "altura": sum(alturas[ini + 1:fim + 1]) + alturas[0],
                "limite": altura_tabela, "ultima": indice == len(fatias) - 1,
            })


def _estilo_tabela(cor):
    s = escala()
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), cor),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COR_LINHA_ALTERNADA]),
        ("GRID", (0, 0), (-1, -1), 0.35 * s, COR_CINZA_LINHA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        # respiro vertical mínimo: a fonte fica no tamanho do SSP e o
        # ganho de linhas por página vem daqui
        ("TOPPADDING", (0, 0), (-1, -1), 1 * s),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6 * s),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5 * s),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5 * s),
    ])


# ----------------------------------------------------------------------------
# DIVISÓRIA, CABEÇALHO E MONTAGEM DA SEÇÃO
# ----------------------------------------------------------------------------

class _TrocarCabecalho(Flowable):
    """Flowable invisível: define qual cabeçalho as PRÓXIMAS páginas usam
    (o cabeçalho de uma página é desenhado antes do conteúdo dela)."""

    def __init__(self, modo):
        Flowable.__init__(self)
        self.modo = modo
        self.width = self.height = 0

    def wrap(self, *args):
        return (0, 0)

    def draw(self):
        pac.MODO_CABECALHO_PDF = self.modo


def desenhar_divisoria(canvas, doc):
    """Página de abertura da seção, no mesmo visual da capa do PAC."""
    canvas.saveState()
    largura, altura = pac.PAGINA
    try:
        canvas.drawImage(pac.caminho_recurso("GOVERNO.PNG"),
                         largura - pac.MARGEM_DIR - 260, pac.MARGEM_ESQ,
                         width=260, height=286, mask="auto",
                         preserveAspectRatio=True, anchor="c")
    except Exception:
        pass
    x = pac.MARGEM_ESQ + 40
    canvas.setFillColor(pac.COR_MARCA_TEAL_ESCURO)
    canvas.setFont(_fonte_negrito(), 96)
    canvas.drawString(x, altura * 0.60, ROTULO_SECAO)
    canvas.setFillColor(colors.HexColor("#444444"))
    canvas.setFont(_fonte(), 44)
    canvas.drawString(x, altura * 0.60 - 70, "RESUMO DAS AÇÕES DO CRONOGRAMA")
    canvas.setStrokeColor(pac.COR_MARCA_TEAL)
    canvas.setLineWidth(3)
    canvas.line(x, altura * 0.60 - 100, largura / 2, altura * 0.60 - 100)
    canvas.setStrokeColor(pac.COR_MARCA_GOLD)
    canvas.line(largura / 2, altura * 0.60 - 100, largura - pac.MARGEM_DIR - 40,
                altura * 0.60 - 100)
    canvas.setFillColor(colors.black)
    canvas.setFont(_fonte(), 16)
    canvas.drawString(x, altura * 0.60 - 130, texto_atualizacao())
    canvas.restoreState()


def desenhar_cabecalho(canvas, doc):
    """Cabeçalho das páginas da seção: logo do Governo, linha teal/dourada
    do PAC e o nome da seção no lugar da logo do PAC."""
    canvas.saveState()
    largura, altura = pac.PAGINA
    try:
        canvas.drawImage(pac.caminho_recurso("GOVERNO.PNG"), pac.MARGEM_ESQ,
                         altura - 37, width=100, height=35, mask="auto",
                         preserveAspectRatio=True, anchor="c")
    except Exception:
        pass
    canvas.setFillColor(pac.COR_MARCA_TEAL_ESCURO)
    canvas.setFont(_fonte_negrito(), 15)
    canvas.drawRightString(largura - pac.MARGEM_DIR, altura - 30, ROTULO_SECAO)
    y = altura - 46
    canvas.setLineWidth(2)
    canvas.setStrokeColor(pac.COR_MARCA_TEAL)
    canvas.line(pac.MARGEM_ESQ, y, largura / 2, y)
    canvas.setStrokeColor(pac.COR_MARCA_GOLD)
    canvas.line(largura / 2, y, largura - pac.MARGEM_DIR, y)
    canvas.setFont(pac.FONTE_PADRAO, 13)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(largura / 2, 25, texto_atualizacao())
    canvas.restoreState()


def secoes_disponiveis(filtros):
    """Itens da janela de seleção de páginas (vazio se não houver base ou
    o recorte do MONITORAMENTO estiver vazio)."""
    df = filtrar(filtros)
    if df is None or df.empty:
        return []
    # Chegam desmarcadas na janela (pedido do usuário em 11/11/2026): o
    # MONITORAMENTO só entra no PDF quando alguém o marca.
    return [dict(s, paginas=None if s["chave"] == "MON_DETALHAMENTO" else 1,
                 marcada=False) for s in SECOES]


def elementos_pdf(filtros, secoes_marcadas, marcador_gestao):
    """Flowables da seção (divisória incluída), ou [] se nada sai.
    marcador_gestao: classe _MarcadorGestao do PAC (numeração por seção)."""
    df = filtrar(filtros)
    chaves = [s["chave"] for s in SECOES
              if secoes_marcadas is None or s["chave"] in secoes_marcadas]
    if df is None or df.empty or not chaves:
        return []
    filtros = filtros or {}
    estilos = _estilos()
    historia = [marcador_gestao(None), Spacer(1, 1),
                _TrocarCabecalho("MONITORAMENTO"), PageBreak(),
                marcador_gestao(ROTULO_SECAO)]
    if "MON_PAINEL_GERAL" in chaves:
        pagina_painel_geral(historia, df, estilos)
    if "MON_RESUMO_FINANCEIRO" in chaves:
        pagina_resumo_financeiro(historia, df, estilos)
    if "MON_DETALHAMENTO" in chaves:
        paginas_detalhamento(historia, df, estilos,
                             filtros.get("COLUNAS"), filtros.get("ORDEM")
                             or "MUNICIPIO")
    return historia
