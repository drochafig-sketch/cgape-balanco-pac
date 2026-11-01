# -*- coding: utf-8 -*-
"""Gera o PPTX de Análise Financeira do PAC para apresentação ao Governo Federal.

Reexecutável: lê a base atual (importa "CGAPE - BALANÇO PAC.py" como módulo,
igual ao servidor_web.py) e recalcula todos os números — quando a planilha
for atualizada, basta rodar este script de novo.

Dois jeitos de usar:
  1. Linha de comando: "python gerar_pptx_pac.py" — carrega a base do zero
     (bootstrap próprio, igual ao servidor_web.py) e salva os dois arquivos
     nesta pasta.
  2. Chamado pelo botão "ANÁLISE FINANCEIRA (PPTX)" do próprio painel (ver
     _construir_pptx_analise_financeira_bytes em "CGAPE - BALANÇO PAC.py"):
     quem chama já tem o módulo carregado (sem reler a planilha de novo) e
     passa a referência via configurar(pac_ja_carregado) antes de chamar
     gerar_pptx_bytes(df).

Escopo e definições decididas com o usuário em 01/11/2026 (ver conversa):
  - INVESTIMENTO/CONTRATADO/A CONTRATAR seguem a mesma regra "ajustada" já
    usada no resto do app (ver _serie_investimento_ajustado no .py original):
    Valor Contratado substitui os 4 componentes quando preenchido, pra não
    contar o mesmo dinheiro duas vezes.
  - FINANCIAMENTO é tratado como parcela FEDERAL (junto com OGU/Apoiado) —
    decisão do usuário, não é um fato da base (a planilha não diferencia
    dívida do Estado com aval federal de financiamento federal direto).
  - Sem dado de desembolso/liberação/empenho na base: os slides NÃO fazem
    nenhuma promessa de "execução financeira" além de Contratado x A
    Contratar x Detalhamento por fonte.
  - Sem cronograma anual confiável (só 30-35% das ações têm data prevista):
    o slide de "necessidade por ano" foi removido deste escopo.
  - "Não discriminado" (contratado sem detalhamento por fonte) aparece como
    categoria própria, não é estimado nem redistribuído.
Saída: ANALISE_FINANCEIRA_PAC.pptx e APOIO_ANALISE_FINANCEIRA_PAC.xlsx, na
mesma pasta deste script.
"""
import importlib.util
import os
import sys
from datetime import date

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------------------
# Referência ao módulo "CGAPE - BALANÇO PAC.py" já carregado (col_valor_
# contratado, df_original, _campos_alerta_qualidade etc.) — None até alguém
# chamar configurar(). Fica só numa variável de módulo (em vez de passada
# explicitamente por parâmetro em cada função) pra não precisar reescrever a
# assinatura de toda função que já usa "pac.xxx" só por causa disso.
# ---------------------------------------------------------------------------
PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
pac = None


def configurar(pac_ja_carregado):
    # Usado pelo botão do painel (_construir_pptx_analise_financeira_bytes em
    # "CGAPE - BALANÇO PAC.py"): injeta o módulo que JÁ está rodando (com a
    # planilha já lida em memória), em vez de recarregar tudo do zero — ver
    # bootstrap próprio em main(), usado só no modo linha de comando.
    global pac
    pac = pac_ja_carregado

# ---------------------------------------------------------------------------
# Identidade visual (paleta fixa — mesma do pedido original)
# ---------------------------------------------------------------------------
NAVY = "0B1F3F"
AZUL = "0B5CAD"
VERDE = "2E7D32"
LARANJA = "E68A00"
VERMELHO = "C62828"
TEAL = "2F7F92"
CINZA_TEXTO = "5B6576"
CINZA_CLARO = "F2F4F7"
CINZA_BORDA = "D8DCE3"
BRANCO = "FFFFFF"
AZUL_CLARO_TEXTO = "C9D6EA"
AZUL_CLARO_MARCA = "4FA3E0"
PENDENCIA_FUNDO = "FDECEC"

COR_OGU = AZUL
COR_FINANCIAMENTO = TEAL
COR_ESTADUAL = LARANJA
COR_NAO_DISCRIMINADO = "8C8C8C"

FONTE = "Roboto"

# Canvas: igual ao "APRESENTAÇÃO - CGAPE.pptx" já existente (20 x 11,25 pol,
# 16:9) — decisão do usuário, pra colar sem reescalar dentro do deck atual.
LARGURA_IN = 20.0
ALTURA_IN = 11.25


def pol(valor):
    return Emu(int(valor * 914400))


def cor(hexstr):
    return RGBColor.from_string(hexstr)


# ---------------------------------------------------------------------------
# Formatação de valores (regras do pedido: Mil/Mi/Bi, 1 casa, vírgula)
# ---------------------------------------------------------------------------
def abreviar_moeda(valor):
    valor = float(valor)
    sinal = "-" if valor < 0 else ""
    valor = abs(valor)
    if valor >= 1e9:
        texto = f"{valor / 1e9:.1f}".replace(".", ",") + " Bi"
    elif valor >= 1e6:
        texto = f"{valor / 1e6:.1f}".replace(".", ",") + " Mi"
    elif valor >= 1e3:
        texto = f"{valor / 1e3:.1f}".replace(".", ",") + " Mil"
    else:
        texto = f"{valor:.0f}"
    return f"{sinal}R$ {texto}"


def pct(parte, total):
    return f"{(parte / total * 100):.1f}".replace(".", ",") + "%" if total else "0,0%"


# ---------------------------------------------------------------------------
# Helpers de desenho (formas nativas — nada de gráfico OOXML, por causa da
# compatibilidade com a importação no Canva)
# ---------------------------------------------------------------------------
def add_rect(slide, x, y, w, h, fill=None, line=None, arredondado=False, line_w=None):
    forma = MSO_SHAPE.ROUNDED_RECTANGLE if arredondado else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(forma, pol(x), pol(y), pol(w), pol(h))
    if arredondado:
        try:
            shp.adjustments[0] = 0.06
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = cor(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = cor(line)
        shp.line.width = Pt(line_w or 1)
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, texto, tam, cor_hex, negrito=False, alinhar=PP_ALIGN.LEFT,
             fonte=FONTE, ancorar=MSO_ANCHOR.TOP, espaco_linha=None):
    caixa = slide.shapes.add_textbox(pol(x), pol(y), pol(w), pol(h))
    tf = caixa.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = ancorar
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    linhas = texto.split("\n")
    for i, linha in enumerate(linhas):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = alinhar
        if espaco_linha:
            p.line_spacing = espaco_linha
        run = p.add_run()
        run.text = linha
        run.font.size = Pt(tam)
        run.font.bold = negrito
        run.font.name = fonte
        run.font.color.rgb = cor(cor_hex)
    return caixa


def add_tag_estimativa(slide, x, y, texto="ESTIMATIVA"):
    add_rect(slide, x, y, 1.1, 0.26, fill=CINZA_CLARO, line=CINZA_BORDA, arredondado=True, line_w=0.75)
    add_text(slide, x, y + 0.02, 1.1, 0.22, texto, 9, CINZA_TEXTO, negrito=True,
              alinhar=PP_ALIGN.CENTER, ancorar=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------------------
# Slide base: cabeçalho navy + rodapé — igual em todo slide
# ---------------------------------------------------------------------------
def novo_slide(prs, tema, pagina, total_paginas):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # layout em branco

    # fundo branco explícito
    add_rect(slide, 0, 0, LARGURA_IN, ALTURA_IN, fill=BRANCO)

    # faixa do cabeçalho
    add_rect(slide, 0, 0, LARGURA_IN, 1.1, fill=NAVY)
    add_text(slide, 0.56, 0.22, 12.5, 0.5, f"PAC - {tema}", 26, BRANCO, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    add_text(slide, 0.56, 0.74, 12.5, 0.3, "COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE PROJETOS ESTRUTURANTES",
              12, AZUL_CLARO_TEXTO, negrito=True)
    add_text(slide, LARGURA_IN - 5.0, 0.18, 4.44, 0.35, "GOVERNO DO ESTADO DA BAHIA", 15,
              AZUL_CLARO_MARCA, negrito=True, alinhar=PP_ALIGN.RIGHT)
    add_text(slide, LARGURA_IN - 5.0, 0.56, 4.44, 0.35, "CASA CIVIL", 15, AZUL_CLARO_MARCA,
              negrito=True, alinhar=PP_ALIGN.RIGHT)

    # rodapé
    add_text(slide, 0.56, ALTURA_IN - 0.42, 10.0, 0.3,
              f"Fonte: {os.path.basename(pac.CAMINHO_PLANILHA) if hasattr(pac, 'CAMINHO_PLANILHA') else 'PANORAMA - PAC ORIGINAL - PAC SELEÇÕES - 2026'} · "
              f"posição em {pac.ultima_atualizacao.strftime('%m/%Y') if hasattr(pac, 'ultima_atualizacao') else date.today().strftime('%m/%Y')}",
              10, CINZA_TEXTO)
    add_text(slide, LARGURA_IN - 2.0, ALTURA_IN - 0.42, 1.44, 0.3, f"{pagina}/{total_paginas}", 10,
              CINZA_TEXTO, negrito=True, alinhar=PP_ALIGN.RIGHT)

    return slide


def add_pendencia_tag(slide, x, y, w=1.8, h=0.4):
    add_rect(slide, x, y, w, h, fill=PENDENCIA_FUNDO, line=VERMELHO, arredondado=True, line_w=1.25)
    add_text(slide, x, y, w, h, "⚠ PENDÊNCIA", 11, VERMELHO, negrito=True, alinhar=PP_ALIGN.CENTER,
              ancorar=MSO_ANCHOR.MIDDLE)


def barra_horizontal(slide, x, y, largura_max, altura, rotulo, valor, total, cor_barra, largura_rotulo=4.2,
                      tam_rotulo=13, largura_valor=2.6):
    # Tudo (rótulo + barra + valor) precisa caber dentro de largura_max, a
    # partir de x — senão o rótulo de valor estoura a borda do slide.
    fracao = (valor / total) if total else 0
    add_text(slide, x, y, largura_rotulo, altura, rotulo, tam_rotulo, NAVY, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    x_barra = x + largura_rotulo + 0.2
    largura_disponivel = largura_max - largura_rotulo - 0.2 - largura_valor - 0.15
    add_rect(slide, x_barra, y + altura * 0.18, largura_disponivel, altura * 0.64, fill=CINZA_CLARO)
    if fracao > 0:
        add_rect(slide, x_barra, y + altura * 0.18, max(largura_disponivel * fracao, 0.05), altura * 0.64,
                  fill=cor_barra)
    add_text(slide, x_barra + largura_disponivel + 0.15, y, largura_valor, altura,
              f"{abreviar_moeda(valor)}  ({pct(valor, total)})", 12, NAVY, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)


def tabela_nativa(slide, x, y, w, h, cabecalho, linhas, larguras=None, tam_fonte=12,
                   linha_bold_indices=None):
    linha_bold_indices = linha_bold_indices or set()
    n_linhas = len(linhas) + 1
    n_cols = len(cabecalho)
    shp = slide.shapes.add_table(n_linhas, n_cols, pol(x), pol(y), pol(w), pol(h))
    tbl = shp.table
    if larguras:
        for i, lw in enumerate(larguras):
            tbl.columns[i].width = pol(lw)
    # cabeçalho
    for j, titulo in enumerate(cabecalho):
        celula = tbl.cell(0, j)
        celula.fill.solid()
        celula.fill.fore_color.rgb = cor(NAVY)
        celula.margin_left = Pt(6)
        celula.margin_right = Pt(6)
        tf = celula.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = titulo
        run.font.size = Pt(tam_fonte)
        run.font.bold = True
        run.font.name = FONTE
        run.font.color.rgb = cor(BRANCO)
    # linhas
    for i, linha in enumerate(linhas):
        zebra = (i % 2 == 1)
        negrito_linha = i in linha_bold_indices
        for j, valor_cel in enumerate(linha):
            celula = tbl.cell(i + 1, j)
            celula.fill.solid()
            celula.fill.fore_color.rgb = cor(CINZA_CLARO if zebra else BRANCO)
            celula.margin_left = Pt(6)
            celula.margin_right = Pt(6)
            tf = celula.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            run = p.add_run()
            run.text = str(valor_cel)
            run.font.size = Pt(tam_fonte)
            run.font.bold = negrito_linha
            run.font.name = FONTE
            run.font.color.rgb = cor(NAVY)
    return shp


# ---------------------------------------------------------------------------
# Métricas (tudo recalculado a partir de df — nada hardcoded)
# ---------------------------------------------------------------------------
def calcular_metricas(df):
    contratado = df[pac.col_valor_contratado]
    apoiado_aj = df[pac.col_apoiado_ajustado]
    financiamento_aj = df[pac.col_financiamento_ajustado]
    contrapartida_aj = df[pac.col_contrapartida_ajustado]
    complementar_aj = df[pac.col_complementar_ajustado]
    a_contratar_linha = apoiado_aj + financiamento_aj + contrapartida_aj + complementar_aj
    investimento_linha = contratado + a_contratar_linha

    m = {}
    m["qtd_acoes"] = int(len(df))
    m["total_investimento"] = float(investimento_linha.sum())
    m["total_contratado"] = float(contratado.sum())
    m["total_a_contratar"] = float(a_contratar_linha.sum())

    # a contratar por fonte (decisão do usuário: financiamento = Federal)
    m["a_contratar_ogu"] = float(apoiado_aj.sum())
    m["a_contratar_financiamento"] = float(financiamento_aj.sum())
    m["a_contratar_estadual"] = float((contrapartida_aj + complementar_aj).sum())
    m["a_contratar_federal"] = m["a_contratar_ogu"] + m["a_contratar_financiamento"]

    # detalhamento do contratado (colunas novas)
    m["contratado_ogu"] = float(df[pac.col_apoiado_contratado].sum())
    m["contratado_financiamento"] = float(df[pac.col_financiamento_contratado].sum())
    m["contratado_estadual"] = float(
        (df[pac.col_contrapartida_contratado] + df[pac.col_complementar_contratado]).sum()
    )
    m["contratado_federal"] = m["contratado_ogu"] + m["contratado_financiamento"]
    soma_detalhado = m["contratado_ogu"] + m["contratado_financiamento"] + m["contratado_estadual"]
    m["contratado_nao_discriminado"] = max(0.0, m["total_contratado"] - soma_detalhado)
    m["cobertura_detalhamento_pct"] = soma_detalhado / m["total_contratado"] if m["total_contratado"] else 0

    # Federal x Estadual (geral — contratado + a contratar)
    m["total_federal"] = m["a_contratar_federal"] + m["contratado_federal"]
    m["total_estadual"] = m["a_contratar_estadual"] + m["contratado_estadual"]

    # contadores por status (4 categorias do padrão visual)
    concluido_mask = df["STATUS_TEXTO"].isin(["CONCLUÍDA", "INAUGURADA"])
    andamento_mask = df["STATUS_TEXTO"] == "ANDAMENTO"
    paralisado_mask = df["STATUS_TEXTO"] == "PARALISADA"
    preparatoria_mask = ~(concluido_mask | andamento_mask | paralisado_mask)
    m["qtd_concluido"] = int(concluido_mask.sum())
    m["qtd_andamento"] = int(andamento_mask.sum())
    m["qtd_preparatoria"] = int(preparatoria_mask.sum())
    m["qtd_paralisado"] = int(paralisado_mask.sum())

    # por eixo
    df_tmp = df.assign(_CONTRATADO=contratado, _A_CONTRATAR=a_contratar_linha,
                        _INVESTIMENTO=investimento_linha)
    por_eixo = (
        df_tmp.groupby(pac.col_eixo)
        .agg(QTD=(pac.col_objeto, "count"), INVESTIMENTO=("_INVESTIMENTO", "sum"),
             CONTRATADO=("_CONTRATADO", "sum"), A_CONTRATAR=("_A_CONTRATAR", "sum"))
        .sort_values("INVESTIMENTO", ascending=False)
    )
    m["por_eixo"] = por_eixo

    # top 10 "a contratar"
    top10 = (
        df_tmp.assign(_A_CONTRATAR=a_contratar_linha)
        .sort_values("_A_CONTRATAR", ascending=False)
        .head(10)
    )
    m["top10_a_contratar"] = top10

    # pendências condicionantes
    m["qtd_clausula_vigente"] = int((df["SITUACAO_CLAUSULA_SUSPENSIVA"] == "Cláusula Vigente").sum())
    m["valor_clausula_vigente"] = float(
        contratado.loc[df["SITUACAO_CLAUSULA_SUSPENSIVA"] == "Cláusula Vigente"].sum()
    )
    m["qtd_obra_nao_iniciada"] = int((df["OBRA_NAO_INICIADA"] == "SIM").sum())
    m["qtd_prazo_pendencia_vencido"] = 0
    m["qtd_sem_detalhamento"] = 0
    hoje = pac.agora_bahia().date()
    for _, row in df.iterrows():
        alertas = pac._campos_alerta_qualidade(row, hoje)
        motivos = set(alertas.values())
        for mo in motivos:
            if "Prazo da Pendência" in mo:
                m["qtd_prazo_pendencia_vencido"] += 1
            if "categoria da origem do recurso" in mo:
                m["qtd_sem_detalhamento"] += 1

    m["gestao_estadual_qtd"] = int((df["GESTAO"] == "GESTÃO ESTADUAL").sum())
    m["gestao_federal_qtd"] = int((df["GESTAO"] == "GESTÃO FEDERAL").sum())

    return m


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------
TOTAL_PAGINAS = 9


def slide_1_panorama(prs, m):
    s = novo_slide(prs, "PANORAMA DA CARTEIRA", 1, TOTAL_PAGINAS)
    add_text(s, 0.9, 1.5, 10, 0.5, "INVESTIMENTO TOTAL DA CARTEIRA", 16, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, 1.85, 10, 1.3, abreviar_moeda(m["total_investimento"]), 60, VERDE, negrito=True)
    add_text(s, 0.9, 3.05, 10, 0.4, f"{m['qtd_acoes']} ações do PAC acompanhadas pelo Estado da Bahia", 14,
              CINZA_TEXTO)

    # barra empilhada contratado x a contratar
    y_barra = 3.9
    add_text(s, 0.9, y_barra - 0.4, 6, 0.3, "CONTRATADO  x  A CONTRATAR", 13, NAVY, negrito=True)
    largura_barra = 17.0
    x0 = 0.9
    frac_contratado = m["total_contratado"] / m["total_investimento"] if m["total_investimento"] else 0
    add_rect(s, x0, y_barra, largura_barra * frac_contratado, 0.9, fill=VERDE)
    add_rect(s, x0 + largura_barra * frac_contratado, y_barra, largura_barra * (1 - frac_contratado), 0.9,
              fill=LARANJA)
    add_text(s, x0 + 0.15, y_barra + 0.12, 6, 0.3, f"CONTRATADO · {abreviar_moeda(m['total_contratado'])}", 14,
              BRANCO, negrito=True)
    add_text(s, x0 + 0.15, y_barra + 0.48, 6, 0.3, pct(m["total_contratado"], m["total_investimento"]), 13,
              BRANCO)
    x_txt2 = x0 + largura_barra * frac_contratado + 0.15
    add_text(s, x_txt2, y_barra + 0.12, 7, 0.3, f"A CONTRATAR · {abreviar_moeda(m['total_a_contratar'])}", 14,
              BRANCO, negrito=True)
    add_text(s, x_txt2, y_barra + 0.48, 7, 0.3, pct(m["total_a_contratar"], m["total_investimento"]), 13,
              BRANCO)

    # contadores por status
    y_cards = 5.4
    cards = [
        ("CONCLUÍDO / INAUGURADO", m["qtd_concluido"], VERDE),
        ("EM OBRAS / ANDAMENTO", m["qtd_andamento"], AZUL),
        ("SELEÇÃO / PREPARATÓRIA / LICITAÇÃO", m["qtd_preparatoria"], LARANJA),
        ("PARALISADO", m["qtd_paralisado"], VERMELHO),
    ]
    largura_card = 4.1
    espaco = 0.3
    for i, (rotulo, qtd, c) in enumerate(cards):
        x = 0.9 + i * (largura_card + espaco)
        add_rect(s, x, y_cards, largura_card, 2.0, fill=BRANCO, line=CINZA_BORDA, arredondado=True, line_w=1)
        add_rect(s, x, y_cards, largura_card, 0.12, fill=c, arredondado=False)
        add_text(s, x + 0.25, y_cards + 0.35, largura_card - 0.5, 0.9, str(qtd), 44, c, negrito=True)
        add_text(s, x + 0.25, y_cards + 1.35, largura_card - 0.5, 0.55, rotulo, 12, CINZA_TEXTO, negrito=True)
    return s


def slide_2_composicao_origem(prs, m):
    s = novo_slide(prs, "COMPOSIÇÃO POR ORIGEM DO RECURSO", 2, TOTAL_PAGINAS)
    total = m["total_investimento"]
    y = 1.9
    add_text(s, 0.9, y, 10, 0.35, "INVESTIMENTO TOTAL, POR FONTE (contratado + a contratar)", 14, CINZA_TEXTO,
              negrito=True)
    y += 0.55
    fontes = [
        ("OGU (Apoiado)", m["contratado_ogu"] + m["a_contratar_ogu"], COR_OGU),
        ("FINANCIAMENTO", m["contratado_financiamento"] + m["a_contratar_financiamento"], COR_FINANCIAMENTO),
        ("RECURSO ESTADUAL (Contrapartida + Complementar)", m["contratado_estadual"] + m["a_contratar_estadual"],
         COR_ESTADUAL),
    ]
    for rotulo, valor, c in fontes:
        barra_horizontal(s, 0.9, y, 17.6, 0.75, rotulo, valor, total, c, largura_rotulo=5.2)
        y += 1.0

    # destaque federal x estadual
    y += 0.3
    add_rect(s, 0.9, y, 8.4, 1.6, fill=CINZA_CLARO, arredondado=True)
    add_text(s, 1.2, y + 0.18, 7.8, 0.3, "PARCELA FEDERAL (OGU + Financiamento)*", 12, CINZA_TEXTO, negrito=True)
    add_text(s, 1.2, y + 0.5, 7.8, 0.8, abreviar_moeda(m["total_federal"]), 32, AZUL, negrito=True)
    add_text(s, 1.2, y + 1.15, 7.8, 0.3, pct(m["total_federal"], total) + " do investimento total", 12,
              CINZA_TEXTO)

    add_rect(s, 9.6, y, 8.4, 1.6, fill=CINZA_CLARO, arredondado=True)
    add_text(s, 9.9, y + 0.18, 7.8, 0.3, "PARCELA ESTADUAL (Contrapartida + Complementar)", 12, CINZA_TEXTO,
              negrito=True)
    add_text(s, 9.9, y + 0.5, 7.8, 0.8, abreviar_moeda(m["total_estadual"]), 32, LARANJA, negrito=True)
    add_text(s, 9.9, y + 1.15, 7.8, 0.3, pct(m["total_estadual"], total) + " do investimento total", 12,
              CINZA_TEXTO)

    add_text(s, 0.9, y + 1.85, 17.6, 0.5,
              "* Financiamento classificado como parcela Federal por definição adotada nesta apresentação — "
              "a base não distingue dívida do Estado com aval federal de financiamento federal direto.",
              10, CINZA_TEXTO)
    return s


def slide_3_execucao_contratado(prs, m):
    s = novo_slide(prs, "EXECUÇÃO DO CONTRATADO, POR FONTE", 3, TOTAL_PAGINAS)
    add_text(s, 0.9, 1.5, 10, 0.4, "TOTAL JÁ CONTRATADO", 14, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, 1.85, 9, 1.0, abreviar_moeda(m["total_contratado"]), 44, VERDE, negrito=True)
    add_text(s, 10.2, 2.1, 8.8, 0.6,
              f"Detalhamento por fonte disponível para {pct(m['total_contratado']*m['cobertura_detalhamento_pct'], m['total_contratado'])} "
              "do valor contratado.", 13, CINZA_TEXTO)

    y = 3.2
    total_c = m["total_contratado"]
    segmentos = [
        ("APOIADO (OGU) CONTRATADO", m["contratado_ogu"], COR_OGU),
        ("FINANCIAMENTO CONTRATADO", m["contratado_financiamento"], COR_FINANCIAMENTO),
        ("RECURSO ESTADUAL CONTRATADO", m["contratado_estadual"], COR_ESTADUAL),
        ("NÃO DISCRIMINADO", m["contratado_nao_discriminado"], COR_NAO_DISCRIMINADO),
    ]
    for rotulo, valor, c in segmentos:
        barra_horizontal(s, 0.9, y, 17.6, 0.75, rotulo, valor, total_c, c, largura_rotulo=5.2)
        y += 1.0

    add_pendencia_tag(s, 0.9, y + 0.2, w=2.6)
    add_text(s, 3.7, y + 0.2, 14.0, 0.75,
              f"{m['qtd_sem_detalhamento']} ação sem nenhuma categoria de origem preenchida e mais ações com "
              "divergência entre Valor Contratado e o detalhamento — já sinalizadas no Controle de Qualidade "
              "do painel, aguardando correção de cadastro.", 12, VERMELHO)
    return s


def slide_4_a_contratar(prs, m):
    s = novo_slide(prs, "O QUE AINDA PRECISA SER CONTRATADO", 4, TOTAL_PAGINAS)
    add_text(s, 0.9, 1.5, 12, 0.4, "VALOR TOTAL A CONTRATAR", 16, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, 1.85, 12, 1.3, abreviar_moeda(m["total_a_contratar"]), 58, LARANJA, negrito=True)
    add_text(s, 0.9, 3.05, 14, 0.4, pct(m["total_a_contratar"], m["total_investimento"]) +
              " do investimento total da carteira ainda não tem instrumento assinado", 14, CINZA_TEXTO)

    y = 4.0
    add_text(s, 0.9, y, 8, 0.3, "POR FONTE", 13, NAVY, negrito=True)
    y2 = y + 0.5
    for rotulo, valor, c in [
        ("OGU (Apoiado)", m["a_contratar_ogu"], COR_OGU),
        ("Financiamento", m["a_contratar_financiamento"], COR_FINANCIAMENTO),
        ("Recurso Estadual", m["a_contratar_estadual"], COR_ESTADUAL),
    ]:
        barra_horizontal(s, 0.9, y2, 8.6, 0.65, rotulo, valor, m["total_a_contratar"], c, largura_rotulo=2.6)
        y2 += 0.85

    add_text(s, 9.9, y, 8.6, 0.3, "POR EIXO (top 5 em valor a contratar)", 13, NAVY, negrito=True)
    por_eixo_ord = m["por_eixo"].sort_values("A_CONTRATAR", ascending=False).head(5)
    y3 = y + 0.5
    for eixo, linha in por_eixo_ord.iterrows():
        barra_horizontal(s, 9.9, y3, 8.6, 0.7, eixo, float(linha["A_CONTRATAR"]), m["total_a_contratar"],
                          TEAL, largura_rotulo=3.9, tam_rotulo=11, largura_valor=2.2)
        y3 += 0.95
    return s


def slide_5_tabela_resumo_fonte(prs, m):
    s = novo_slide(prs, "TABELA-RESUMO POR FONTE", 5, TOTAL_PAGINAS)
    cabecalho = ["FONTE", "CONTRATADO", "A CONTRATAR", "TOTAL", "% DO TOTAL"]
    total_geral = m["total_investimento"]
    linhas_dados = [
        ("OGU (Apoiado)", m["contratado_ogu"], m["a_contratar_ogu"]),
        ("Financiamento", m["contratado_financiamento"], m["a_contratar_financiamento"]),
        ("Recurso Estadual", m["contratado_estadual"], m["a_contratar_estadual"]),
        ("Não discriminado (contratado)", m["contratado_nao_discriminado"], 0.0),
    ]
    linhas = []
    for nome, c_val, a_val in linhas_dados:
        tot = c_val + a_val
        linhas.append([nome, abreviar_moeda(c_val), abreviar_moeda(a_val), abreviar_moeda(tot),
                        pct(tot, total_geral)])
    linhas.append(["SUBTOTAL FEDERAL (OGU + Financiamento)", abreviar_moeda(m["contratado_federal"]),
                    abreviar_moeda(m["a_contratar_federal"]), abreviar_moeda(m["total_federal"]),
                    pct(m["total_federal"], total_geral)])
    linhas.append(["SUBTOTAL ESTADUAL", abreviar_moeda(m["contratado_estadual"]),
                    abreviar_moeda(m["a_contratar_estadual"]), abreviar_moeda(m["total_estadual"]),
                    pct(m["total_estadual"], total_geral)])
    linhas.append(["TOTAL GERAL", abreviar_moeda(m["total_contratado"]), abreviar_moeda(m["total_a_contratar"]),
                    abreviar_moeda(total_geral), "100,0%"])

    tabela_nativa(s, 0.9, 2.0, 18.2, 5.6, cabecalho, linhas,
                  larguras=[6.2, 3.0, 3.0, 3.0, 3.0],
                  tam_fonte=14, linha_bold_indices={4, 5, 6})
    add_text(s, 0.9, 7.9, 17, 0.4,
              "\"Não discriminado\" = contratado sem detalhamento por fonte preenchido na base (ver slide 3).",
              11, CINZA_TEXTO)
    return s


def slide_6_por_eixo(prs, m):
    s = novo_slide(prs, "DETALHE POR EIXO DO PAC", 6, TOTAL_PAGINAS)
    cabecalho = ["EIXO", "Nº AÇÕES", "INVESTIMENTO", "CONTRATADO", "A CONTRATAR"]
    linhas = []
    for eixo, linha in m["por_eixo"].iterrows():
        linhas.append([eixo, int(linha["QTD"]), abreviar_moeda(linha["INVESTIMENTO"]),
                        abreviar_moeda(linha["CONTRATADO"]), abreviar_moeda(linha["A_CONTRATAR"])])
    tabela_nativa(s, 0.9, 1.9, 18.2, 6.6, cabecalho, linhas,
                  larguras=[6.5, 2.2, 3.17, 3.17, 3.16], tam_fonte=13)
    return s


def slide_7_top10(prs, m):
    s = novo_slide(prs, "MAIORES AÇÕES A CONTRATAR", 7, TOTAL_PAGINAS)
    cabecalho = ["AÇÃO", "MUNICÍPIO", "A CONTRATAR", "STATUS", "FASE"]
    linhas = []
    top10 = m["top10_a_contratar"]
    for _, r in top10.iterrows():
        linhas.append([str(r[pac.col_objeto]), str(r["MUNICIPIO"]), abreviar_moeda(r["_A_CONTRATAR"]),
                        str(r["STATUS_TEXTO"]).title(), str(r["FASE_TEXTO"]).title()])
    tabela_nativa(s, 0.9, 1.7, 18.2, 6.9, cabecalho, linhas,
                  larguras=[5.2, 3.5, 3.0, 3.2, 3.3], tam_fonte=12)

    concluidas_com_saldo = top10[top10["STATUS_TEXTO"].isin(["CONCLUÍDA", "INAUGURADA"])]
    if len(concluidas_com_saldo) > 0:
        add_pendencia_tag(s, 0.9, 8.75, w=2.6)
        add_text(s, 3.7, 8.75, 14.5, 0.5,
                  f"{len(concluidas_com_saldo)} ação(ões) nesta lista aparece(m) como Concluída/Inaugurada mas "
                  "ainda com saldo a contratar na base — verificar cadastro antes de divulgar.",
                  12, VERMELHO)
    return s


def slide_8_pendencias(prs, m):
    s = novo_slide(prs, "PENDÊNCIAS QUE CONDICIONAM A CONTRATAÇÃO", 8, TOTAL_PAGINAS)
    cards = [
        ("CLÁUSULA SUSPENSIVA VIGENTE", m["qtd_clausula_vigente"], abreviar_moeda(m["valor_clausula_vigente"]),
         "ações com condição contratual ainda não resolvida"),
        ("OBRA NÃO INICIADA", m["qtd_obra_nao_iniciada"], None,
          "mais de 40 dias desde a Emissão de O.S., sem avanço registrado"),
        ("PRAZO DE PENDÊNCIA/TAREFA VENCIDO", m["qtd_prazo_pendencia_vencido"], None,
          "tarefas com prazo já ultrapassado"),
        ("SEM CATEGORIA DE ORIGEM DO RECURSO", m["qtd_sem_detalhamento"], None,
          "impede identificar a fonte do valor contratado"),
    ]
    largura_card = 4.1
    espaco = 0.3
    y = 2.0
    for i, (rotulo, qtd, valor, desc) in enumerate(cards):
        x = 0.9 + i * (largura_card + espaco)
        add_rect(s, x, y, largura_card, 3.8, fill=PENDENCIA_FUNDO, line=VERMELHO, arredondado=True, line_w=1.25)
        add_text(s, x + 0.25, y + 0.25, largura_card - 0.5, 0.9, str(qtd), 44, VERMELHO, negrito=True)
        add_text(s, x + 0.25, y + 1.2, largura_card - 0.5, 0.3, "ações", 12, CINZA_TEXTO)
        if valor:
            add_text(s, x + 0.25, y + 1.55, largura_card - 0.5, 0.4, valor, 18, VERMELHO, negrito=True)
        add_text(s, x + 0.25, y + 2.15, largura_card - 0.5, 0.55, rotulo, 12, NAVY, negrito=True)
        add_text(s, x + 0.25, y + 2.75, largura_card - 0.5, 0.95, desc, 10, CINZA_TEXTO)
    return s


def slide_9_sintese(prs, m):
    s = novo_slide(prs, "SÍNTESE E PLEITO", 9, TOTAL_PAGINAS)
    mensagens = [
        (abreviar_moeda(m["total_a_contratar"]), "ainda precisam ser contratados para concluir a carteira do PAC-BA"),
        (abreviar_moeda(m["a_contratar_federal"]), "desse valor dependem de fonte Federal (OGU + Financiamento)"),
        (f"{m['gestao_federal_qtd']}", "ações sob execução direta de órgãos federais na carteira"),
        (abreviar_moeda(m["valor_clausula_vigente"]), f"já contratados têm cláusula suspensiva vigente "
          f"({m['qtd_clausula_vigente']} ações) — recurso condicionado"),
    ]
    y = 1.9
    for valor, texto in mensagens:
        add_rect(s, 0.9, y, 17.6, 1.3, fill=CINZA_CLARO, arredondado=True)
        add_text(s, 1.2, y + 0.12, 5.0, 1.06, valor, 30, VERDE, negrito=True, ancorar=MSO_ANCHOR.MIDDLE)
        add_text(s, 6.4, y + 0.12, 11.8, 1.06, texto, 15, NAVY, ancorar=MSO_ANCHOR.MIDDLE)
        y += 1.5
    return s


# ---------------------------------------------------------------------------
# Planilha de apoio
# ---------------------------------------------------------------------------
def gerar_xlsx_apoio(m, caminho_saida):
    import pandas as pd

    resumo_fonte = pd.DataFrame([
        {"FONTE": "OGU (Apoiado)", "CONTRATADO": m["contratado_ogu"], "A_CONTRATAR": m["a_contratar_ogu"]},
        {"FONTE": "Financiamento", "CONTRATADO": m["contratado_financiamento"],
         "A_CONTRATAR": m["a_contratar_financiamento"]},
        {"FONTE": "Recurso Estadual", "CONTRATADO": m["contratado_estadual"],
         "A_CONTRATAR": m["a_contratar_estadual"]},
        {"FONTE": "Não discriminado (contratado)", "CONTRATADO": m["contratado_nao_discriminado"],
         "A_CONTRATAR": 0.0},
    ])
    resumo_fonte["TOTAL"] = resumo_fonte["CONTRATADO"] + resumo_fonte["A_CONTRATAR"]

    por_eixo = m["por_eixo"].reset_index()
    top10 = m["top10_a_contratar"][[pac.col_objeto, "MUNICIPIO", "_A_CONTRATAR", "STATUS_TEXTO", "FASE_TEXTO"]]

    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        resumo_fonte.to_excel(writer, sheet_name="Resumo por Fonte", index=False)
        por_eixo.to_excel(writer, sheet_name="Por Eixo", index=False)
        top10.to_excel(writer, sheet_name="Top 10 A Contratar", index=False)
        pd.DataFrame([m]).drop(columns=["por_eixo", "top10_a_contratar"]).to_excel(
            writer, sheet_name="Metricas Gerais", index=False
        )


# ---------------------------------------------------------------------------
# API pública — usada tanto pelo botão do painel quanto pelo main() abaixo.
# Exige que configurar(pac) já tenha sido chamado antes.
# ---------------------------------------------------------------------------
def montar_apresentacao(df):
    m = calcular_metricas(df)

    prs = Presentation()
    prs.slide_width = pol(LARGURA_IN)
    prs.slide_height = pol(ALTURA_IN)

    slide_1_panorama(prs, m)
    slide_2_composicao_origem(prs, m)
    slide_3_execucao_contratado(prs, m)
    slide_4_a_contratar(prs, m)
    slide_5_tabela_resumo_fonte(prs, m)
    slide_6_por_eixo(prs, m)
    slide_7_top10(prs, m)
    slide_8_pendencias(prs, m)
    slide_9_sintese(prs, m)

    return prs, m


def gerar_pptx_bytes(df):
    # Usado pelo botão "ANÁLISE FINANCEIRA (PPTX)" do painel (ver
    # _construir_pptx_analise_financeira_bytes em "CGAPE - BALANÇO PAC.py") —
    # monta tudo em memória, sem tocar o disco, igual ao padrão já usado para
    # o PDF (_api_gerar_pdf_bytes).
    import io
    prs, m = montar_apresentacao(df)
    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue(), m


# ---------------------------------------------------------------------------
# main — só usado no modo linha de comando ("python gerar_pptx_pac.py"),
# fora do painel. Faz o próprio bootstrap (recarrega a base do zero, mesma
# técnica do servidor_web.py) porque não há nenhum módulo já carregado pra
# reaproveitar nesse caso.
# ---------------------------------------------------------------------------
def main():
    if pac is None:
        os.environ.setdefault("PAC_WEB_MODE", "1")
        sys.path.insert(0, PASTA_BASE)
        caminho_script_original = os.path.join(PASTA_BASE, "CGAPE - BALANÇO PAC.py")
        spec = importlib.util.spec_from_file_location("cgape_balanco_pac", caminho_script_original)
        modulo = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = modulo
        spec.loader.exec_module(modulo)
        configurar(modulo)

    prs, m = montar_apresentacao(pac.df_original)

    caminho_pptx = os.path.join(PASTA_BASE, "ANALISE_FINANCEIRA_PAC.pptx")
    prs.save(caminho_pptx)
    print("PPTX gerado:", caminho_pptx)

    caminho_xlsx = os.path.join(PASTA_BASE, "APOIO_ANALISE_FINANCEIRA_PAC.xlsx")
    gerar_xlsx_apoio(m, caminho_xlsx)
    print("XLSX de apoio gerado:", caminho_xlsx)


if __name__ == "__main__":
    main()
