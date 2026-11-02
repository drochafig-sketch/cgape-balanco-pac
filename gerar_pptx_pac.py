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
# Variações "A CONTRATAR" (mesma família, mais claras) — ficam com o mesmo
# valor do "cheio" até _sincronizar_cores_com_painel() rodar (ver abaixo).
COR_OGU_CLARO = AZUL
COR_FINANCIAMENTO_CLARO = TEAL
COR_ESTADUAL_CLARO = LARANJA

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
    add_text(slide, x, y, w, h, "⚠ PENDÊNCIA", 13, VERMELHO, negrito=True, alinhar=PP_ALIGN.CENTER,
              ancorar=MSO_ANCHOR.MIDDLE)


def barra_horizontal(slide, x, y, largura_max, altura, rotulo, valor, total, cor_barra, largura_rotulo=4.2,
                      tam_rotulo=16, largura_valor=2.9, tam_valor=16):
    # Tudo (rótulo + barra + valor) precisa caber dentro de largura_max, a
    # partir de x — senão o rótulo de valor estoura a borda do slide.
    # Tamanhos de fonte grandes de propósito: o canvas é 20x11,25 pol (bem
    # maior que o 13,33x7,5 padrão) — um corpo de texto "normal" de
    # apresentação (11-13pt) fica minúsculo nessa escala.
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
              f"{abreviar_moeda(valor)}  ({pct(valor, total)})", tam_valor, NAVY, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)


def barra_horizontal_dupla(slide, x, y, largura_max, altura, rotulo, valor_contratado, valor_a_contratar,
                            total, cor_cheia, cor_clara, largura_rotulo=5.4, tam_rotulo=16,
                            largura_valor=3.6, tam_valor=16):
    # Igual a barra_horizontal, mas o segmento colorido (proporcional ao
    # total desta fonte sobre o total geral) é dividido em dois: CONTRATADO
    # (cor cheia, primeiro) e A CONTRATAR (cor clara, depois) — mesma família
    # de cor, só a tonalidade muda. Pedido do usuário em 02/11/2026.
    total_item = valor_contratado + valor_a_contratar
    fracao_item = (total_item / total) if total else 0
    add_text(slide, x, y, largura_rotulo, altura, rotulo, tam_rotulo, NAVY, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    x_barra = x + largura_rotulo + 0.2
    largura_disponivel = largura_max - largura_rotulo - 0.2 - largura_valor - 0.15
    add_rect(slide, x_barra, y + altura * 0.18, largura_disponivel, altura * 0.64, fill=CINZA_CLARO)
    largura_item = largura_disponivel * fracao_item
    if largura_item > 0:
        fracao_contratado = (valor_contratado / total_item) if total_item else 0
        largura_c = largura_item * fracao_contratado
        largura_a = largura_item - largura_c
        if largura_c > 0:
            add_rect(slide, x_barra, y + altura * 0.18, max(largura_c, 0.03), altura * 0.64, fill=cor_cheia)
        if largura_a > 0:
            add_rect(slide, x_barra + largura_c, y + altura * 0.18, max(largura_a, 0.03), altura * 0.64,
                      fill=cor_clara)
    add_text(slide, x_barra + largura_disponivel + 0.15, y, largura_valor, altura,
              f"{abreviar_moeda(total_item)}  ({pct(total_item, total)})", tam_valor, NAVY, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    add_text(slide, x_barra, y + altura + 0.04, largura_disponivel + largura_valor + 0.15, 0.3,
              f"Contratado: {abreviar_moeda(valor_contratado)} ({pct(valor_contratado, total_item)})   ·   "
              f"A contratar: {abreviar_moeda(valor_a_contratar)} ({pct(valor_a_contratar, total_item)})",
              12, CINZA_TEXTO)


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
# Cores — decisão do usuário em 02/11/2026: a apresentação deixa de ter
# identidade própria para OGU/Financiamento/Estadual e passa a usar
# exatamente os mesmos tons do painel/PDF (pac.TONS_FINANCEIRO), pra não
# existir uma "cor da apresentação" diferente da "cor do painel" pra Apoiado,
# Financiamento e Recurso Estadual. Só pode rodar depois de configurar(pac).
# ---------------------------------------------------------------------------
def _sincronizar_cores_com_painel():
    global COR_OGU, COR_FINANCIAMENTO, COR_ESTADUAL, COR_NAO_DISCRIMINADO
    global COR_OGU_CLARO, COR_FINANCIAMENTO_CLARO, COR_ESTADUAL_CLARO
    tons = pac.TONS_FINANCEIRO
    COR_OGU = tons["apoiado"]["contratado"].lstrip("#")
    COR_OGU_CLARO = tons["apoiado"]["aContratar"].lstrip("#")
    COR_FINANCIAMENTO = tons["financiamento"]["contratado"].lstrip("#")
    COR_FINANCIAMENTO_CLARO = tons["financiamento"]["aContratar"].lstrip("#")
    COR_ESTADUAL = tons["estadual"]["contratado"].lstrip("#")
    COR_ESTADUAL_CLARO = tons["estadual"]["aContratar"].lstrip("#")
    COR_NAO_DISCRIMINADO = tons["naoDiscriminado"].lstrip("#")


# ---------------------------------------------------------------------------
# Métricas (tudo recalculado a partir de df — nada hardcoded)
# ---------------------------------------------------------------------------
def calcular_metricas(df):
    _sincronizar_cores_com_painel()
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
    add_text(s, 0.9, y_barra - 0.45, 6, 0.35, "CONTRATADO  x  A CONTRATAR", 17, NAVY, negrito=True)
    largura_barra = 17.0
    x0 = 0.9
    frac_contratado = m["total_contratado"] / m["total_investimento"] if m["total_investimento"] else 0
    add_rect(s, x0, y_barra, largura_barra * frac_contratado, 0.9, fill=VERDE)
    add_rect(s, x0 + largura_barra * frac_contratado, y_barra, largura_barra * (1 - frac_contratado), 0.9,
              fill=LARANJA)
    add_text(s, x0 + 0.15, y_barra + 0.1, 6, 0.35, f"CONTRATADO · {abreviar_moeda(m['total_contratado'])}", 17,
              BRANCO, negrito=True)
    add_text(s, x0 + 0.15, y_barra + 0.5, 6, 0.3, pct(m["total_contratado"], m["total_investimento"]), 15,
              BRANCO)
    x_txt2 = x0 + largura_barra * frac_contratado + 0.15
    add_text(s, x_txt2, y_barra + 0.1, 7, 0.35, f"A CONTRATAR · {abreviar_moeda(m['total_a_contratar'])}", 17,
              BRANCO, negrito=True)
    add_text(s, x_txt2, y_barra + 0.5, 7, 0.3, pct(m["total_a_contratar"], m["total_investimento"]), 15,
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
        add_text(s, x + 0.25, y_cards + 1.3, largura_card - 0.5, 0.6, rotulo, 14, CINZA_TEXTO, negrito=True)
    return s


def slide_2_composicao_origem(prs, m):
    s = novo_slide(prs, "COMPOSIÇÃO POR ORIGEM DO RECURSO", 2, TOTAL_PAGINAS)
    total = m["total_investimento"]
    y = 1.9
    add_text(s, 0.9, y, 14, 0.4, "INVESTIMENTO TOTAL, POR FONTE (contratado + a contratar)", 17, CINZA_TEXTO,
              negrito=True)
    add_text(s, 13.6, y, 4.9, 0.4, "tom forte = CONTRATADO · tom claro = A CONTRATAR", 12, CINZA_TEXTO,
              alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
    y += 0.65
    fontes = [
        ("OGU (Apoiado)", m["contratado_ogu"], m["a_contratar_ogu"], COR_OGU, COR_OGU_CLARO),
        ("FINANCIAMENTO", m["contratado_financiamento"], m["a_contratar_financiamento"], COR_FINANCIAMENTO,
         COR_FINANCIAMENTO_CLARO),
        ("RECURSO ESTADUAL (Contrapartida + Complementar)", m["contratado_estadual"], m["a_contratar_estadual"],
         COR_ESTADUAL, COR_ESTADUAL_CLARO),
    ]
    for rotulo, c_val, a_val, cheia, clara in fontes:
        barra_horizontal_dupla(s, 0.9, y, 17.6, 0.75, rotulo, c_val, a_val, total, cheia, clara,
                                largura_rotulo=5.4)
        y += 1.4

    # destaque federal x estadual
    y += 0.3
    add_rect(s, 0.9, y, 8.4, 1.6, fill=CINZA_CLARO, arredondado=True)
    add_text(s, 1.2, y + 0.16, 7.8, 0.35, "PARCELA FEDERAL (OGU + Financiamento)*", 14, CINZA_TEXTO, negrito=True)
    add_text(s, 1.2, y + 0.52, 7.8, 0.8, abreviar_moeda(m["total_federal"]), 32, AZUL, negrito=True)
    add_text(s, 1.2, y + 1.18, 7.8, 0.35, pct(m["total_federal"], total) + " do investimento total", 14,
              CINZA_TEXTO)

    add_rect(s, 9.6, y, 8.4, 1.6, fill=CINZA_CLARO, arredondado=True)
    add_text(s, 9.9, y + 0.16, 7.8, 0.35, "PARCELA ESTADUAL (Contrapartida + Complementar)", 14, CINZA_TEXTO,
              negrito=True)
    add_text(s, 9.9, y + 0.52, 7.8, 0.8, abreviar_moeda(m["total_estadual"]), 32, LARANJA, negrito=True)
    add_text(s, 9.9, y + 1.18, 7.8, 0.35, pct(m["total_estadual"], total) + " do investimento total", 14,
              CINZA_TEXTO)

    add_text(s, 0.9, y + 1.85, 17.6, 0.5,
              "* Financiamento classificado como parcela Federal por definição adotada nesta apresentação — "
              "a base não distingue dívida do Estado com aval federal de financiamento federal direto.",
              12, CINZA_TEXTO)
    return s


def slide_3_execucao_contratado(prs, m):
    s = novo_slide(prs, "EXECUÇÃO DO CONTRATADO, POR FONTE", 3, TOTAL_PAGINAS)
    add_text(s, 0.9, 1.5, 10, 0.4, "TOTAL JÁ CONTRATADO", 17, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, 1.85, 9, 1.0, abreviar_moeda(m["total_contratado"]), 44, VERDE, negrito=True)
    add_text(s, 10.2, 2.15, 8.8, 0.7,
              f"Detalhamento por fonte disponível para {pct(m['total_contratado']*m['cobertura_detalhamento_pct'], m['total_contratado'])} "
              "do valor contratado.", 15, CINZA_TEXTO)

    y = 3.3
    total_c = m["total_contratado"]
    segmentos = [
        ("APOIADO (OGU) CONTRATADO", m["contratado_ogu"], COR_OGU),
        ("FINANCIAMENTO CONTRATADO", m["contratado_financiamento"], COR_FINANCIAMENTO),
        ("RECURSO ESTADUAL CONTRATADO", m["contratado_estadual"], COR_ESTADUAL),
        ("NÃO DISCRIMINADO", m["contratado_nao_discriminado"], COR_NAO_DISCRIMINADO),
    ]
    for rotulo, valor, c in segmentos:
        barra_horizontal(s, 0.9, y, 17.6, 0.75, rotulo, valor, total_c, c, largura_rotulo=5.4)
        y += 1.05

    add_pendencia_tag(s, 0.9, y + 0.2, w=2.6)
    add_text(s, 3.7, y + 0.2, 14.0, 0.85,
              f"{m['qtd_sem_detalhamento']} ação sem nenhuma categoria de origem preenchida e mais ações com "
              "divergência entre Valor Contratado e o detalhamento — já sinalizadas no Controle de Qualidade "
              "do painel, aguardando correção de cadastro.", 14, VERMELHO)
    return s


def slide_4_a_contratar(prs, m):
    s = novo_slide(prs, "O QUE AINDA PRECISA SER CONTRATADO", 4, TOTAL_PAGINAS)
    add_text(s, 0.9, 1.5, 12, 0.4, "VALOR TOTAL A CONTRATAR", 17, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, 1.85, 12, 1.3, abreviar_moeda(m["total_a_contratar"]), 58, LARANJA, negrito=True)
    add_text(s, 0.9, 3.1, 17.6, 0.4, pct(m["total_a_contratar"], m["total_investimento"]) +
              " do investimento total da carteira ainda não tem instrumento assinado", 15, CINZA_TEXTO)

    y = 4.1
    add_text(s, 0.9, y, 8, 0.35, "POR FONTE", 17, NAVY, negrito=True)
    y2 = y + 0.6
    for rotulo, valor, c in [
        ("OGU (Apoiado)", m["a_contratar_ogu"], COR_OGU_CLARO),
        ("Financiamento", m["a_contratar_financiamento"], COR_FINANCIAMENTO_CLARO),
        ("Recurso Estadual", m["a_contratar_estadual"], COR_ESTADUAL_CLARO),
    ]:
        barra_horizontal(s, 0.9, y2, 8.6, 0.7, rotulo, valor, m["total_a_contratar"], c, largura_rotulo=2.9,
                          tam_rotulo=14, largura_valor=2.5, tam_valor=14)
        y2 += 0.95

    add_text(s, 9.9, y, 8.6, 0.35, "POR EIXO (top 5 em valor a contratar)", 17, NAVY, negrito=True)
    por_eixo_ord = m["por_eixo"].sort_values("A_CONTRATAR", ascending=False).head(5)
    y3 = y + 0.6
    for eixo, linha in por_eixo_ord.iterrows():
        barra_horizontal(s, 9.9, y3, 8.6, 0.75, eixo, float(linha["A_CONTRATAR"]), m["total_a_contratar"],
                          TEAL, largura_rotulo=3.7, tam_rotulo=13, largura_valor=2.3, tam_valor=13)
        y3 += 1.0
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
                  tam_fonte=16, linha_bold_indices={4, 5, 6})
    add_text(s, 0.9, 7.9, 17, 0.4,
              "\"Não discriminado\" = contratado sem detalhamento por fonte preenchido na base (ver slide 3).",
              13, CINZA_TEXTO)
    return s


def slide_6_por_eixo(prs, m):
    s = novo_slide(prs, "DETALHE POR EIXO DO PAC", 6, TOTAL_PAGINAS)
    cabecalho = ["EIXO", "Nº AÇÕES", "INVESTIMENTO", "CONTRATADO", "A CONTRATAR"]
    linhas = []
    for eixo, linha in m["por_eixo"].iterrows():
        linhas.append([eixo, int(linha["QTD"]), abreviar_moeda(linha["INVESTIMENTO"]),
                        abreviar_moeda(linha["CONTRATADO"]), abreviar_moeda(linha["A_CONTRATAR"])])
    tabela_nativa(s, 0.9, 1.9, 18.2, 6.6, cabecalho, linhas,
                  larguras=[6.5, 2.2, 3.17, 3.17, 3.16], tam_fonte=15)
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
                  larguras=[5.2, 3.5, 3.0, 3.2, 3.3], tam_fonte=14)

    concluidas_com_saldo = top10[top10["STATUS_TEXTO"].isin(["CONCLUÍDA", "INAUGURADA"])]
    if len(concluidas_com_saldo) > 0:
        add_pendencia_tag(s, 0.9, 8.75, w=2.6)
        add_text(s, 3.7, 8.75, 14.5, 0.5,
                  f"{len(concluidas_com_saldo)} ação(ões) nesta lista aparece(m) como Concluída/Inaugurada mas "
                  "ainda com saldo a contratar na base — verificar cadastro antes de divulgar.",
                  14, VERMELHO)
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
        add_text(s, x + 0.25, y + 1.2, largura_card - 0.5, 0.3, "ações", 14, CINZA_TEXTO)
        if valor:
            add_text(s, x + 0.25, y + 1.55, largura_card - 0.5, 0.4, valor, 20, VERMELHO, negrito=True)
        add_text(s, x + 0.25, y + 2.15, largura_card - 0.5, 0.55, rotulo, 14, NAVY, negrito=True)
        add_text(s, x + 0.25, y + 2.75, largura_card - 0.5, 0.95, desc, 12, CINZA_TEXTO)
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
        add_text(s, 1.2, y + 0.12, 5.0, 1.06, valor, 32, VERDE, negrito=True, ancorar=MSO_ANCHOR.MIDDLE)
        add_text(s, 6.4, y + 0.12, 11.8, 1.06, texto, 17, NAVY, ancorar=MSO_ANCHOR.MIDDLE)
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
# Versão HTML da mesma apresentação — mesmos 9 "slides", mesmas métricas
# (reaproveita calcular_metricas), mesma paleta de cores. Existe por causa de
# uma limitação real do PPTX: texto em caixa de posição/tamanho FIXOS (como
# o PowerPoint exige) tem que ter o tamanho de fonte calculado à mão pra não
# estourar — o HTML não tem esse problema (flexbox quebra linha sozinho,
# a barra usa % de largura em vez de polegada calculada), então serve tanto
# como conferência visual rápida (abre em qualquer navegador, sem precisar
# de PowerPoint/Canva) quanto como versão alternativa pra quem só precisa
# mostrar os números na tela, sem editar/enviar um arquivo de slides.
# ---------------------------------------------------------------------------
def _html_esc(texto):
    import html as _html_mod
    return _html_mod.escape(str(texto))


def _html_barra(rotulo, valor, total, cor_barra, tam_rotulo_px=280):
    fracao = (valor / total * 100) if total else 0
    return f'''<div class="barra-linha">
      <div class="barra-rotulo" style="width:{tam_rotulo_px}px">{_html_esc(rotulo)}</div>
      <div class="barra-trilho"><div class="barra-fill" style="width:{fracao:.2f}%;background:#{cor_barra}"></div></div>
      <div class="barra-valor">{_html_esc(abreviar_moeda(valor))}&nbsp;&nbsp;({_html_esc(pct(valor, total))})</div>
    </div>'''


def _html_barra_dupla(rotulo, valor_contratado, valor_a_contratar, total, cor_cheia, cor_clara,
                       tam_rotulo_px=340):
    # Igual a _html_barra, mas o trecho colorido vira dois tons empilhados
    # (CONTRATADO cheio + A CONTRATAR claro) — mesma ideia da barra dupla do
    # PPTX (barra_horizontal_dupla). Pedido do usuário em 02/11/2026.
    total_item = valor_contratado + valor_a_contratar
    fracao_item = (total_item / total * 100) if total else 0
    frac_c = (valor_contratado / total_item * 100) if total_item else 0
    frac_a = 100 - frac_c if total_item else 0
    return f'''<div class="barra-linha">
      <div class="barra-rotulo" style="width:{tam_rotulo_px}px">{_html_esc(rotulo)}</div>
      <div class="barra-trilho">
        <div class="barra-fill-grupo" style="width:{fracao_item:.4f}%">
          <div class="barra-fill" style="width:{frac_c:.2f}%;background:#{cor_cheia}" title="Contratado"></div>
          <div class="barra-fill" style="width:{frac_a:.2f}%;background:#{cor_clara}" title="A contratar"></div>
        </div>
      </div>
      <div class="barra-valor">{_html_esc(abreviar_moeda(total_item))}&nbsp;&nbsp;({_html_esc(pct(total_item, total))})</div>
    </div>
    <div class="barra-detalhe">Contratado {_html_esc(abreviar_moeda(valor_contratado))}
      ({_html_esc(pct(valor_contratado, total_item))}) &nbsp;·&nbsp;
      A contratar {_html_esc(abreviar_moeda(valor_a_contratar))}
      ({_html_esc(pct(valor_a_contratar, total_item))})</div>'''


def _html_tabela(cabecalho, linhas, linhas_bold=()):
    ths = "".join(f"<th>{_html_esc(c)}</th>" for c in cabecalho)
    trs = []
    for i, linha in enumerate(linhas):
        classe = " class=\"linha-bold\"" if i in linhas_bold else ""
        tds = "".join(f"<td>{_html_esc(v)}</td>" for v in linha)
        trs.append(f"<tr{classe}>{tds}</tr>")
    return f'<table class="pac-tabela"><thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table>'


def _html_cabecalho_rodape(tema, pagina, total_paginas, conteudo_html):
    nome_planilha = os.path.basename(pac.CAMINHO_PLANILHA) if hasattr(pac, "CAMINHO_PLANILHA") else \
        "PANORAMA - PAC ORIGINAL - PAC SELEÇÕES - 2026"
    posicao = pac.ultima_atualizacao.strftime("%m/%Y") if hasattr(pac, "ultima_atualizacao") else date.today().strftime("%m/%Y")
    return f'''<section class="slide">
      <div class="cabecalho">
        <div class="titulo">PAC - {_html_esc(tema)}</div>
        <div class="subtitulo">COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE PROJETOS ESTRUTURANTES</div>
        <div class="governo"><div>GOVERNO DO ESTADO DA BAHIA</div><div>CASA CIVIL</div></div>
      </div>
      <div class="conteudo">{conteudo_html}</div>
      <div class="rodape">
        <span>Fonte: {_html_esc(nome_planilha)} · posição em {_html_esc(posicao)}</span>
        <span>{pagina}/{total_paginas}</span>
      </div>
    </section>'''


def montar_html_apresentacao(df):
    m = calcular_metricas(df)
    total_paginas = 9
    total_investimento = m["total_investimento"]

    # --- slide 1 ---
    cards_status = [
        ("CONCLUÍDO / INAUGURADO", m["qtd_concluido"], "2E7D32"),
        ("EM OBRAS / ANDAMENTO", m["qtd_andamento"], "0B5CAD"),
        ("SELEÇÃO / PREPARATÓRIA / LICITAÇÃO", m["qtd_preparatoria"], "E68A00"),
        ("PARALISADO", m["qtd_paralisado"], "C62828"),
    ]
    frac_contratado = (m["total_contratado"] / total_investimento * 100) if total_investimento else 0
    slide1 = _html_cabecalho_rodape("PANORAMA DA CARTEIRA", 1, total_paginas, f'''
      <div class="rotulo-secao">INVESTIMENTO TOTAL DA CARTEIRA</div>
      <div class="valor-principal" style="color:#2E7D32">{_html_esc(abreviar_moeda(total_investimento))}</div>
      <div class="texto-secundario">{m["qtd_acoes"]} ações do PAC acompanhadas pelo Estado da Bahia</div>
      <div class="rotulo-secao" style="margin-top:28px">CONTRATADO&nbsp;&nbsp;x&nbsp;&nbsp;A CONTRATAR</div>
      <div class="barra-empilhada">
        <div class="barra-empilhada-seg" style="width:{frac_contratado:.2f}%;background:#2E7D32">
          <div class="be-titulo">CONTRATADO · {_html_esc(abreviar_moeda(m["total_contratado"]))}</div>
          <div class="be-sub">{_html_esc(pct(m["total_contratado"], total_investimento))}</div>
        </div>
        <div class="barra-empilhada-seg" style="width:{100 - frac_contratado:.2f}%;background:#E68A00">
          <div class="be-titulo">A CONTRATAR · {_html_esc(abreviar_moeda(m["total_a_contratar"]))}</div>
          <div class="be-sub">{_html_esc(pct(m["total_a_contratar"], total_investimento))}</div>
        </div>
      </div>
      <div class="cartoes-grade">
        {"".join(f'<div class="cartao-status"><div class="cs-barra" style="background:#{c}"></div><div class="cs-num" style="color:#{c}">{qtd}</div><div class="cs-rotulo">{_html_esc(rot)}</div></div>' for rot, qtd, c in cards_status)}
      </div>
    ''')

    # --- slide 2 ---
    slide2 = _html_cabecalho_rodape("COMPOSIÇÃO POR ORIGEM DO RECURSO", 2, total_paginas, f'''
      <div class="rotulo-secao" style="display:flex;justify-content:space-between;align-items:baseline">
        <span>INVESTIMENTO TOTAL, POR FONTE (contratado + a contratar)</span>
        <span style="font-size:13px;font-weight:400;color:#5B6576">tom forte = CONTRATADO · tom claro = A CONTRATAR</span>
      </div>
      <div class="bloco-barras">
        {_html_barra_dupla("OGU (Apoiado)", m["contratado_ogu"], m["a_contratar_ogu"], total_investimento, COR_OGU, COR_OGU_CLARO, 340)}
        {_html_barra_dupla("FINANCIAMENTO", m["contratado_financiamento"], m["a_contratar_financiamento"], total_investimento, COR_FINANCIAMENTO, COR_FINANCIAMENTO_CLARO, 340)}
        {_html_barra_dupla("RECURSO ESTADUAL (Contrapartida + Complementar)", m["contratado_estadual"], m["a_contratar_estadual"], total_investimento, COR_ESTADUAL, COR_ESTADUAL_CLARO, 340)}
      </div>
      <div class="duas-colunas" style="margin-top:24px">
        <div class="cartao-destaque">
          <div class="cd-rotulo">PARCELA FEDERAL (OGU + Financiamento)*</div>
          <div class="cd-valor" style="color:#0B5CAD">{_html_esc(abreviar_moeda(m["total_federal"]))}</div>
          <div class="cd-rodape">{_html_esc(pct(m["total_federal"], total_investimento))} do investimento total</div>
        </div>
        <div class="cartao-destaque">
          <div class="cd-rotulo">PARCELA ESTADUAL (Contrapartida + Complementar)</div>
          <div class="cd-valor" style="color:#E68A00">{_html_esc(abreviar_moeda(m["total_estadual"]))}</div>
          <div class="cd-rodape">{_html_esc(pct(m["total_estadual"], total_investimento))} do investimento total</div>
        </div>
      </div>
      <div class="nota-rodape">* Financiamento classificado como parcela Federal por definição adotada nesta
      apresentação — a base não distingue dívida do Estado com aval federal de financiamento federal direto.</div>
    ''')

    # --- slide 3 ---
    total_contratado = m["total_contratado"]
    slide3 = _html_cabecalho_rodape("EXECUÇÃO DO CONTRATADO, POR FONTE", 3, total_paginas, f'''
      <div class="duas-colunas-desiguais">
        <div>
          <div class="rotulo-secao">TOTAL JÁ CONTRATADO</div>
          <div class="valor-principal" style="color:#2E7D32">{_html_esc(abreviar_moeda(total_contratado))}</div>
        </div>
        <div class="texto-secundario" style="align-self:center">Detalhamento por fonte disponível para
        {_html_esc(pct(total_contratado * m["cobertura_detalhamento_pct"], total_contratado))} do valor contratado.</div>
      </div>
      <div class="bloco-barras" style="margin-top:20px">
        {_html_barra("APOIADO (OGU) CONTRATADO", m["contratado_ogu"], total_contratado, COR_OGU, 340)}
        {_html_barra("FINANCIAMENTO CONTRATADO", m["contratado_financiamento"], total_contratado, COR_FINANCIAMENTO, 340)}
        {_html_barra("RECURSO ESTADUAL CONTRATADO", m["contratado_estadual"], total_contratado, COR_ESTADUAL, 340)}
        {_html_barra("NÃO DISCRIMINADO", m["contratado_nao_discriminado"], total_contratado, COR_NAO_DISCRIMINADO, 340)}
      </div>
      <div class="caixa-pendencia">⚠ PENDÊNCIA — {m['qtd_sem_detalhamento']} ação sem nenhuma categoria de origem
      preenchida e mais ações com divergência entre Valor Contratado e o detalhamento — já sinalizadas no
      Controle de Qualidade do painel, aguardando correção de cadastro.</div>
    ''')

    # --- slide 4 ---
    por_eixo_ord = m["por_eixo"].sort_values("A_CONTRATAR", ascending=False).head(5)
    slide4 = _html_cabecalho_rodape("O QUE AINDA PRECISA SER CONTRATADO", 4, total_paginas, f'''
      <div class="rotulo-secao">VALOR TOTAL A CONTRATAR</div>
      <div class="valor-principal" style="color:#E68A00">{_html_esc(abreviar_moeda(m["total_a_contratar"]))}</div>
      <div class="texto-secundario">{_html_esc(pct(m["total_a_contratar"], total_investimento))} do investimento
      total da carteira ainda não tem instrumento assinado</div>
      <div class="duas-colunas" style="margin-top:24px; align-items:start">
        <div>
          <div class="rotulo-secao" style="margin-bottom:14px">POR FONTE</div>
          <div class="bloco-barras">
            {_html_barra("OGU (Apoiado)", m["a_contratar_ogu"], m["total_a_contratar"], COR_OGU_CLARO, 170)}
            {_html_barra("Financiamento", m["a_contratar_financiamento"], m["total_a_contratar"], COR_FINANCIAMENTO_CLARO, 170)}
            {_html_barra("Recurso Estadual", m["a_contratar_estadual"], m["total_a_contratar"], COR_ESTADUAL_CLARO, 170)}
          </div>
        </div>
        <div>
          <div class="rotulo-secao" style="margin-bottom:14px">POR EIXO (top 5 em valor a contratar)</div>
          <div class="bloco-barras">
            {"".join(_html_barra(eixo, float(linha["A_CONTRATAR"]), m["total_a_contratar"], "2F7F92", 230) for eixo, linha in por_eixo_ord.iterrows())}
          </div>
        </div>
      </div>
    ''')

    # --- slide 5 ---
    linhas5 = []
    for nome, c_val, a_val in [
        ("OGU (Apoiado)", m["contratado_ogu"], m["a_contratar_ogu"]),
        ("Financiamento", m["contratado_financiamento"], m["a_contratar_financiamento"]),
        ("Recurso Estadual", m["contratado_estadual"], m["a_contratar_estadual"]),
        ("Não discriminado (contratado)", m["contratado_nao_discriminado"], 0.0),
    ]:
        tot = c_val + a_val
        linhas5.append([nome, abreviar_moeda(c_val), abreviar_moeda(a_val), abreviar_moeda(tot),
                         pct(tot, total_investimento)])
    linhas5.append(["SUBTOTAL FEDERAL (OGU + Financiamento)", abreviar_moeda(m["contratado_federal"]),
                     abreviar_moeda(m["a_contratar_federal"]), abreviar_moeda(m["total_federal"]),
                     pct(m["total_federal"], total_investimento)])
    linhas5.append(["SUBTOTAL ESTADUAL", abreviar_moeda(m["contratado_estadual"]),
                     abreviar_moeda(m["a_contratar_estadual"]), abreviar_moeda(m["total_estadual"]),
                     pct(m["total_estadual"], total_investimento)])
    linhas5.append(["TOTAL GERAL", abreviar_moeda(m["total_contratado"]), abreviar_moeda(m["total_a_contratar"]),
                     abreviar_moeda(total_investimento), "100,0%"])
    slide5 = _html_cabecalho_rodape("TABELA-RESUMO POR FONTE", 5, total_paginas, f'''
      {_html_tabela(["FONTE", "CONTRATADO", "A CONTRATAR", "TOTAL", "% DO TOTAL"], linhas5, linhas_bold={4, 5, 6})}
      <div class="nota-rodape">"Não discriminado" = contratado sem detalhamento por fonte preenchido na base
      (ver slide 3).</div>
    ''')

    # --- slide 6 ---
    linhas6 = []
    for eixo, linha in m["por_eixo"].iterrows():
        linhas6.append([eixo, int(linha["QTD"]), abreviar_moeda(linha["INVESTIMENTO"]),
                         abreviar_moeda(linha["CONTRATADO"]), abreviar_moeda(linha["A_CONTRATAR"])])
    slide6 = _html_cabecalho_rodape("DETALHE POR EIXO DO PAC", 6, total_paginas,
        _html_tabela(["EIXO", "Nº AÇÕES", "INVESTIMENTO", "CONTRATADO", "A CONTRATAR"], linhas6))

    # --- slide 7 ---
    top10 = m["top10_a_contratar"]
    linhas7 = []
    for _, r in top10.iterrows():
        linhas7.append([str(r[pac.col_objeto]), str(r["MUNICIPIO"]), abreviar_moeda(r["_A_CONTRATAR"]),
                         str(r["STATUS_TEXTO"]).title(), str(r["FASE_TEXTO"]).title()])
    concluidas_com_saldo = top10[top10["STATUS_TEXTO"].isin(["CONCLUÍDA", "INAUGURADA"])]
    aviso7 = ""
    if len(concluidas_com_saldo) > 0:
        aviso7 = (f'<div class="caixa-pendencia">⚠ PENDÊNCIA — {len(concluidas_com_saldo)} ação(ões) nesta '
                  "lista aparece(m) como Concluída/Inaugurada mas ainda com saldo a contratar na base — "
                  "verificar cadastro antes de divulgar.</div>")
    slide7 = _html_cabecalho_rodape("MAIORES AÇÕES A CONTRATAR", 7, total_paginas,
        _html_tabela(["AÇÃO", "MUNICÍPIO", "A CONTRATAR", "STATUS", "FASE"], linhas7) + aviso7)

    # --- slide 8 ---
    cards8 = [
        ("CLÁUSULA SUSPENSIVA VIGENTE", m["qtd_clausula_vigente"], abreviar_moeda(m["valor_clausula_vigente"]),
         "ações com condição contratual ainda não resolvida"),
        ("OBRA NÃO INICIADA", m["qtd_obra_nao_iniciada"], None,
         "mais de 40 dias desde a Emissão de O.S., sem avanço registrado"),
        ("PRAZO DE PENDÊNCIA/TAREFA VENCIDO", m["qtd_prazo_pendencia_vencido"], None,
         "tarefas com prazo já ultrapassado"),
        ("SEM CATEGORIA DE ORIGEM DO RECURSO", m["qtd_sem_detalhamento"], None,
         "impede identificar a fonte do valor contratado"),
    ]
    slide8 = _html_cabecalho_rodape("PENDÊNCIAS QUE CONDICIONAM A CONTRATAÇÃO", 8, total_paginas, f'''
      <div class="cartoes-pendencia">
        {"".join(f"""<div class="cartao-pendencia">
          <div class="cp-num">{qtd}</div>
          <div class="cp-acoes">ações</div>
          {f'<div class="cp-valor">{_html_esc(valor)}</div>' if valor else ''}
          <div class="cp-rotulo">{_html_esc(rotulo)}</div>
          <div class="cp-desc">{_html_esc(desc)}</div>
        </div>""" for rotulo, qtd, valor, desc in cards8)}
      </div>
    ''')

    # --- slide 9 ---
    mensagens9 = [
        (abreviar_moeda(m["total_a_contratar"]), "ainda precisam ser contratados para concluir a carteira do PAC-BA"),
        (abreviar_moeda(m["a_contratar_federal"]), "desse valor dependem de fonte Federal (OGU + Financiamento)"),
        (f"{m['gestao_federal_qtd']}", "ações sob execução direta de órgãos federais na carteira"),
        (abreviar_moeda(m["valor_clausula_vigente"]),
         f"já contratados têm cláusula suspensiva vigente ({m['qtd_clausula_vigente']} ações) — recurso condicionado"),
    ]
    slide9 = _html_cabecalho_rodape("SÍNTESE E PLEITO", 9, total_paginas, f'''
      <div class="lista-sintese">
        {"".join(f'''<div class="linha-sintese">
          <div class="ls-valor">{_html_esc(valor)}</div>
          <div class="ls-texto">{_html_esc(texto)}</div>
        </div>''' for valor, texto in mensagens9)}
      </div>
    ''')

    css = '''
    * { box-sizing: border-box; }
    body { margin:0; padding:40px 16px; background:#e9ebef; font-family:'Roboto','Segoe UI',Arial,sans-serif; }
    .slide { width:1480px; max-width:100%; aspect-ratio:20/11.25; margin:0 auto 40px; background:#fff;
      box-shadow:0 4px 24px rgba(0,0,0,.2); position:relative; overflow:hidden; border-radius:6px; }
    .cabecalho { background:#0B1F3F; color:#fff; padding:22px 40px 14px; position:relative; }
    .cabecalho .titulo { font-size:30px; font-weight:700; }
    .cabecalho .subtitulo { font-size:14px; font-weight:700; color:#C9D6EA; margin-top:8px; }
    .cabecalho .governo { position:absolute; top:22px; right:40px; text-align:right; font-weight:700;
      font-size:17px; color:#4FA3E0; line-height:1.7; }
    .conteudo { padding: 28px 40px; }
    .rodape { position:absolute; bottom:14px; left:40px; right:40px; display:flex; justify-content:space-between;
      font-size:12px; color:#5B6576; }
    .rotulo-secao { font-size:18px; font-weight:700; color:#5B6576; text-transform:uppercase; }
    .valor-principal { font-size:68px; font-weight:900; margin-top:6px; }
    .texto-secundario { font-size:17px; color:#5B6576; margin-top:6px; }
    .barra-empilhada { display:flex; height:90px; border-radius:4px; overflow:hidden; margin-top:14px; }
    .barra-empilhada-seg { padding:14px 18px; color:#fff; }
    .be-titulo { font-size:19px; font-weight:700; }
    .be-sub { font-size:16px; margin-top:6px; }
    .cartoes-grade { display:flex; gap:20px; margin-top:28px; }
    .cartao-status { flex:1; background:#fff; border:1px solid #D8DCE3; border-radius:10px; overflow:hidden; }
    .cs-barra { height:8px; }
    .cs-num { font-size:48px; font-weight:900; padding:14px 18px 0; }
    .cs-rotulo { font-size:15px; font-weight:700; color:#5B6576; padding:10px 18px 18px; }
    .bloco-barras { margin-top:10px; }
    .barra-linha { display:flex; align-items:center; gap:18px; margin-bottom:16px; }
    .barra-rotulo { font-weight:700; color:#0B1F3F; font-size:17px; flex-shrink:0; }
    .barra-trilho { flex:1; height:30px; background:#F2F4F7; border-radius:4px; overflow:hidden; display:flex; }
    .barra-fill { height:100%; }
    .barra-fill-grupo { height:100%; display:flex; }
    .barra-valor { width:220px; text-align:right; font-weight:700; color:#0B1F3F; font-size:17px; flex-shrink:0; }
    .barra-detalhe { margin:-8px 0 16px 358px; font-size:13px; color:#5B6576; }
    .duas-colunas { display:grid; grid-template-columns:1fr 1fr; gap:24px; }
    .duas-colunas-desiguais { display:grid; grid-template-columns:1fr 1fr; gap:24px; }
    .cartao-destaque { background:#F2F4F7; border-radius:10px; padding:20px 24px; }
    .cd-rotulo { font-size:15px; font-weight:700; color:#5B6576; }
    .cd-valor { font-size:38px; font-weight:900; margin-top:8px; }
    .cd-rodape { font-size:15px; color:#5B6576; margin-top:8px; }
    .nota-rodape { font-size:13px; color:#5B6576; margin-top:18px; }
    .caixa-pendencia { background:#FDECEC; border:1px solid #C62828; border-radius:10px; padding:14px 20px;
      color:#C62828; font-weight:700; font-size:15px; margin-top:18px; }
    table.pac-tabela { width:100%; border-collapse:collapse; font-size:17px; margin-top:6px; }
    table.pac-tabela th { background:#0B1F3F; color:#fff; text-align:right; padding:12px 16px; font-size:16px; }
    table.pac-tabela th:first-child, table.pac-tabela td:first-child { text-align:left; }
    table.pac-tabela td { padding:10px 16px; text-align:right; border-bottom:1px solid #D8DCE3; }
    table.pac-tabela tr:nth-child(even) td { background:#F2F4F7; }
    table.pac-tabela tr.linha-bold td { font-weight:700; border-top:2px solid #0B1F3F; }
    .cartoes-pendencia { display:flex; gap:20px; margin-top:10px; }
    .cartao-pendencia { flex:1; background:#FDECEC; border:1px solid #C62828; border-radius:10px; padding:20px; }
    .cp-num { font-size:48px; font-weight:900; color:#C62828; }
    .cp-acoes { font-size:15px; color:#5B6576; }
    .cp-valor { font-size:22px; font-weight:900; color:#C62828; margin-top:10px; }
    .cp-rotulo { font-size:15px; font-weight:700; color:#0B1F3F; margin-top:16px; }
    .cp-desc { font-size:13px; color:#5B6576; margin-top:8px; }
    .lista-sintese { display:flex; flex-direction:column; gap:16px; margin-top:8px; }
    .linha-sintese { display:flex; align-items:center; gap:28px; background:#F2F4F7; border-radius:10px;
      padding:18px 24px; }
    .ls-valor { font-size:34px; font-weight:900; color:#2E7D32; width:260px; flex-shrink:0; }
    .ls-texto { font-size:18px; color:#0B1F3F; }
    '''

    return f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Análise Financeira do PAC - Apresentação</title>
<style>{css}</style>
</head>
<body>
{slide1}{slide2}{slide3}{slide4}{slide5}{slide6}{slide7}{slide8}{slide9}
</body>
</html>'''


def gerar_html_bytes(df):
    # Mesmo padrão de gerar_pptx_bytes: monta tudo em memória, sem tocar o
    # disco — usado pelo botão "ANÁLISE FINANCEIRA (HTML)" do painel.
    html_texto = montar_html_apresentacao(df)
    return html_texto.encode("utf-8")


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

    caminho_html = os.path.join(PASTA_BASE, "ANALISE_FINANCEIRA_PAC.html")
    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(montar_html_apresentacao(pac.df_original))
    print("HTML gerado:", caminho_html)


if __name__ == "__main__":
    main()
