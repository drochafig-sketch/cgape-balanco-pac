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
  - FINANCIAMENTO é tratado como parcela ESTADUAL (junto com Contrapartida/
    Complementar) — decisão do usuário (atualizada em 02/11/2026; antes era
    Federal), não é um fato da base (a planilha não diferencia dívida do
    Estado com aval federal de financiamento federal direto).
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
        if texto == "0":  # pedido do usuário em 04/11/2026: valor zerado aparece como "-"
            return "-"
    return f"{sinal}R$ {texto}"


def pct(parte, total):
    if not total:
        return "0,0%"
    # O "+ 0.0" evita o "-0,0%" quando a parte é zero e o total é negativo.
    return f"{round(parte / total * 100, 1) + 0.0:.1f}".replace(".", ",") + "%"


# ---------------------------------------------------------------------------
# Recorte (filtros do painel) — pedido do usuário em 03/11/2026: a
# apresentação segue os filtros ativos no painel, igual ao relatório PDF.
# Quando há filtro, o rodapé de cada slide diz qual é, pra um recorte nunca
# ser lido como a carteira inteira. Mesmas chaves de _filtrar_dataframe.
# ---------------------------------------------------------------------------
ROTULOS_FILTROS = [
    ("GESTAO", "Gestão"), ("EIXO", "Eixo"), ("STATUS", "Status"), ("FASE", "Fase"),
    ("ORGAO", "Órgão"), ("EXECUTOR", "Executor"), ("OBJETO", "Objeto"), ("MUNICIPIO", "Município"),
    ("FONTE", "Fonte"), ("CLAUSULA_SUSPENSIVA", "Cláusula suspensiva"),
    ("TERMO_COMPROMISSO", "Termo de compromisso"), ("OBRA_NAO_INICIADA", "Obra não iniciada"),
    ("DATAS_CONCLUSAO_FASE", "Conclusão da fase"), ("DATAS_VIGENCIA", "Vigência"),
    ("DATAS_CONCLUSAO_ATUAL", "Conclusão atual"),
]


def descrever_filtros(filtros, limite=150):
    if not filtros:
        return ""
    partes = []
    for chave, rotulo in ROTULOS_FILTROS:
        valores = filtros.get(chave) or []
        if not valores:
            continue
        if chave.startswith("DATAS_"):
            partes.append(f"{rotulo}: {len(valores)} data(s)")
        elif len(valores) <= 3:
            partes.append(f"{rotulo}: {', '.join(str(v) for v in valores)}")
        else:
            partes.append(f"{rotulo}: {len(valores)} selecionados")
    texto = "; ".join(partes)
    return texto if len(texto) <= limite else texto[:limite - 1].rstrip() + "…"


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
def novo_slide(prs, tema):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # layout em branco
    # numeração sequencial: depende de quais slides foram escolhidos no painel
    prs._pagina_atual += 1
    pagina, total_paginas = prs._pagina_atual, prs._total_paginas

    # fundo branco explícito
    add_rect(slide, 0, 0, LARGURA_IN, ALTURA_IN, fill=BRANCO)

    # faixa do cabeçalho
    add_rect(slide, 0, 0, LARGURA_IN, 1.1, fill=NAVY)
    add_text(slide, 0.56, 0.22, 12.5, 0.5, f"PAC - {tema}", 26, BRANCO, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    add_text(slide, 0.56, 0.74, 12.5, 0.3, "COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS ESTRATÉGICAS",
              12, AZUL_CLARO_TEXTO, negrito=True)
    add_text(slide, LARGURA_IN - 5.0, 0.18, 4.44, 0.35, "GOVERNO DO ESTADO DA BAHIA", 15,
              AZUL_CLARO_MARCA, negrito=True, alinhar=PP_ALIGN.RIGHT)
    add_text(slide, LARGURA_IN - 5.0, 0.56, 4.44, 0.35, "CASA CIVIL", 15, AZUL_CLARO_MARCA,
              negrito=True, alinhar=PP_ALIGN.RIGHT)

    # rodapé (+ recorte, quando a apresentação foi gerada com filtros —
    # guardado no próprio prs por montar_apresentacao)
    recorte = getattr(prs, "_recorte_pac", "")
    add_text(slide, 0.56, ALTURA_IN - 0.42, 16.5, 0.3,
              f"Fonte: {os.path.basename(pac.CAMINHO_PLANILHA) if hasattr(pac, 'CAMINHO_PLANILHA') else 'PANORAMA - PAC ORIGINAL - PAC SELEÇÕES - 2026'} · "
              f"posição em {pac.ultima_atualizacao.strftime('%m/%Y') if hasattr(pac, 'ultima_atualizacao') else date.today().strftime('%m/%Y')}"
              + (f" · Filtros: {recorte}" if recorte else ""),
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


def tabela_nativa(slide, x, y, w, h, cabecalho, linhas, larguras=None, tam_fonte=12,
                   linha_bold_indices=None, cor_cabecalho=NAVY):
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
        celula.fill.fore_color.rgb = cor(cor_cabecalho)
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
# Fases (slide "STATUS POR FASE E ORIGEM DO RECURSO") — valores da coluna
# FASE da base, na ordem do processo. Decisão do usuário em 03/11/2026:
# CONCLUÍDA entra como 4ª fase (sem ela o total de cada fonte não bate com o
# slide 2). Atualizado no mesmo dia: o tom da cor indica CONTRATADO (forte) x
# A CONTRATAR (claro), igual ao slide 2 — a fase é só o rótulo da linha.
# ---------------------------------------------------------------------------
FASES_ORIGEM = [
    ("CAPTAÇÃO DE RECURSO", "Captação de Recurso"),
    ("LICITAÇÃO", "Licitação"),
    ("EXECUÇÃO DO OBJETO", "Execução do Objeto"),
    ("CONCLUÍDA", "Concluída"),
]


def rotulo_fase(fase):
    return dict(FASES_ORIGEM).get(fase, str(fase).title())


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

    # a contratar por fonte (OGU/Financiamento/Estadual — sempre separados,
    # independente de qual "parcela" cada um entra abaixo)
    m["a_contratar_ogu"] = float(apoiado_aj.sum())
    m["a_contratar_financiamento"] = float(financiamento_aj.sum())
    m["a_contratar_estadual"] = float((contrapartida_aj + complementar_aj).sum())

    # detalhamento do contratado (colunas novas)
    m["contratado_ogu"] = float(df[pac.col_apoiado_contratado].sum())
    m["contratado_financiamento"] = float(df[pac.col_financiamento_contratado].sum())
    m["contratado_estadual"] = float(
        (df[pac.col_contrapartida_contratado] + df[pac.col_complementar_contratado]).sum()
    )
    # fonte x fase x (contratado, a contratar) — mesmas colunas e mesmo
    # critério da linha de detalhe do slide 2 (contratado detalhado por fonte
    # / a contratar ajustado por fonte), só que agrupadas pela FASE da ação;
    # por isso as somas por fonte batem, por construção, com o slide 2. Fase
    # fora de FASES_ORIGEM (não existe hoje) entra no fim, em vez de sumir.
    # Só entram as fases presentes no recorte: com o filtro de Fase do painel
    # tirando Concluída (por ex.), a linha some em vez de aparecer zerada.
    fase = df["FASE_TEXTO"].fillna("NÃO INFORMADA")
    presentes = set(fase)
    valores_fonte = {
        "ogu": (df[pac.col_apoiado_contratado], apoiado_aj),
        "financiamento": (df[pac.col_financiamento_contratado], financiamento_aj),
        "estadual": (df[pac.col_contrapartida_contratado] + df[pac.col_complementar_contratado],
                     contrapartida_aj + complementar_aj),
    }
    conhecidas = [f for f, _ in FASES_ORIGEM]
    m["fases_ordem"] = [f for f in conhecidas if f in presentes] + sorted(presentes - set(conhecidas))
    m["fonte_fase"] = {
        chave: {f: (float(c[fase == f].sum()), float(a[fase == f].sum())) for f in m["fases_ordem"]}
        for chave, (c, a) in valores_fonte.items()
    }
    # "Não discriminado" (pedido do usuário em 03/11/2026: categoria própria
    # nos gráficos de composição) = Valor Contratado menos o que está
    # detalhado por fonte, por fase. Só tem contratado (a contratar já vem
    # sempre por fonte). Com ele, as composições somam o investimento total.
    detalhado_linha = sum(c for c, _ in valores_fonte.values())
    nd_linha = contratado - detalhado_linha
    # Saldo líquido, sem truncar em zero: fica negativo onde o detalhamento
    # passa do Valor Contratado — é o que faz as composições fecharem sempre
    # com o investimento total, em qualquer recorte.
    m["fonte_fase"]["nao_discriminado"] = {
        f: (float(nd_linha[fase == f].sum()), 0.0) for f in m["fases_ordem"]
    }

    # fonte x eixo x (contratado, a contratar) — mesmo critério do fonte x fase,
    # pros slides de eixo. Eixos do maior pro menor investimento.
    eixo = df[pac.col_eixo].fillna("NÃO INFORMADO").astype(str)
    fontes_eixo = dict(valores_fonte, nao_discriminado=(nd_linha, nd_linha * 0))
    m["fonte_eixo"] = {
        chave: {e: (float(c[eixo == e].sum()), float(a[eixo == e].sum())) for e in set(eixo)}
        for chave, (c, a) in fontes_eixo.items()
    }
    m["eixos_ordem"] = sorted(set(eixo), key=lambda e: (-sum(sum(m["fonte_eixo"][k][e]) for k in m["fonte_eixo"]), e))

    # fase > objeto > fonte > (contratado, a contratar) — slides de composição
    # por fase (um por fase), mesmo critério dos cruzamentos acima
    objeto = df[pac.col_objeto].fillna("NÃO INFORMADO").astype(str)
    m["fase_objeto"] = {
        f: {o: {k: (float(c[(fase == f) & (objeto == o)].sum()), float(a[(fase == f) & (objeto == o)].sum()))
                for k, (c, a) in fontes_eixo.items()}
            for o in set(objeto[fase == f])}
        for f in m["fases_ordem"]
    }

    soma_detalhado = m["contratado_ogu"] + m["contratado_financiamento"] + m["contratado_estadual"]
    m["contratado_nao_discriminado"] = max(0.0, m["total_contratado"] - soma_detalhado)
    m["cobertura_detalhamento_pct"] = soma_detalhado / m["total_contratado"] if m["total_contratado"] else 0

    # Parcela Federal x Parcela Estadual (geral — contratado + a contratar):
    # decisão do usuário em 02/11/2026 — Financiamento passa a contar como
    # parcela ESTADUAL (antes contava como Federal, junto com OGU). Não é um
    # fato da base, é definição adotada só pra esta apresentação.
    m["contratado_parcela_federal"] = m["contratado_ogu"]
    m["a_contratar_parcela_federal"] = m["a_contratar_ogu"]
    m["total_parcela_federal"] = m["contratado_parcela_federal"] + m["a_contratar_parcela_federal"]

    m["contratado_parcela_estadual"] = m["contratado_financiamento"] + m["contratado_estadual"]
    m["a_contratar_parcela_estadual"] = m["a_contratar_financiamento"] + m["a_contratar_estadual"]
    m["total_parcela_estadual"] = m["contratado_parcela_estadual"] + m["a_contratar_parcela_estadual"]

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
# Faixa do slide Visão Geral (pedido do usuário em 05/11/2026): acima da barra
# Contratado x A Contratar, metade da espessura dela e alinhada com ela, abre
# cada status nas fontes (tom forte = contratado, tom claro = a contratar). O
# valor só aparece se couber no segmento; as cores estão na legenda.
ROTULO_DESTAQUE_FINANCIAMENTO = "FINANCIAMENTO A CONTRATAR"
ALT_FAIXA_FONTES = 0.45
FOLGA_FAIXA_FONTES = 0.03
ROTULOS_FAIXA_FONTES = {"ogu": "OGU", "financiamento": "Financiamento", "estadual": "Estadual",
                        "nao_discriminado": "Não discriminado"}


def _segmentos_faixa_fontes(m):
    # [(texto, cor de fundo, cor do texto, fração da barra), ...]. Saldo negativo
    # do Não discriminado fica de fora: cada status se divide entre as fontes positivas.
    total = m["total_investimento"]
    fontes = _fontes_por_fase(m)
    if total <= 0:
        return []
    segmentos = []
    for indice_valor, indice_cor, cor_texto, total_status in ((4, 2, BRANCO, m["total_contratado"]),
                                                              (5, 3, NAVY, m["total_a_contratar"])):
        positivos = sum(max(f[indice_valor], 0) for f in fontes)
        for fonte in fontes:
            valor = fonte[indice_valor]
            if valor > 0 and positivos > 0 and total_status > 0:
                fracao = total_status / total * valor / positivos
                segmentos.append((_texto_segmento_faixa(fonte[1], valor, 17.6 * fracao), fonte[indice_cor], cor_texto,
                                  fracao))
    return segmentos


def _texto_segmento_faixa(chave, valor, largura_pol):
    # Largura estimada pra fonte de 13 pt; a mesma decisão vale pro PPTX e pro HTML
    for texto in (f"{ROTULOS_FAIXA_FONTES[chave]} · {abreviar_moeda(valor)}", abreviar_moeda(valor)):
        if len(texto) * 0.095 + 0.2 <= largura_pol:
            return texto
    return ""


def _fontes_legenda_faixa(m):
    # (rótulo, cor) só das fontes que têm segmento na faixa
    return [(ROTULOS_FAIXA_FONTES[f[1]], f[2]) for f in _fontes_por_fase(m) if f[4] > 0 or f[5] > 0]


def _desenhar_legenda_faixa(s, m, y):
    x = 18.5
    for rotulo, cheia in reversed(_fontes_legenda_faixa(m)):
        largura = len(rotulo) * 0.1 + 0.3
        x -= largura
        add_text(s, x, y, largura, 0.3, rotulo, 13, CINZA_TEXTO, ancorar=MSO_ANCHOR.MIDDLE)
        x -= 0.22
        add_rect(s, x, y + 0.07, 0.16, 0.16, fill=cheia)
        x -= 0.2


def _desenhar_faixa_fontes(s, m, y):
    x = 0.9
    for texto, fundo, cor_texto, fracao in _segmentos_faixa_fontes(m):
        largura = 17.6 * fracao
        add_rect(s, x, y, largura, ALT_FAIXA_FONTES, fill=fundo)
        if texto:
            add_text(s, x + 0.05, y, largura - 0.05, ALT_FAIXA_FONTES, texto, 13, cor_texto, negrito=True,
                      ancorar=MSO_ANCHOR.MIDDLE)
        x += largura


def _desenhar_panorama(s, m, y, compacto=False):
    # Investimento total, barra contratado x a contratar e cards de status. A
    # versão compacta é a do slide Visão Geral, abaixo do quadro por fase.
    tam_valor, alt_valor = (44, 0.85) if compacto else (60, 1.2)
    add_text(s, 0.9, y, 10, 0.5, "INVESTIMENTO TOTAL DA CARTEIRA", 18 if compacto else 16, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, y + 0.35, 10, alt_valor + 0.1, abreviar_moeda(m["total_investimento"]), tam_valor, VERDE,
              negrito=True)
    if compacto and abreviar_moeda(m["a_contratar_financiamento"]) != "-":
        # Destaque pedido pelo usuário em 06/11/2026; some se o recorte não tiver esse valor
        add_text(s, 10.5, y, 8.0, 0.5, ROTULO_DESTAQUE_FINANCIAMENTO, 18, CINZA_TEXTO, negrito=True,
                  alinhar=PP_ALIGN.RIGHT)
        add_text(s, 10.5, y + 0.35, 8.0, alt_valor + 0.1, abreviar_moeda(m["a_contratar_financiamento"]), tam_valor,
                  COR_FINANCIAMENTO, negrito=True, alinhar=PP_ALIGN.RIGHT)
    y_texto = y + 0.35 + alt_valor
    add_text(s, 0.9, y_texto, 10, 0.4, f"{m['qtd_acoes']} ações do PAC acompanhadas pelo Estado da Bahia",
              16 if compacto else 14, CINZA_TEXTO)

    # barra empilhada contratado x a contratar
    y_barra = y_texto + 0.85
    add_text(s, 0.9, y_barra - 0.45, 6, 0.35, "CONTRATADO  x  A CONTRATAR", 17, NAVY, negrito=True)
    if compacto:
        _desenhar_legenda_faixa(s, m, y_barra - 0.4)
        _desenhar_faixa_fontes(s, m, y_barra)
        y_barra += ALT_FAIXA_FONTES + FOLGA_FAIXA_FONTES
    largura_barra = 17.6 if compacto else 17.0
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
    y_cards = y_barra + (1.1 if compacto else 1.5)
    alt_card, tam_num, y_num, y_rotulo = (1.3, 34, 0.12, 0.82) if compacto else (2.0, 44, 0.35, 1.3)
    cards = [
        ("CONCLUÍDO / INAUGURADO", m["qtd_concluido"], VERDE),
        ("EM OBRAS / ANDAMENTO", m["qtd_andamento"], AZUL),
        ("SELEÇÃO / PREPARATÓRIA / LICITAÇÃO", m["qtd_preparatoria"], LARANJA),
        ("PARALISADO", m["qtd_paralisado"], VERMELHO),
    ]
    # status sem nenhuma ação no recorte não ganha card (pedido do usuário em
    # 03/11/2026); os que sobram dividem a mesma largura total de 4 cards
    cards = [c for c in cards if c[1] > 0]
    espaco = 0.3
    largura_total = 17.6 if compacto else 4 * 4.1 + 3 * espaco
    largura_card = (largura_total - espaco * (len(cards) - 1)) / max(len(cards), 1)
    for i, (rotulo, qtd, c) in enumerate(cards):
        x = 0.9 + i * (largura_card + espaco)
        add_rect(s, x, y_cards, largura_card, alt_card, fill=BRANCO, line=CINZA_BORDA, arredondado=True, line_w=1)
        add_rect(s, x, y_cards, largura_card, 0.12, fill=c, arredondado=False)
        add_text(s, x + 0.25, y_cards + y_num, largura_card - 0.5, 0.9, str(qtd), tam_num, c, negrito=True)
        add_text(s, x + 0.25, y_cards + y_rotulo, largura_card - 0.5, 0.5, rotulo, 16 if compacto else 14,
                  CINZA_TEXTO, negrito=True)


def slide_1_panorama(prs, m):
    s = novo_slide(prs, "PANORAMA DA CARTEIRA")
    _desenhar_panorama(s, m, 1.5)
    return s


def slide_2_composicao_origem(prs, m):
    s = novo_slide(prs, "COMPOSIÇÃO POR ORIGEM DO RECURSO")
    total = m["total_investimento"]
    # mesmo gráfico dos slides por fase (pedido do usuário em 06/11/2026): barra
    # do total e, abaixo, as fontes proporcionais a ele, em tamanho maior
    y = _desenhar_grafico_fase(s, m, 1.55, _grafico_total_fontes(m), ROTULO_GRAFICO_ORIGEM, LEGENDA_FASES,
                               AMPLIADO_ORIGEM)

    # destaque federal x estadual x não discriminado — 3 cards lado a lado;
    # rótulo com espaço pra 2 linhas (o da Parcela Estadual quebra)
    y += 0.2
    for i, (rotulo, valor, cor_valor) in enumerate(_parcelas(m)):
        x = 0.9 + i * 5.85
        add_rect(s, x, y, 5.55, 1.9, fill=CINZA_CLARO, arredondado=True)
        add_text(s, x + 0.3, y + 0.16, 4.95, 0.55, rotulo, 14, CINZA_TEXTO, negrito=True)
        add_text(s, x + 0.3, y + 0.72, 4.95, 0.75, abreviar_moeda(valor), 32, cor_valor, negrito=True)
        add_text(s, x + 0.3, y + 1.45, 4.95, 0.35, pct(valor, total) + " do investimento total", 14, CINZA_TEXTO)

    add_text(s, 0.9, y + 2.1, 17.6, 0.5,
              "* Financiamento classificado como parcela Estadual por definição adotada nesta apresentação — "
              "a base não distingue dívida do Estado com aval federal de financiamento federal direto.",
              12, CINZA_TEXTO)
    return s


# Textos e linhas do slide de fases — compartilhados entre PPTX e HTML pra
# os dois saírem com exatamente os mesmos rótulos e números.
LEGENDA_FASES = "tom forte = CONTRATADO · tom claro = A CONTRATAR"


def _fontes_por_fase(m):
    # Fontes dos gráficos de composição (slides 2, 3 e 4), na ordem. "Não
    # discriminado" = contratado sem fonte detalhada; total = soma por fase,
    # pra bater com as linhas de fase (só tem contratado).
    nd = sum(c for c, _ in m["fonte_fase"]["nao_discriminado"].values())
    return [
        ("OGU (Apoiado)", "ogu", COR_OGU, COR_OGU_CLARO, m["contratado_ogu"], m["a_contratar_ogu"]),
        ("FINANCIAMENTO", "financiamento", COR_FINANCIAMENTO, COR_FINANCIAMENTO_CLARO,
         m["contratado_financiamento"], m["a_contratar_financiamento"]),
        ("RECURSO ESTADUAL (Contrapartida + Complementar)", "estadual", COR_ESTADUAL, COR_ESTADUAL_CLARO,
         m["contratado_estadual"], m["a_contratar_estadual"]),
        ("NÃO DISCRIMINADO (contratado sem fonte informada)", "nao_discriminado", COR_NAO_DISCRIMINADO,
         pac.clarear_hex(COR_NAO_DISCRIMINADO).lstrip("#"), nd, 0.0),
    ]


def _parcelas(m):
    # Cards do slide 2: Federal x Estadual x Não discriminado (pedido do
    # usuário em 03/11/2026) — os três somam o investimento total.
    nd = _fontes_por_fase(m)[-1][4]
    return [
        ("PARCELA FEDERAL (OGU)", m["total_parcela_federal"], AZUL),
        ("PARCELA ESTADUAL (Financiamento + Contrapartida + Complementar)*", m["total_parcela_estadual"], LARANJA),
        ("NÃO DISCRIMINADO (contratado sem fonte informada)", nd, COR_NAO_DISCRIMINADO),
    ]


# Geometria dos slides de composição com 2 níveis (fase x fonte), em pol:
# compacta pra caber 4 grupos x 4 linhas desde a entrada do Não discriminado.
ALT_NIVEL_1 = 0.5
AVANCO_NIVEL_1 = 0.56
ALT_NIVEL_2 = 0.29
ESPACO_GRUPO = 0.18   # respiro entre grupos (fontes / fases) — pedido do usuário

# Colunas das tabelas de status por fase e de eixo (pedido do usuário em
# 06/11/2026): rótulo | CONTRATADO | barra | A CONTRATAR | TOTAL. A barra não é
# mais 100% da linha: mostra o peso dela (ver _desenhar_barra_proporcional), com
# cada segmento encostado na coluna do seu valor.
X_COL_CONTRATADO = 6.4
LARG_COL_CONTRATADO = 1.9
X_BARRA_TABELA = 8.45
LARG_BARRA_TABELA = 5.45
X_COL_A_CONTRATAR = 14.05
LARG_COL_A_CONTRATAR = 2.1
X_COL_TOTAL = 16.25
LARG_COL_TOTAL = 2.25


def _slide_execucao(m):
    # Nº do slide de Execução do Contratado; None se não foi escolhido no painel
    return m.get("num_slide", {}).get("EXECUCAO")


def _nota_nao_discriminado(m, com_remissao=True):
    return ("* Não discriminado: valor contratado sem detalhamento por fonte de recurso na base, "
            "mostrado como categoria própria para a composição somar o investimento total."
            + (f" Ver slide {_slide_execucao(m)}." if com_remissao and _slide_execucao(m) else ""))


def _cabecalho_tabela(s, y):
    add_text(s, X_COL_CONTRATADO, y, LARG_COL_CONTRATADO, 0.28, "CONTRATADO", 11, CINZA_TEXTO, negrito=True,
              alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
    add_text(s, X_COL_A_CONTRATAR, y, LARG_COL_A_CONTRATAR, 0.28, "A CONTRATAR", 11, CINZA_TEXTO, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    add_text(s, X_COL_TOTAL, y, LARG_COL_TOTAL, 0.28, "TOTAL", 11, CINZA_TEXTO, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    add_rect(s, 0.9, y + 0.32, 17.6, 0.012, fill=CINZA_BORDA)
    return y + 0.42


def _fracoes_barra(c_val, a_val, base):
    # Partes da barra (contratado, a contratar) em relação à base. Saldo
    # negativo do Não discriminado não tem barra; a soma nunca passa de 100%.
    if base <= 0:
        return 0.0, 0.0
    fc = min(max(c_val, 0) / base, 1.0)
    return fc, min(max(a_val, 0) / base, 1.0 - fc)


def _desenhar_barra_proporcional(s, x, y, largura, altura, c_val, a_val, base, cheia, clara):
    # Trilho cinza e, por cima, contratado + a contratar; segmento com valor nunca some
    add_rect(s, x, y, largura, altura, fill=CINZA_CLARO)
    fc, fa = _fracoes_barra(c_val, a_val, base)
    larg_c = max(largura * fc, 0.03) if fc > 0 else 0
    larg_a = min(max(largura * fa, 0.03), largura - larg_c) if fa > 0 else 0
    if larg_c:
        add_rect(s, x, y, larg_c, altura, fill=cheia)
    if larg_a:
        add_rect(s, x + larg_c, y, larg_a, altura, fill=clara)


def _linha_tabela(s, nivel, y, rotulo, c_val, a_val, base, cheia, clara, altura=None):
    # Uma linha da tabela. Base da barra: nível 1 = investimento total, nível 2
    # = o grupo. Só valores, sem percentuais; destaque (negrito) no A CONTRATAR.
    x = 1.35 if nivel == 2 else 0.9
    altura = altura or (ALT_NIVEL_2 if nivel == 2 else ALT_NIVEL_1)
    tam_rotulo = {1: 16, 2: 13}[nivel]
    tam_status = {1: 14, 2: 12}[nivel]
    add_text(s, x, y, X_COL_CONTRATADO - 0.1 - x, altura, rotulo, tam_rotulo, NAVY, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    add_text(s, X_COL_CONTRATADO, y, LARG_COL_CONTRATADO, altura, abreviar_moeda(c_val), tam_status, CINZA_TEXTO,
              alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
    _desenhar_barra_proporcional(s, X_BARRA_TABELA, y + altura * 0.18, LARG_BARRA_TABELA, altura * 0.64, c_val, a_val,
                                 base, cheia, clara)
    add_text(s, X_COL_A_CONTRATAR, y, LARG_COL_A_CONTRATAR, altura, abreviar_moeda(a_val), tam_status, NAVY,
              negrito=True, ancorar=MSO_ANCHOR.MIDDLE)
    add_text(s, X_COL_TOTAL, y, LARG_COL_TOTAL, altura, abreviar_moeda(c_val + a_val), tam_status, CINZA_TEXTO,
              ancorar=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------------------
# Slide "COMPOSIÇÃO POR FASE" — pedido do usuário em 03/11/2026: hierarquia
# de orçamento invertida em relação ao slide "STATUS POR FASE E ORIGEM":
# cada FASE (nível 1) aberta nas FONTES (nível 2), com barras 100% (ver
# _linha_tabela). Na fase, só contratado (cinza escuro)
# x a contratar (cinza claro), sem distinguir fonte; nas fontes, a cor
# identifica a fonte (mesmas do slide 2) e o tom, contratado (forte) x a
# contratar (claro). Mesmos números de m["fonte_fase"]. Sem os cards do
# slide 2 — 4 fases x 3 fontes não cabem com eles.
# ---------------------------------------------------------------------------
LEGENDA_COMPOSICAO_FASE = "cinza escuro / tom forte = CONTRATADO · cinza claro / tom claro = A CONTRATAR"
CINZA_FASE_CONTRATADO = "5B6576"
CINZA_FASE_A_CONTRATAR = "C3C8D0"


ROTULOS_CURTOS_FONTE = {"ogu": "OGU (Apoiado)", "financiamento": "Financiamento", "estadual": "Recurso Estadual",
                        "nao_discriminado": "Não discriminado"}


def _fases_por_fonte(m):
    # [(rótulo da fase, c, a, [(rótulo curto da fonte, c, a, cheia, clara), ...]), ...]
    linhas = []
    for f in m["fases_ordem"]:
        fontes = [(ROTULOS_CURTOS_FONTE[k], *m["fonte_fase"][k][f], cheia, clara)
                  for _, k, cheia, clara, _, _ in _fontes_por_fase(m)]
        linhas.append((rotulo_fase(f), sum(x[1] for x in fontes), sum(x[2] for x in fontes), fontes))
    return linhas


def slide_status_fase_origem(prs, m):
    # Pedido do usuário em 03/11/2026: cada fonte (nível 1) aberta nas fases
    # (nível 2), em tabela com barras 100% (ver _linha_tabela). Tom forte =
    # contratado, tom claro = a contratar; a fase é só o rótulo.
    s = novo_slide(prs, "STATUS POR FASE E ORIGEM DO RECURSO")
    y = 1.55
    add_text(s, 0.9, y, 9.0, 0.4, "INVESTIMENTO POR FONTE E FASE (contratado + a contratar)*", 17, CINZA_TEXTO,
              negrito=True)
    add_text(s, 9.9, y, 8.6, 0.4, LEGENDA_FASES, 12, CINZA_TEXTO, alinhar=PP_ALIGN.RIGHT,
              ancorar=MSO_ANCHOR.MIDDLE)
    y = _cabecalho_tabela(s, y + 0.5)
    for rotulo, chave, cheia, clara, c_val, a_val in _fontes_por_fase(m):
        _linha_tabela(s, 1, y, rotulo, c_val, a_val, m["total_investimento"], cheia, clara)
        y += AVANCO_NIVEL_1
        for f in m["fases_ordem"]:
            c_fase, a_fase = m["fonte_fase"][chave][f]
            _linha_tabela(s, 2, y, rotulo_fase(f), c_fase, a_fase, c_val + a_val, cheia, clara)
            y += ALT_NIVEL_2
        y += ESPACO_GRUPO
    add_text(s, 0.9, y + 0.05, 17.6, 0.3, _nota_nao_discriminado(m), 12, CINZA_TEXTO)
    return s


# ---------------------------------------------------------------------------
# Slides de EIXO — pedido do usuário em 04/11/2026: mesmo layout de tabela dos
# slides de fase, em duas hierarquias (eixo > fonte e fonte > eixo). São 9
# eixos: mesmo omitindo as linhas zeradas, não cabem num slide só com letra
# legível, então a tabela continua no slide seguinte, sem partir um grupo.
# ---------------------------------------------------------------------------
LEGENDA_EIXO_FONTE = LEGENDA_COMPOSICAO_FASE
ALT_EIXO_NIVEL_1 = 0.45
AVANCO_EIXO_NIVEL_1 = 0.5
ALT_EIXO_NIVEL_2 = 0.25
ESPACO_EIXO_GRUPO = 0.12
CAPACIDADE_TABELA_EIXO = 7.2   # altura útil de um slide (pol); vale pro HTML também


def _linha_com_valor(c_val, a_val):
    return abs(c_val) + abs(a_val) >= 0.5


def _grupos_eixo_fonte(m):
    # [(rótulo, c, a, cheia, clara, [(rótulo, c, a, cheia, clara), ...]), ...]
    grupos = []
    for e in m["eixos_ordem"]:
        linhas = [(ROTULOS_CURTOS_FONTE[k], *m["fonte_eixo"][k][e], cheia, clara)
                  for _, k, cheia, clara, _, _ in _fontes_por_fase(m) if _linha_com_valor(*m["fonte_eixo"][k][e])]
        if linhas:
            grupos.append((e, sum(x[1] for x in linhas), sum(x[2] for x in linhas), CINZA_FASE_CONTRATADO,
                           CINZA_FASE_A_CONTRATAR, linhas))
    return grupos


def _grupos_fonte_eixo(m):
    grupos = []
    for rotulo, k, cheia, clara, _, _ in _fontes_por_fase(m):
        linhas = [(e, *m["fonte_eixo"][k][e], cheia, clara) for e in m["eixos_ordem"]
                  if _linha_com_valor(*m["fonte_eixo"][k][e])]
        if linhas:
            grupos.append((rotulo, sum(x[1] for x in linhas), sum(x[2] for x in linhas), cheia, clara, linhas))
    return grupos


def _paginar_grupos(grupos):
    paginas, usado = [[]], 0.0
    for grupo in grupos:
        altura = AVANCO_EIXO_NIVEL_1 + ALT_EIXO_NIVEL_2 * len(grupo[5]) + ESPACO_EIXO_GRUPO
        if paginas[-1] and usado + altura > CAPACIDADE_TABELA_EIXO:
            paginas.append([])
            usado = 0.0
        paginas[-1].append(grupo)
        usado += altura
    return paginas


def _tema_paginado(tema, i, n):
    return f"{tema} ({i}/{n})" if n > 1 else tema


def _nota_eixos(m, grupos):
    nota = _nota_nao_discriminado(m) + " Linhas sem valor são omitidas."
    if any(linha[1] < 0 for grupo in grupos for linha in grupo[5]):
        nota += " Valor negativo: detalhamento por fonte maior que o Valor Contratado na base."
    return nota


def _desenhar_grupos(s, y, pagina, total):
    # Grupos de 2 níveis de uma página (rótulo > linhas); devolve o y final.
    # Barra do grupo em relação ao total; a das linhas, em relação ao grupo.
    for rotulo, c_val, a_val, cheia, clara, linhas in pagina:
        _linha_tabela(s, 1, y, rotulo, c_val, a_val, total, cheia, clara, altura=ALT_EIXO_NIVEL_1)
        y += AVANCO_EIXO_NIVEL_1
        for rotulo_2, c_2, a_2, cheia_2, clara_2 in linhas:
            _linha_tabela(s, 2, y, rotulo_2, c_2, a_2, c_val + a_val, cheia_2, clara_2, altura=ALT_EIXO_NIVEL_2)
            y += ALT_EIXO_NIVEL_2
        y += ESPACO_EIXO_GRUPO
    return y


def _slides_tabela_eixo(prs, m, tema, rotulo_secao, legenda, grupos):
    paginas = _paginar_grupos(grupos)
    for i, pagina in enumerate(paginas, 1):
        s = novo_slide(prs, _tema_paginado(tema, i, len(paginas)))
        y = 1.55
        add_text(s, 0.9, y, 9.0, 0.4, rotulo_secao, 17, CINZA_TEXTO, negrito=True)
        add_text(s, 9.9, y, 8.6, 0.4, legenda, 12, CINZA_TEXTO, alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
        y = _desenhar_grupos(s, _cabecalho_tabela(s, y + 0.5), pagina, m["total_investimento"])
        add_text(s, 0.9, y + 0.05, 17.6, 0.5, _nota_eixos(m, grupos), 12, CINZA_TEXTO)


# ---------------------------------------------------------------------------
# Slides "COMPOSIÇÃO POR FASE" — pedido do usuário em 05/11/2026: um slide por
# fase. Começa com o gráfico da fase aberto por fonte (barra da fase em relação
# ao investimento total; barra da fonte em relação à fase; tom forte =
# contratado, tom claro = a contratar). Abaixo, duas tabelas de objetos, com as
# fontes em colunas: primeiro A CONTRATAR, depois CONTRATADOS. Um objeto com
# as duas situações aparece nas duas. O que não couber continua no slide
# seguinte, sem repetir o gráfico.
# ---------------------------------------------------------------------------
ROTULO_GRAFICO_FASE = "INVESTIMENTO TOTAL, POR FASE E FONTE (contratado + a contratar)*"
COLUNAS_FONTES_OBJETOS = ["OGU (APOIADO)", "FINANCIAMENTO", "RECURSO ESTADUAL", "NÃO DISCRIMINADO*", "TOTAL"]
LARGURAS_OBJETOS = [4.6, 2.6, 2.6, 2.6, 2.6, 2.6]
# (título da 1ª coluna, índice em (contratado, a contratar), cor do cabeçalho =
# a da barra Contratado x A Contratar do panorama)
BLOCOS_OBJETOS = [("OBJETO A CONTRATAR", 1, LARANJA), ("OBJETO CONTRATADO", 0, VERDE)]
ALT_LINHA_OBJETOS = 0.36
ESPACO_BLOCO = 0.25
Y_INICIO_FASE = 1.4
Y_APOS_GRAFICO = 4.05
Y_LIMITE_TABELAS = 10.2   # abaixo disso, só a nota e o rodapé
ROTULO_GRAFICO_ORIGEM = "INVESTIMENTO TOTAL, POR FONTE (contratado + a contratar)"
# (altura da linha do total, (fonte do rótulo, fonte dos valores), o mesmo pras fontes)
MEDIDAS_GRAFICO_FASE = (ALT_NIVEL_1, (19, 16), ALT_NIVEL_2, (14, 13))
AMPLIADO_ORIGEM = (0.9, (22, 18), 0.8, (18, 16))
X_BARRA_GRAFICO = 6.4
LARG_BARRA_GRAFICO = 5.2
X_GRAFICO_TOTAL = 11.85
X_GRAFICO_CONTRATADO = 14.4
X_GRAFICO_A_CONTRATAR = 16.9
LARG_GRAFICO_VALOR = 1.6


def _linhas_objetos(m, fase, indice):
    # Linhas de uma tabela de objetos, do maior pro menor, com o TOTAL no fim;
    # vazia se a fase não tem objeto nessa situação
    chaves = [k for _, k, *_ in _fontes_por_fase(m)]
    objetos = [(objeto, [fontes[k][indice] for k in chaves]) for objeto, fontes in m["fase_objeto"][fase].items()]
    objetos = sorted((o for o in objetos if any(abs(v) >= 0.5 for v in o[1])), key=lambda o: (-sum(o[1]), o[0]))
    if not objetos:
        return []
    totais = [sum(coluna) for coluna in zip(*(valores for _, valores in objetos))]
    return [[rotulo] + [abreviar_moeda(v) for v in valores + [sum(valores)]]
            for rotulo, valores in objetos + [("TOTAL", totais)]]


def _paginar_blocos_objetos(blocos):
    # [[(cabeçalho, linhas, tem_total, cor), ...], ...]: uma tabela só começa num slide
    # se couber um pedaço útil, e a linha de TOTAL nunca fica sozinha no seguinte
    paginas = [[]]
    livre = Y_LIMITE_TABELAS - Y_APOS_GRAFICO
    for titulo, linhas, cor_bloco in blocos:
        resto, continuacao = list(linhas), False
        while resto:
            cabem = int(livre / ALT_LINHA_OBJETOS + 1e-9) - 1   # -1 do cabeçalho
            if cabem < min(3, len(resto)):
                paginas.append([])
                livre = Y_LIMITE_TABELAS - Y_INICIO_FASE
                continue
            n = cabem - 1 if len(resto) - cabem == 1 else min(cabem, len(resto))
            parte, resto = resto[:n], resto[n:]
            cabecalho = [titulo + (" (cont.)" if continuacao else "")] + COLUNAS_FONTES_OBJETOS
            paginas[-1].append((cabecalho, parte, not resto, cor_bloco))
            livre -= ALT_LINHA_OBJETOS * (len(parte) + 1) + ESPACO_BLOCO
            continuacao = True
    return paginas


def _paginas_composicao_fase(m):
    # [(tema do slide, (fase, c, a, fontes) só no 1º slide da fase, blocos, nota), ...]
    saida = []
    for fase, grafico in zip(m["fases_ordem"], _fases_por_fonte(m)):
        paginas = _paginar_blocos_objetos([(titulo, _linhas_objetos(m, fase, i), c) for titulo, i, c in BLOCOS_OBJETOS])
        nota = _nota_nao_discriminado(m)
        if any(fontes["nao_discriminado"][0] <= -0.5 for fontes in m["fase_objeto"][fase].values()):
            nota += " Valor negativo: detalhamento por fonte maior que o Valor Contratado na base."
        for i, pagina in enumerate(paginas, 1):
            tema = _tema_paginado(f"COMPOSIÇÃO POR FASE - {grafico[0].upper()}", i, len(paginas))
            saida.append((tema, grafico if i == 1 else None, pagina, nota))
    return saida


def _linha_grafico_fase(s, nivel, y, rotulo, c_val, a_val, base, cheia, clara, altura, tam_rotulo, tam_valor):
    x = 0.9 if nivel == 1 else 1.35
    add_text(s, x, y, X_BARRA_GRAFICO - 0.1 - x, altura, rotulo, tam_rotulo, NAVY, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    _desenhar_barra_proporcional(s, X_BARRA_GRAFICO, y + altura * 0.18, LARG_BARRA_GRAFICO, altura * 0.64, c_val, a_val,
                                 base, cheia, clara)
    add_text(s, X_GRAFICO_TOTAL, y, LARG_GRAFICO_VALOR + 0.5, altura, abreviar_moeda(c_val + a_val), tam_valor, NAVY,
              negrito=True, ancorar=MSO_ANCHOR.MIDDLE)
    for x_valor, valor in ((X_GRAFICO_CONTRATADO, c_val), (X_GRAFICO_A_CONTRATAR, a_val)):
        add_text(s, x_valor, y, LARG_GRAFICO_VALOR, altura, abreviar_moeda(valor), tam_valor, CINZA_TEXTO,
                  ancorar=MSO_ANCHOR.MIDDLE)


def _grafico_total_fontes(m):
    # Gráfico do slide de origem do recurso: o total da carteira no lugar da fase
    fontes = [(ROTULOS_CURTOS_FONTE[k], c_val, a_val, cheia, clara)
              for _, k, cheia, clara, c_val, a_val in _fontes_por_fase(m)]
    return ("Total da carteira", sum(f[1] for f in fontes), sum(f[2] for f in fontes), fontes)


def _desenhar_grafico_fase(s, m, y, grafico, rotulo_secao=ROTULO_GRAFICO_FASE, legenda=None, medidas=MEDIDAS_GRAFICO_FASE):
    # Devolve o y logo abaixo da última linha
    rotulo, c_val, a_val, fontes = grafico
    alt_1, tam_1, alt_2, tam_2 = medidas
    if legenda:
        add_text(s, 0.9, y, 9.0, 0.4, rotulo_secao, 17, CINZA_TEXTO, negrito=True)
        add_text(s, 9.9, y, 8.6, 0.4, legenda, 12, CINZA_TEXTO, alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
    else:
        add_text(s, 0.9, y, 17.6, 0.4, rotulo_secao, 17, CINZA_TEXTO, negrito=True)
    y += 0.42
    for x_valor, titulo in ((X_GRAFICO_TOTAL, "TOTAL"), (X_GRAFICO_CONTRATADO, "CONTRATADO"),
                            (X_GRAFICO_A_CONTRATAR, "A CONTRATAR")):
        add_text(s, x_valor, y, LARG_GRAFICO_VALOR, 0.28, titulo, 11, CINZA_TEXTO, negrito=True,
                  ancorar=MSO_ANCHOR.MIDDLE)
    add_rect(s, 0.9, y + 0.32, 17.6, 0.012, fill=CINZA_BORDA)
    y += 0.4
    _linha_grafico_fase(s, 1, y, rotulo, c_val, a_val, m["total_investimento"], CINZA_FASE_CONTRATADO,
                        CINZA_FASE_A_CONTRATAR, alt_1, *tam_1)
    y += alt_1 + AVANCO_NIVEL_1 - ALT_NIVEL_1
    for rotulo_fonte, c_fonte, a_fonte, cheia, clara in fontes:
        _linha_grafico_fase(s, 2, y, rotulo_fonte, c_fonte, a_fonte, c_val + a_val, cheia, clara, alt_2, *tam_2)
        y += alt_2
    return y


def slide_composicao_fase(prs, m):
    for tema, grafico, blocos, nota in _paginas_composicao_fase(m):
        s = novo_slide(prs, tema)
        y = Y_INICIO_FASE
        if grafico:
            _desenhar_grafico_fase(s, m, y, grafico)
            y = Y_APOS_GRAFICO
        for cabecalho, linhas, tem_total, cor_bloco in blocos:
            altura = ALT_LINHA_OBJETOS * (len(linhas) + 1)
            tabela_nativa(s, 0.9, y, 17.6, altura, cabecalho, linhas, larguras=LARGURAS_OBJETOS, tam_fonte=14,
                          linha_bold_indices={len(linhas) - 1} if tem_total else None, cor_cabecalho=cor_bloco)
            y += altura + ESPACO_BLOCO
        add_text(s, 0.9, y - ESPACO_BLOCO + 0.1, 17.6, 0.5, nota, 12, CINZA_TEXTO)


def slide_eixo_fonte(prs, m):
    _slides_tabela_eixo(prs, m, "INVESTIMENTO POR EIXO E FONTE",
                        "INVESTIMENTO TOTAL, POR EIXO E FONTE (contratado + a contratar)*", LEGENDA_EIXO_FONTE,
                        _grupos_eixo_fonte(m))


def slide_fonte_eixo(prs, m):
    _slides_tabela_eixo(prs, m, "INVESTIMENTO POR FONTE E EIXO",
                        "INVESTIMENTO TOTAL, POR FONTE E EIXO (contratado + a contratar)*", LEGENDA_FASES,
                        _grupos_fonte_eixo(m))


def slide_3_execucao_contratado(prs, m):
    s = novo_slide(prs, "EXECUÇÃO DO CONTRATADO, POR FONTE")
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
    s = novo_slide(prs, "O QUE AINDA PRECISA SER CONTRATADO")
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
    s = novo_slide(prs, "TABELA-RESUMO POR FONTE")
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
    linhas.append(["SUBTOTAL FEDERAL (OGU)", abreviar_moeda(m["contratado_parcela_federal"]),
                    abreviar_moeda(m["a_contratar_parcela_federal"]), abreviar_moeda(m["total_parcela_federal"]),
                    pct(m["total_parcela_federal"], total_geral)])
    linhas.append(["SUBTOTAL ESTADUAL (Financiamento + Contrapartida + Complementar)",
                    abreviar_moeda(m["contratado_parcela_estadual"]),
                    abreviar_moeda(m["a_contratar_parcela_estadual"]), abreviar_moeda(m["total_parcela_estadual"]),
                    pct(m["total_parcela_estadual"], total_geral)])
    linhas.append(["TOTAL GERAL", abreviar_moeda(m["total_contratado"]), abreviar_moeda(m["total_a_contratar"]),
                    abreviar_moeda(total_geral), "100,0%"])

    tabela_nativa(s, 0.9, 2.0, 18.2, 5.6, cabecalho, linhas,
                  larguras=[6.2, 3.0, 3.0, 3.0, 3.0],
                  tam_fonte=16, linha_bold_indices={4, 5, 6})
    add_text(s, 0.9, 7.9, 17, 0.4,
              "\"Não discriminado\" = contratado sem detalhamento por fonte preenchido na base"
              + (f" (ver slide {_slide_execucao(m)})." if _slide_execucao(m) else "."),
              13, CINZA_TEXTO)
    return s


# ---------------------------------------------------------------------------
# Slide "QUADRO RESUMO DE INVESTIMENTOS POR FASE" — pedido do usuário em
# 04/11/2026, com o formato do quadro "Consolidado Recurso Federal Novo PAC",
# mas calculado da planilha: mesmos números dos slides de composição (fase x
# fonte, contratado + a contratar), seguindo os filtros. Não discriminado em
# coluna própria pra fechar com o investimento total.
# ---------------------------------------------------------------------------
TEMA_QUADRO_RESUMO = "QUADRO RESUMO DE INVESTIMENTOS POR FASE"
SUBTITULO_QUADRO_RESUMO = "INVESTIMENTO POR FASE E FONTE (contratado + a contratar)*"
LEGENDA_QUADRO_RESUMO = "Valores expressos em reais (R$)"
LEGENDA_QUADRO_ABREVIADA = "Valores expressos em reais (Mi = milhões · Bi = bilhões)"
CABECALHO_QUADRO_RESUMO = ["FASE", "OGU (APOIADO)", "FINANCIAMENTO", "CONTRAPARTIDA + COMPLEMENTAR",
                           "NÃO DISCRIMINADO", "TOTAL DA FASE"]


def _reais_inteiros(valor):
    # Reais inteiros, como no quadro de referência; zero vira "-", como no resto do deck
    arredondado = round(float(valor))
    if arredondado == 0:
        return "-"
    return ("-" if arredondado < 0 else "") + "R$ " + f"{abs(arredondado):,}".replace(",", ".")


def _mi_bi_2_casas(valor):
    # R$ x,xx Bi a partir de 1 bilhão (já contando o arredondamento pra 2
    # casas) e R$ x,xx Mi abaixo disso — slide Visão Geral. O que arredonda pra
    # zero vira "-", como no resto do deck.
    absoluto = abs(float(valor))
    if absoluto >= 999_995_000:
        numero, unidade = absoluto / 1e9, "Bi"
    else:
        numero, unidade = absoluto / 1e6, "Mi"
    texto = f"{numero:,.2f}"
    if texto == "0.00":
        return "-"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if valor < 0 else "") + f"R$ {texto} {unidade}"


def _linhas_quadro_resumo(m, formatar=_reais_inteiros):
    # Fontes na ordem de _fontes_por_fase: OGU, Financiamento, Estadual, Não discriminado
    por_fase = [(rotulo, [c + a for _, c, a, _, _ in fontes]) for rotulo, _, _, fontes in _fases_por_fonte(m)]
    totais = [sum(col) for col in zip(*(valores for _, valores in por_fase))]
    return [[rotulo] + [formatar(v) for v in valores + [sum(valores)]]
            for rotulo, valores in por_fase + [("TOTAL GERAL", totais)]]


def _desenhar_quadro_resumo(s, m, y, alt_linha=0.7, folga_tabela=0.6, formatar=_reais_inteiros,
                            legenda=LEGENDA_QUADRO_RESUMO, tam_tabela=16, tam_nota=12):
    # Subtítulo, legenda, tabela e nota; devolve o y logo abaixo da nota
    add_text(s, 0.9, y, 9.0, 0.4, SUBTITULO_QUADRO_RESUMO, 17, CINZA_TEXTO, negrito=True)
    add_text(s, 9.9, y, 8.6, 0.4, legenda, tam_nota, CINZA_TEXTO, alinhar=PP_ALIGN.RIGHT,
              ancorar=MSO_ANCHOR.MIDDLE)
    linhas = _linhas_quadro_resumo(m, formatar)
    alt_tabela = alt_linha * (len(linhas) + 1)
    tabela_nativa(s, 0.9, y + folga_tabela, 17.6, alt_tabela, CABECALHO_QUADRO_RESUMO, linhas,
                  larguras=[3.2, 2.9, 3.0, 3.1, 2.4, 3.0], tam_fonte=tam_tabela, linha_bold_indices={len(linhas) - 1})
    y_nota = y + folga_tabela + alt_tabela + 0.2
    add_text(s, 0.9, y_nota, 17.6, 0.3, _nota_nao_discriminado(m, com_remissao=False), tam_nota, CINZA_TEXTO)
    return y_nota + 0.3


def slide_quadro_resumo_fase(prs, m):
    s = novo_slide(prs, TEMA_QUADRO_RESUMO)
    _desenhar_quadro_resumo(s, m, 1.55)
    return s


def slide_visao_geral(prs, m):
    # Pedido do usuário em 05/11/2026: quadro por fase em cima, panorama embaixo
    s = novo_slide(prs, "VISÃO GERAL DA CARTEIRA")
    y = _desenhar_quadro_resumo(s, m, 1.35, alt_linha=0.55, folga_tabela=0.45, formatar=_mi_bi_2_casas,
                                legenda=LEGENDA_QUADRO_ABREVIADA, tam_tabela=20, tam_nota=14)
    _desenhar_panorama(s, m, y + 0.15, compacto=True)
    return s


def slide_6_por_eixo(prs, m):
    s = novo_slide(prs, "DETALHE POR EIXO DO PAC")
    cabecalho = ["EIXO", "Nº AÇÕES", "INVESTIMENTO", "CONTRATADO", "A CONTRATAR"]
    linhas = []
    for eixo, linha in m["por_eixo"].iterrows():
        linhas.append([eixo, int(linha["QTD"]), abreviar_moeda(linha["INVESTIMENTO"]),
                        abreviar_moeda(linha["CONTRATADO"]), abreviar_moeda(linha["A_CONTRATAR"])])
    tabela_nativa(s, 0.9, 1.9, 18.2, 6.6, cabecalho, linhas,
                  larguras=[6.5, 2.2, 3.17, 3.17, 3.16], tam_fonte=15)
    return s


def slide_7_top10(prs, m):
    s = novo_slide(prs, "MAIORES AÇÕES A CONTRATAR")
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
    s = novo_slide(prs, "PENDÊNCIAS QUE CONDICIONAM A CONTRATAÇÃO")
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
    s = novo_slide(prs, "SÍNTESE E PLEITO")
    mensagens = [
        (abreviar_moeda(m["total_a_contratar"]), "ainda precisam ser contratados para concluir a carteira do PAC-BA"),
        (abreviar_moeda(m["a_contratar_parcela_federal"]), "desse valor dependem de fonte Federal (OGU)"),
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


def _html_segmentos_barra(c_val, a_val, base, cheia, clara):
    # Contratado + a contratar em relação à base; segmento com valor nunca some
    fc, fa = _fracoes_barra(c_val, a_val, base)
    return "".join(f'<div class="barra-fill" style="width:{f * 100:.4f}%;min-width:3px;background:#{c}" title="{t}"></div>'
                   for f, c, t in ((fc, cheia, "Contratado"), (fa, clara, "A contratar")) if f > 0)


def _html_linha_tabela(rotulo, c_val, a_val, base, cheia, clara):
    # Igual ao PPTX, ver _linha_tabela. O nível 2 só muda o estilo, via CSS .nivel-2.
    return f'''<div class="barra-linha">
      <div class="barra-rotulo" style="width:340px">{_html_esc(rotulo)}</div>
      <div class="barra-col-valor col-contratado">{_html_esc(abreviar_moeda(c_val))}</div>
      <div class="barra-trilho">{_html_segmentos_barra(c_val, a_val, base, cheia, clara)}</div>
      <div class="barra-col-valor col-a-contratar">{_html_esc(abreviar_moeda(a_val))}</div>
      <div class="barra-col-valor">{_html_esc(abreviar_moeda(c_val + a_val))}</div>
    </div>'''


def _html_blocos_grupos(pagina, total):
    # Grupos de 2 níveis de uma página (rótulo > linhas), ver _desenhar_grupos
    return "".join(f'''<div class="bloco-fonte-fase compacto denso">
      {_html_linha_tabela(rotulo, c_val, a_val, total, cheia, clara)}
      <div class="nivel-2">{"".join(_html_linha_tabela(r, c, a, c_val + a_val, f, l) for r, c, a, f, l in linhas)}</div>
    </div>''' for rotulo, c_val, a_val, cheia, clara, linhas in pagina)


def _html_linha_grafico_fase(nivel, rotulo, c_val, a_val, base, cheia, clara):
    # Linha do gráfico dos slides por fase (igual ao PPTX, ver _linha_grafico_fase)
    return f'''<div class="grafico-fase-linha nivel-{nivel}">
      <div class="gf-rotulo">{_html_esc(rotulo)}</div>
      <div class="barra-trilho">{_html_segmentos_barra(c_val, a_val, base, cheia, clara)}</div>
      <div class="gf-total">{_html_esc(abreviar_moeda(c_val + a_val))}</div>
      <div>{_html_esc(abreviar_moeda(c_val))}</div>
      <div>{_html_esc(abreviar_moeda(a_val))}</div>
    </div>'''


def _html_grafico_fase(m, grafico, rotulo_secao=ROTULO_GRAFICO_FASE, legenda=None):
    rotulo, c_val, a_val, fontes = grafico
    html_legenda = (f'<span style="font-size:13px;font-weight:400;color:#5B6576;text-transform:none">{_html_esc(legenda)}</span>'
                    if legenda else "")
    return f'''<div class="rotulo-secao" style="display:flex;justify-content:space-between;align-items:baseline">
        <span>{_html_esc(rotulo_secao)}</span>{html_legenda}
      </div>
      <div class="grafico-fase-linha gf-cabecalho"><div></div><div></div><div>TOTAL</div><div>CONTRATADO</div><div>A CONTRATAR</div></div>
      {_html_linha_grafico_fase(1, rotulo, c_val, a_val, m["total_investimento"], CINZA_FASE_CONTRATADO, CINZA_FASE_A_CONTRATAR)}
      {"".join(_html_linha_grafico_fase(2, r, c, a, c_val + a_val, cheia, clara) for r, c, a, cheia, clara in fontes)}'''


def _html_cabecalho_tabela():
    return ('<div class="barra-linha cabecalho-tabela"><div class="barra-rotulo" style="width:340px"></div>'
            '<div class="barra-col-valor col-contratado">CONTRATADO</div>'
            '<div class="barra-trilho" style="background:none;height:auto"></div>'
            '<div class="barra-col-valor">A CONTRATAR</div><div class="barra-col-valor">TOTAL</div></div>')


def _html_tabela(cabecalho, linhas, linhas_bold=(), cor_cabecalho=None):
    estilo = f' style="background:#{cor_cabecalho}"' if cor_cabecalho else ""
    ths = "".join(f"<th{estilo}>{_html_esc(c)}</th>" for c in cabecalho)
    trs = []
    for i, linha in enumerate(linhas):
        classe = " class=\"linha-bold\"" if i in linhas_bold else ""
        tds = "".join(f"<td>{_html_esc(v)}</td>" for v in linha)
        trs.append(f"<tr{classe}>{tds}</tr>")
    return f'<table class="pac-tabela"><thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table>'


def _html_cabecalho_rodape_base(tema, pagina, total_paginas, conteudo_html, recorte=""):
    nome_planilha = os.path.basename(pac.CAMINHO_PLANILHA) if hasattr(pac, "CAMINHO_PLANILHA") else \
        "PANORAMA - PAC ORIGINAL - PAC SELEÇÕES - 2026"
    posicao = pac.ultima_atualizacao.strftime("%m/%Y") if hasattr(pac, "ultima_atualizacao") else date.today().strftime("%m/%Y")
    return f'''<section class="slide">
      <div class="cabecalho">
        <div class="titulo">PAC - {_html_esc(tema)}</div>
        <div class="subtitulo">COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS ESTRATÉGICAS</div>
        <div class="governo"><div>GOVERNO DO ESTADO DA BAHIA</div><div>CASA CIVIL</div></div>
      </div>
      <div class="conteudo">{conteudo_html}</div>
      <div class="rodape">
        <span>Fonte: {_html_esc(nome_planilha)} · posição em {_html_esc(posicao)}{(" · Filtros: " + _html_esc(recorte)) if recorte else ""}</span>
        <span>{pagina}/{total_paginas}</span>
      </div>
    </section>'''


def montar_html_apresentacao(df, filtros=None, secoes=None):
    m = calcular_metricas(df)
    escolhidas = _preparar_plano(m, secoes)
    recorte = descrever_filtros(filtros)

    # todo slide abaixo leva o mesmo recorte no rodapé; o número da página é
    # preenchido no fim, quando já se sabe quais slides entram
    def _html_cabecalho_rodape(tema, conteudo):
        return _html_cabecalho_rodape_base(tema, "§PAG§", m["total_slides"], conteudo, recorte=recorte)

    total_investimento = m["total_investimento"]

    # --- slide 1 ---
    cards_status = [
        ("CONCLUÍDO / INAUGURADO", m["qtd_concluido"], "2E7D32"),
        ("EM OBRAS / ANDAMENTO", m["qtd_andamento"], "0B5CAD"),
        ("SELEÇÃO / PREPARATÓRIA / LICITAÇÃO", m["qtd_preparatoria"], "E68A00"),
        ("PARALISADO", m["qtd_paralisado"], "C62828"),
    ]
    frac_contratado = (m["total_contratado"] / total_investimento * 100) if total_investimento else 0
    panorama_html = f'''
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
        {"".join(f'<div class="cartao-status"><div class="cs-barra" style="background:#{c}"></div><div class="cs-num" style="color:#{c}">{qtd}</div><div class="cs-rotulo">{_html_esc(rot)}</div></div>' for rot, qtd, c in cards_status if qtd > 0)}
      </div>
    '''
    slide1 = _html_cabecalho_rodape("PANORAMA DA CARTEIRA", panorama_html)

    # --- slide 2 ---
    slide2 = _html_cabecalho_rodape("COMPOSIÇÃO POR ORIGEM DO RECURSO", f'''
      <div class="grafico-origem">{_html_grafico_fase(m, _grafico_total_fontes(m), ROTULO_GRAFICO_ORIGEM, LEGENDA_FASES)}</div>
      <div class="tres-colunas" style="margin-top:24px">
        {"".join(f"""<div class="cartao-destaque">
          <div class="cd-rotulo">{_html_esc(rot)}</div>
          <div class="cd-valor" style="color:#{cor_v}">{_html_esc(abreviar_moeda(valor))}</div>
          <div class="cd-rodape">{_html_esc(pct(valor, total_investimento))} do investimento total</div>
        </div>""" for rot, valor, cor_v in _parcelas(m))}
      </div>
      <div class="nota-rodape">* Financiamento classificado como parcela Estadual por definição adotada nesta
      apresentação — a base não distingue dívida do Estado com aval federal de financiamento federal direto.</div>
    ''')

    # --- slides por fase: gráfico da fase > tabelas de objetos a contratar e contratados ---
    slides_composicao_fase = [_html_cabecalho_rodape(tema, f'''
      <div class="composicao-fase">
        {_html_grafico_fase(m, grafico) if grafico else ""}
        {"".join(_html_tabela(cabecalho, linhas, linhas_bold={len(linhas) - 1} if tem_total else (), cor_cabecalho=cor_bloco) for cabecalho, linhas, tem_total, cor_bloco in blocos)}
        <div class="nota-rodape" style="margin-top:6px">{_html_esc(nota)}</div>
      </div>
    ''') for tema, grafico, blocos, nota in _paginas_composicao_fase(m)]

    # --- slide 4 (novo): status por fase e origem do recurso ---
    blocos_fase = []
    for rotulo, chave, cheia, clara, c_val, a_val in _fontes_por_fase(m):
        linhas_fase = "".join(_html_linha_tabela(rotulo_fase(f), *m["fonte_fase"][chave][f], c_val + a_val, cheia, clara)
                              for f in m["fases_ordem"])
        blocos_fase.append(f'''<div class="bloco-fonte-fase compacto">
          {_html_linha_tabela(rotulo, c_val, a_val, total_investimento, cheia, clara)}
          <div class="nivel-2">{linhas_fase}</div>
        </div>''')
    slide2b = _html_cabecalho_rodape("STATUS POR FASE E ORIGEM DO RECURSO", f'''
      <div class="rotulo-secao" style="display:flex;justify-content:space-between;align-items:baseline">
        <span>INVESTIMENTO POR FONTE E FASE (contratado + a contratar)*</span>
        <span style="font-size:13px;font-weight:400;color:#5B6576;text-transform:none">{_html_esc(LEGENDA_FASES)}</span>
      </div>
      <div class="bloco-barras tabela-fases">{_html_cabecalho_tabela()}{"".join(blocos_fase)}</div>
      <div class="nota-rodape" style="margin-top:4px">{_html_esc(_nota_nao_discriminado(m))}</div>
    ''')

    # --- slides de eixo (novos): tabela paginada, ver _slides_tabela_eixo ---
    def _html_slides_eixo(tema, rotulo_secao, legenda, grupos):
        paginas = _paginar_grupos(grupos)
        saida = []
        for i, pagina in enumerate(paginas, 1):
            blocos = _html_blocos_grupos(pagina, total_investimento)
            saida.append(_html_cabecalho_rodape(_tema_paginado(tema, i, len(paginas)), f'''
              <div class="rotulo-secao" style="display:flex;justify-content:space-between;align-items:baseline">
                <span>{_html_esc(rotulo_secao)}</span>
                <span style="font-size:13px;font-weight:400;color:#5B6576;text-transform:none">{_html_esc(legenda)}</span>
              </div>
              <div class="bloco-barras tabela-fases">{_html_cabecalho_tabela()}{blocos}</div>
              <div class="nota-rodape" style="margin-top:4px">{_html_esc(_nota_eixos(m, grupos))}</div>
            '''))
        return saida

    slides_eixo_fonte = _html_slides_eixo("INVESTIMENTO POR EIXO E FONTE",
                                          "INVESTIMENTO TOTAL, POR EIXO E FONTE (contratado + a contratar)*",
                                          LEGENDA_EIXO_FONTE, _grupos_eixo_fonte(m))
    slides_fonte_eixo = _html_slides_eixo("INVESTIMENTO POR FONTE E EIXO",
                                          "INVESTIMENTO TOTAL, POR FONTE E EIXO (contratado + a contratar)*",
                                          LEGENDA_FASES, _grupos_fonte_eixo(m))

    # --- slide 3 ---
    total_contratado = m["total_contratado"]
    slide3 = _html_cabecalho_rodape("EXECUÇÃO DO CONTRATADO, POR FONTE", f'''
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
    slide4 = _html_cabecalho_rodape("O QUE AINDA PRECISA SER CONTRATADO", f'''
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
    linhas5.append(["SUBTOTAL FEDERAL (OGU)", abreviar_moeda(m["contratado_parcela_federal"]),
                     abreviar_moeda(m["a_contratar_parcela_federal"]), abreviar_moeda(m["total_parcela_federal"]),
                     pct(m["total_parcela_federal"], total_investimento)])
    linhas5.append(["SUBTOTAL ESTADUAL (Financiamento + Contrapartida + Complementar)",
                     abreviar_moeda(m["contratado_parcela_estadual"]),
                     abreviar_moeda(m["a_contratar_parcela_estadual"]), abreviar_moeda(m["total_parcela_estadual"]),
                     pct(m["total_parcela_estadual"], total_investimento)])
    linhas5.append(["TOTAL GERAL", abreviar_moeda(m["total_contratado"]), abreviar_moeda(m["total_a_contratar"]),
                     abreviar_moeda(total_investimento), "100,0%"])
    slide5 = _html_cabecalho_rodape("TABELA-RESUMO POR FONTE", f'''
      {_html_tabela(["FONTE", "CONTRATADO", "A CONTRATAR", "TOTAL", "% DO TOTAL"], linhas5, linhas_bold={4, 5, 6})}
      <div class="nota-rodape">"Não discriminado" = contratado sem detalhamento por fonte preenchido na base{f" (ver slide {_slide_execucao(m)})." if _slide_execucao(m) else "."}</div>
    ''')

    # --- quadro resumo por fase ---
    def _html_quadro(formatar, legenda):
        linhas_quadro = _linhas_quadro_resumo(m, formatar)
        return f'''
      <div class="rotulo-secao" style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:18px">
        <span>{_html_esc(SUBTITULO_QUADRO_RESUMO)}</span>
        <span style="font-size:13px;font-weight:400;color:#5B6576;text-transform:none">{_html_esc(legenda)}</span>
      </div>
      {_html_tabela(CABECALHO_QUADRO_RESUMO, linhas_quadro, linhas_bold={len(linhas_quadro) - 1})}
      <div class="nota-rodape">{_html_esc(_nota_nao_discriminado(m, com_remissao=False))}</div>
    '''

    slide_quadro = _html_cabecalho_rodape(TEMA_QUADRO_RESUMO, _html_quadro(_reais_inteiros, LEGENDA_QUADRO_RESUMO))

    # --- visão geral: quadro em cima, panorama embaixo (versão compacta via CSS) ---
    rotulo_barra = 'CONTRATADO&nbsp;&nbsp;x&nbsp;&nbsp;A CONTRATAR</div>'
    assert panorama_html.count(rotulo_barra) == 1
    legenda_faixa = "".join(f'<i style="background:#{cheia}"></i>{_html_esc(rotulo)}'
                            for rotulo, cheia in _fontes_legenda_faixa(m))
    faixa_fontes = "".join(
        f'<div style="width:{fracao * 100:.4f}%;background:#{fundo};color:#{cor_texto}">{_html_esc(texto)}</div>'
        for texto, fundo, cor_texto, fracao in _segmentos_faixa_fontes(m))
    panorama_com_faixa = panorama_html.replace(
        rotulo_barra,
        f'<span class="legenda-faixa">{legenda_faixa}</span>{rotulo_barra}<div class="faixa-fontes">{faixa_fontes}</div>')
    valor_destaque = abreviar_moeda(m["a_contratar_financiamento"])
    destaque_financiamento = "" if valor_destaque == "-" else (
        f'<div class="destaque-financiamento"><div class="rotulo-secao">{_html_esc(ROTULO_DESTAQUE_FINANCIAMENTO)}</div>'
        f'<div class="valor-principal" style="color:#{COR_FINANCIAMENTO}">{_html_esc(valor_destaque)}</div></div>')
    slide_visao_geral_html = _html_cabecalho_rodape(
        "VISÃO GERAL DA CARTEIRA",
        f'<div class="visao-geral">{_html_quadro(_mi_bi_2_casas, LEGENDA_QUADRO_ABREVIADA)}'
        f'{destaque_financiamento}{panorama_com_faixa}</div>')

    # --- slide 6 ---
    linhas6 = []
    for eixo, linha in m["por_eixo"].iterrows():
        linhas6.append([eixo, int(linha["QTD"]), abreviar_moeda(linha["INVESTIMENTO"]),
                         abreviar_moeda(linha["CONTRATADO"]), abreviar_moeda(linha["A_CONTRATAR"])])
    slide6 = _html_cabecalho_rodape("DETALHE POR EIXO DO PAC",
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
    slide7 = _html_cabecalho_rodape("MAIORES AÇÕES A CONTRATAR",
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
    slide8 = _html_cabecalho_rodape("PENDÊNCIAS QUE CONDICIONAM A CONTRATAÇÃO", f'''
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
        (abreviar_moeda(m["a_contratar_parcela_federal"]), "desse valor dependem de fonte Federal (OGU)"),
        (f"{m['gestao_federal_qtd']}", "ações sob execução direta de órgãos federais na carteira"),
        (abreviar_moeda(m["valor_clausula_vigente"]),
         f"já contratados têm cláusula suspensiva vigente ({m['qtd_clausula_vigente']} ações) — recurso condicionado"),
    ]
    slide9 = _html_cabecalho_rodape("SÍNTESE E PLEITO", f'''
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
    .barra-valor { width:220px; text-align:right; font-weight:700; color:#0B1F3F; font-size:17px; flex-shrink:0; }
    /* tabelas de status por fase e de eixo: rótulo | CONTRATADO | barra | A CONTRATAR | TOTAL,
       com cabeçalho único; 4 grupos x 4 linhas desde o Não discriminado */
    .bloco-fonte-fase.compacto { margin-bottom:22px; }
    .bloco-fonte-fase.compacto.denso { margin-bottom:12px; }
    .compacto > .barra-linha { margin-bottom:8px; }
    .compacto > .barra-linha .barra-trilho { height:24px; }
    .compacto .nivel-2 .barra-linha { margin-bottom:3px; }
    .compacto .nivel-2 .barra-trilho { height:15px; }
    .grafico-fase-linha { display:grid; grid-template-columns:400px 1fr 195px 185px 115px; column-gap:14px;
      align-items:center; color:#5B6576; font-size:14px; margin-bottom:2px; }
    .grafico-fase-linha .barra-trilho { flex:none; height:13px; }
    .grafico-fase-linha.nivel-1 { font-size:17px; margin-bottom:4px; }
    .grafico-fase-linha.nivel-1 .barra-trilho { height:20px; }
    .gf-rotulo { font-weight:700; color:#0B1F3F; font-size:14px; padding-left:28px; }
    .nivel-1 .gf-rotulo { font-size:19px; padding-left:0; }
    .gf-total { font-weight:700; color:#0B1F3F; }
    .grafico-origem .grafico-fase-linha { font-size:17px; margin-bottom:12px; }
    .grafico-origem .grafico-fase-linha .barra-trilho { height:30px; }
    .grafico-origem .nivel-1 { font-size:19px; }
    .grafico-origem .nivel-1 .barra-trilho { height:36px; }
    .grafico-origem .gf-rotulo { font-size:18px; }
    .grafico-origem .nivel-1 .gf-rotulo { font-size:22px; }
    .grafico-origem .gf-cabecalho { font-size:12px; margin:4px 0 12px; }
    .gf-cabecalho { font-size:12px; font-weight:700; border-bottom:1px solid #D8DCE3; padding-bottom:3px; margin:4px 0 4px; }
    .composicao-fase table.pac-tabela { font-size:15px; margin-top:14px; table-layout:fixed; }
    .composicao-fase table.pac-tabela th { padding:3px 14px; font-size:14px; }
    .composicao-fase table.pac-tabela td { padding:2px 14px; }
    .composicao-fase table.pac-tabela th:first-child { width:26%; }
    .visao-geral .rotulo-secao { margin:0 0 8px !important; font-size:16px; }
    .visao-geral .nota-rodape { margin:8px 0 14px; font-size:12px; }
    .visao-geral table.pac-tabela { font-size:15px; margin-top:0; }
    .visao-geral table.pac-tabela th { padding:7px 14px; font-size:14px; }
    .visao-geral table.pac-tabela td { padding:5px 14px; }
    .visao-geral .valor-principal { font-size:44px; margin-top:0; line-height:1.1; }
    .visao-geral .texto-secundario { font-size:14px; margin:2px 0 10px; }
    .visao-geral .barra-empilhada { height:64px; margin-top:2px; }
    .destaque-financiamento { float:right; text-align:right; }
    .faixa-fontes { display:flex; height:32px; clear:both; }
    .faixa-fontes > div { display:flex; align-items:center; padding-left:6px; font-size:13px; font-weight:700;
      white-space:nowrap; overflow:hidden; }
    .legenda-faixa { float:right; font-size:13px; font-weight:400; color:#5B6576; text-transform:none; }
    .legenda-faixa i { display:inline-block; width:11px; height:11px; margin:0 5px 0 14px; }
    .visao-geral .barra-empilhada-seg { padding:8px 16px; }
    .visao-geral .be-titulo { font-size:16px; }
    .visao-geral .be-sub { font-size:14px; margin-top:2px; }
    .visao-geral .cartoes-grade { margin-top:12px; }
    .visao-geral .cs-barra { height:6px; }
    .visao-geral .cs-num { font-size:28px; padding:2px 16px 0; }
    .visao-geral .cs-rotulo { font-size:13px; padding:0 16px 5px; }
    .barra-col-valor { width:130px; flex-shrink:0; font-size:15px; color:#5B6576; }
    .col-contratado { text-align:right; }
    .col-a-contratar { font-weight:700; color:#0B1F3F; }
    .cabecalho-tabela { border-bottom:1px solid #D8DCE3; padding-bottom:4px; margin-bottom:10px; }
    .cabecalho-tabela .barra-col-valor { font-size:12px; font-weight:700; color:#5B6576; }
    .nivel-2 .barra-col-valor { font-size:14px; }
    .nivel-2 .barra-rotulo { font-size:14px; padding-left:28px; white-space:nowrap; }
    .duas-colunas { display:grid; grid-template-columns:1fr 1fr; gap:24px; }
    .tres-colunas { display:grid; grid-template-columns:1fr 1fr 1fr; gap:24px; }
    .tres-colunas .cd-rotulo { min-height:2.6em; }  /* valores alinhados mesmo se um rótulo quebra */
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

    /* Impressão: um slide por folha A4 paisagem. Os tamanhos são em px
       fixos pensados para 1480px de largura, então em vez de deixar o
       max-width:100% espremer (e cortar) o conteúdo na largura do papel,
       o slide mantém os 1480px e é reduzido por inteiro com zoom
       (277mm úteis ≈ 1047px → 1047/1480 ≈ 0,707). */
    @page { size: A4 landscape; margin: 10mm; }
    @media print {
      body { padding:0; background:#fff; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
      .slide { width:1480px; max-width:none; zoom:0.705; margin:0; box-shadow:none; border-radius:0;
        break-after:page; page-break-after:always; break-inside:avoid; page-break-inside:avoid; }
      .slide:last-of-type { break-after:auto; page-break-after:auto; }
    }
    '''

    # só os slides escolhidos, numerados na ordem
    por_chave = {"PANORAMA": [slide1], "VISAO_GERAL": [slide_visao_geral_html], "COMP_ORIGEM": [slide2], "COMP_FASE": slides_composicao_fase,
                 "STATUS_FASE": [slide2b],
                 "EIXO_FONTE": slides_eixo_fonte, "FONTE_EIXO": slides_fonte_eixo, "EXECUCAO": [slide3],
                 "A_CONTRATAR": [slide4], "TABELA_RESUMO": [slide5], "QUADRO_FASE": [slide_quadro],
                 "DETALHE_EIXO": [slide6],
                 "MAIORES_ACOES": [slide7], "PENDENCIAS": [slide8], "SINTESE": [slide9]}
    corpo = "".join(html_slide.replace("§PAG§", str(numero)) for numero, html_slide in enumerate(
        (h for chave in escolhidas for h in por_chave[chave]), 1))

    return f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Análise Financeira do PAC - Apresentação</title>
<style>{css}</style>
</head>
<body>
{corpo}
</body>
</html>'''


def gerar_html_bytes(df, filtros=None, secoes=None):
    # Mesmo padrão de gerar_pptx_bytes: monta tudo em memória, sem tocar o
    # disco — usado pelo botão "ANÁLISE FINANCEIRA (HTML)" do painel.
    html_texto = montar_html_apresentacao(df, filtros, secoes)
    return html_texto.encode("utf-8")


# ---------------------------------------------------------------------------
# Catálogo de slides, na ordem do deck — alimenta a janela de seleção do painel
# (mesma do relatório gerencial; "layout" escolhe o desenho da miniatura).
# ---------------------------------------------------------------------------
SLIDES_APRESENTACAO = [
    {"chave": "PANORAMA", "titulo": "Panorama da Carteira", "layout": "cards_tabela"},
    {"chave": "VISAO_GERAL", "titulo": "Visão Geral da Carteira", "layout": "cards_tabela"},
    {"chave": "COMP_ORIGEM", "titulo": "Composição por Origem do Recurso", "layout": "barras"},
    {"chave": "COMP_FASE", "titulo": "Composição por Fase", "layout": "barras"},
    {"chave": "STATUS_FASE", "titulo": "Status por Fase e Origem do Recurso", "layout": "barras"},
    {"chave": "EIXO_FONTE", "titulo": "Investimento por Eixo e Fonte", "layout": "barras"},
    {"chave": "FONTE_EIXO", "titulo": "Investimento por Fonte e Eixo", "layout": "barras"},
    {"chave": "EXECUCAO", "titulo": "Execução do Contratado, por Fonte", "layout": "barras"},
    {"chave": "A_CONTRATAR", "titulo": "O que Ainda Precisa ser Contratado", "layout": "barras"},
    {"chave": "TABELA_RESUMO", "titulo": "Tabela-Resumo por Fonte", "layout": "tabela"},
    {"chave": "QUADRO_FASE", "titulo": "Quadro Resumo de Investimentos por Fase", "layout": "tabela"},
    {"chave": "DETALHE_EIXO", "titulo": "Detalhe por Eixo do PAC", "layout": "tabela"},
    {"chave": "MAIORES_ACOES", "titulo": "Maiores Ações a Contratar", "layout": "tabela"},
    {"chave": "PENDENCIAS", "titulo": "Pendências que Condicionam a Contratação", "layout": "cards_tabela"},
    {"chave": "SINTESE", "titulo": "Síntese e Pleito", "layout": "texto"},
]


def _paginas_por_chave(m):
    return {"COMP_FASE": len(_paginas_composicao_fase(m)),
            "EIXO_FONTE": len(_paginar_grupos(_grupos_eixo_fonte(m))),
            "FONTE_EIXO": len(_paginar_grupos(_grupos_fonte_eixo(m)))}


def _preparar_plano(m, secoes):
    # Decide quais slides entram (None/vazio = todos; chave desconhecida é
    # ignorada) e guarda em m a numeração, usada nas remissões e nos rodapés.
    pedidas = set(secoes or [])
    escolhidas = [s["chave"] for s in SLIDES_APRESENTACAO if s["chave"] in pedidas]
    escolhidas = escolhidas or [s["chave"] for s in SLIDES_APRESENTACAO]
    paginas = _paginas_por_chave(m)
    m["num_slide"], proximo = {}, 1
    for chave in escolhidas:
        m["num_slide"][chave] = proximo
        proximo += paginas.get(chave, 1)
    m["total_slides"] = proximo - 1
    return escolhidas


def listar_slides(df):
    paginas = _paginas_por_chave(calcular_metricas(df))
    return [dict(s, paginas=paginas.get(s["chave"], 1), marcada=True) for s in SLIDES_APRESENTACAO]


# ---------------------------------------------------------------------------
# API pública — usada tanto pelo botão do painel quanto pelo main() abaixo.
# Exige que configurar(pac) já tenha sido chamado antes.
# ---------------------------------------------------------------------------
def montar_apresentacao(df, filtros=None, secoes=None):
    m = calcular_metricas(df)
    escolhidas = _preparar_plano(m, secoes)

    prs = Presentation()
    prs._recorte_pac = descrever_filtros(filtros)  # lido por novo_slide (rodapé)
    prs._pagina_atual, prs._total_paginas = 0, m["total_slides"]  # idem (numeração)
    prs.slide_width = pol(LARGURA_IN)
    prs.slide_height = pol(ALTURA_IN)

    desenhar = {"PANORAMA": slide_1_panorama, "VISAO_GERAL": slide_visao_geral, "COMP_ORIGEM": slide_2_composicao_origem,
                "COMP_FASE": slide_composicao_fase, "STATUS_FASE": slide_status_fase_origem,
                "EIXO_FONTE": slide_eixo_fonte, "FONTE_EIXO": slide_fonte_eixo,
                "EXECUCAO": slide_3_execucao_contratado, "A_CONTRATAR": slide_4_a_contratar,
                "TABELA_RESUMO": slide_5_tabela_resumo_fonte, "QUADRO_FASE": slide_quadro_resumo_fase,
                "DETALHE_EIXO": slide_6_por_eixo,
                "MAIORES_ACOES": slide_7_top10, "PENDENCIAS": slide_8_pendencias, "SINTESE": slide_9_sintese}
    for chave in escolhidas:
        desenhar[chave](prs, m)

    return prs, m


def gerar_pptx_bytes(df, filtros=None, secoes=None):
    # Usado pelo botão "ANÁLISE FINANCEIRA (PPTX)" do painel (ver
    # _construir_pptx_analise_financeira_bytes em "CGAPE - BALANÇO PAC.py") —
    # monta tudo em memória, sem tocar o disco, igual ao padrão já usado para
    # o PDF (_api_gerar_pdf_bytes).
    import io
    prs, m = montar_apresentacao(df, filtros, secoes)
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
