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
import base64
import importlib.util
import io
import os
import re
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
# No .exe (PyInstaller --onefile), __file__ aponta pra pasta TEMPORÁRIA onde
# o executável se descompacta, não pra pasta dele — a "APRESENTAÇÃO - CGAPE"
# (páginas do Canva) e o resto dos arquivos de apoio ficam AO LADO do .exe,
# mesmo critério do PASTA_BASE do script principal. Sem isso, o .exe não
# achava as páginas do Canva (bug corrigido em 09/11/2026).
if getattr(sys, "frozen", False):
    PASTA_BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
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

# Link do design/mockup no Canva — mostrado no rodapé de cada slide do HTML,
# à esquerda da numeração de página (pedido do usuário em 08/11/2026).
LINK_CANVA = "https://sea-lion-m85670.my.canva.site/"

# Páginas extras do HTML com as páginas do site do Canva exportadas como
# imagem (o Canva bloqueia ser embutido via iframe — tentativa anterior
# abandonada em 08/11/2026), uma por arquivo nesta pasta, pra poder
# intercalar com as seções do dash no painel de seleção/ordenação. O nome
# do arquivo é o número da página no Canva (ex.: "12.jpg"), não
# necessariamente sequencial, pois nem toda página precisa ser exportada.
# Não existe no catálogo do PPTX (SLIDES_APRESENTACAO) — pedido do usuário
# em 08/11/2026.
PASTA_SLIDES_CANVA = os.path.join(PASTA_BASE, "APRESENTAÇÃO - CGAPE")
_RE_ARQUIVO_SLIDE_CANVA = re.compile(r"^(\d+)\.(jpe?g|png)$", re.IGNORECASE)


def _chave_canva(numero):
    return f"CANVA_{numero}"


def _paginas_canva_disponiveis():
    if not os.path.isdir(PASTA_SLIDES_CANVA):
        return []
    paginas = []
    for nome in os.listdir(PASTA_SLIDES_CANVA):
        m = _RE_ARQUIVO_SLIDE_CANVA.match(nome)
        if m:
            paginas.append((int(m.group(1)), nome))
    paginas.sort()
    return paginas


def _data_uri_imagem(caminho):
    # Reduz/recomprime antes de embutir em base64 — o Canva exporta em
    # qualidade alta (1920x1080, até ~1,8 MB por página), o que deixaria o
    # HTML pesado demais pra compartilhar; se o Pillow não conseguir abrir
    # por algum motivo, usa o arquivo original sem mexer.
    with open(caminho, "rb") as f:
        dados_originais = f.read()
    mime = "image/png" if caminho.lower().endswith(".png") else "image/jpeg"
    dados = dados_originais
    try:
        from PIL import Image
        imagem = Image.open(io.BytesIO(dados_originais))
        if imagem.width > 1600:
            nova_altura = round(imagem.height * 1600 / imagem.width)
            imagem = imagem.resize((1600, nova_altura), Image.LANCZOS)
        buffer = io.BytesIO()
        imagem.convert("RGB").save(buffer, format="JPEG", quality=82, optimize=True)
        dados = buffer.getvalue()
        mime = "image/jpeg"
    except Exception:
        pass
    return f"data:{mime};base64,{base64.b64encode(dados).decode('ascii')}"
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
def texto_base_atualizada():
    # Linha "Base atualizada em" logo abaixo de "COORDENAÇÃO..." no cabeçalho
    # do PPTX e do HTML (pedido do usuário em 09/11/2026). Data/hora de
    # pac.ultima_atualizacao: o "Modificado em" que o Excel grava na própria
    # planilha ao salvar (não a data do arquivo no disco) — a mesma usada no PDF.
    data = getattr(pac, "ultima_atualizacao", None)
    if data is None:
        return ""
    return f"Base atualizada em: {data.strftime('%d/%m/%Y')} às {data.strftime('%Hh%M')}"


def novo_slide(prs, tema):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # layout em branco
    # numeração sequencial: depende de quais slides foram escolhidos no painel
    prs._pagina_atual += 1
    pagina, total_paginas = prs._pagina_atual, prs._total_paginas

    # fundo branco explícito
    add_rect(slide, 0, 0, LARGURA_IN, ALTURA_IN, fill=BRANCO)

    # faixa do cabeçalho
    add_rect(slide, 0, 0, LARGURA_IN, 1.1, fill=NAVY)
    # Título e subtítulo subiram um pouco pra caber a linha "Base atualizada
    # em" dentro da mesma faixa de 1,1" — o conteúdo dos slides não se mexe.
    add_text(slide, 0.56, 0.1, 12.5, 0.5, f"PAC - {tema}", 26, BRANCO, negrito=True,
              ancorar=MSO_ANCHOR.MIDDLE)
    add_text(slide, 0.56, 0.6, 12.5, 0.24, "COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS ESTRATÉGICAS",
              12, AZUL_CLARO_TEXTO, negrito=True)
    base_atualizada = texto_base_atualizada()
    if base_atualizada:
        add_text(slide, 0.56, 0.84, 12.5, 0.2, base_atualizada, 10, AZUL_CLARO_TEXTO)
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
    # Regra única do Termo de Compromisso (decisão do usuário em 09/11/2026):
    # T.C. FIRMADO x NÃO FIRMADO vem só de SINALIZACAO_TC (coluna do TC/CT
    # FINANCIAMENTO preenchida = firmado; vazia, "-" ou "0" = não firmado — ver
    # pac.classificar_termo_compromisso). Toda a apresentação chama o lado
    # "contratado" de T.C. FIRMADO e o "a contratar" de T.C. NÃO FIRMADO, então
    # o valor de cada ação vai INTEIRO pro lado da sua sinalização. Ação com TC
    # firmado mas VALOR CONTRATADO zerado na planilha entra no lado firmado com
    # o valor planejado de cada fonte (o mesmo que o painel já usa como valor
    # dela) — os totais não mudam, só o lado em que o valor aparece.
    tc_firmado = df["SINALIZACAO_TC"].eq("SIM")

    def _firmado(serie):
        return serie.where(tc_firmado, 0.0)

    def _nao_firmado(serie):
        return serie.where(~tc_firmado, 0.0)

    apoiado_plan = df[pac.col_apoiado_ajustado]
    financiamento_plan = df[pac.col_financiamento_ajustado]
    contrapartida_plan = df[pac.col_contrapartida_ajustado]
    complementar_plan = df[pac.col_complementar_ajustado]
    plan_linha = apoiado_plan + financiamento_plan + contrapartida_plan + complementar_plan

    contratado = df[pac.col_valor_contratado] + _firmado(plan_linha)
    contratado_ogu_linha = df[pac.col_apoiado_contratado] + _firmado(apoiado_plan)
    contratado_fin_linha = df[pac.col_financiamento_contratado] + _firmado(financiamento_plan)
    contratado_est_linha = (df[pac.col_contrapartida_contratado] + df[pac.col_complementar_contratado]
                            + _firmado(contrapartida_plan + complementar_plan))
    apoiado_aj = _nao_firmado(apoiado_plan)
    financiamento_aj = _nao_firmado(financiamento_plan)
    contrapartida_aj = _nao_firmado(contrapartida_plan)
    complementar_aj = _nao_firmado(complementar_plan)
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
    m["contratado_ogu"] = float(contratado_ogu_linha.sum())
    m["contratado_financiamento"] = float(contratado_fin_linha.sum())
    m["contratado_estadual"] = float(contratado_est_linha.sum())
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
        "ogu": (contratado_ogu_linha, apoiado_aj),
        "financiamento": (contratado_fin_linha, financiamento_aj),
        "estadual": (contratado_est_linha, contrapartida_aj + complementar_aj),
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
    _calcular_clausulas(df, m, investimento_linha, eixo, hoje)
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
# Cláusulas suspensivas — seção própria do HTML, em cards agrupados por eixo
# (pedido do usuário em 10/11/2026). Entra em card toda ação cuja cláusula
# ainda não foi totalmente retirada (Cláusula Vigente, Retirada Parcial ou
# qualquer outra situação nova que a planilha venha a ter); Retirada Total e
# NÃO DEFINIDA só aparecem na contagem do topo.
# ---------------------------------------------------------------------------
SITUACOES_CLAUSULA_SEM_CARD = {"RETIRADA TOTAL", "NÃO DEFINIDA"}
DIAS_ALERTA_VIGENCIA = 90


def _texto_celula(valor):
    if valor is None:
        return ""
    try:
        import pandas as pd
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    return "" if texto.upper() in ("NAN", "NAT", "-") else texto


def _data_celula(valor):
    try:
        import pandas as pd
        data = pd.to_datetime(valor, errors="coerce", dayfirst=True)
        return None if pd.isna(data) else data.date()
    except (TypeError, ValueError):
        return None


def _calcular_clausulas(df, m, investimento_linha, eixo, hoje):
    situacao = df["SITUACAO_CLAUSULA_SUSPENSIVA"].fillna("NÃO DEFINIDA").astype(str)
    situacao_up = situacao.str.upper()
    m["qtd_clausula_retirada_total"] = int((situacao_up == "RETIRADA TOTAL").sum())
    m["qtd_clausula_nao_definida"] = int((situacao_up == "NÃO DEFINIDA").sum())

    acoes = []
    for idx in df.index[~situacao_up.isin(SITUACOES_CLAUSULA_SEM_CARD)]:
        linha = df.loc[idx]
        vigencia = _data_celula(linha.get(pac.col_vigencia))
        dias = (vigencia - hoje).days if vigencia else None
        if dias is None:
            prazo = "sem_data"
        elif dias < 0:
            prazo = "vencida"
        elif dias <= DIAS_ALERTA_VIGENCIA:
            prazo = "proxima"
        else:
            prazo = "no_prazo"
        acoes.append({
            "eixo": eixo.loc[idx],
            "objeto": _texto_celula(linha.get(pac.col_objeto)) or "NÃO INFORMADO",
            "municipio": _texto_celula(linha.get(pac.col_municipio)),
            "situacao": situacao.loc[idx],
            "status": _texto_celula(linha.get("STATUS_TEXTO")),
            "fase": _texto_celula(linha.get("FASE_TEXTO")),
            "orgao": _texto_celula(linha.get(pac.col_orgao)),
            "motivo": _texto_celula(linha.get(pac.col_motivo_clausula_suspensiva)),
            "valor": float(investimento_linha.loc[idx]),
            "vigencia": vigencia,
            "dias": dias,
            "prazo": prazo,
        })

    grupos = {}
    for acao in acoes:
        grupos.setdefault(acao["eixo"], []).append(acao)
    for lista in grupos.values():
        lista.sort(key=lambda a: (a["vigencia"] is None, a["vigencia"] or hoje, -a["valor"]))
    m["clausulas_por_eixo"] = sorted(grupos.items(), key=lambda g: (-sum(a["valor"] for a in g[1]), g[0]))
    m["clausulas_qtd"] = len(acoes)
    m["clausulas_valor"] = float(sum(a["valor"] for a in acoes))
    m["clausulas_prazo"] = {k: sum(1 for a in acoes if a["prazo"] == k)
                            for k in ("vencida", "proxima", "no_prazo", "sem_data")}


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------
# Topo do slide Visão Geral (print do usuário em 06/11/2026): investimento total
# à esquerda; à direita, Recursos a Contratar aberto por fonte. A barra passa a
# falar em Termo de Compromisso, que é o que define contratado x a contratar.
ROTULO_RECURSOS_A_CONTRATAR = "TERMO DE COMPROMISSO NÃO FIRMADO"
X_RECURSOS_A_CONTRATAR = 13.2
LARG_RECURSOS_A_CONTRATAR = 5.3
ALT_LINHA_RECURSOS = 0.42
ALT_TOPO_VISAO_GERAL = 2.55
LARG_BARRA_VISAO_GERAL = 17.6


def _linhas_recursos_a_contratar(m):
    # Na ordem do print; fonte sem valor no recorte não aparece
    linhas = []
    for valor, rotulo in ((m["a_contratar_financiamento"], "Financiamento"), (m["a_contratar_ogu"], "OGU"),
                          (m["a_contratar_estadual"], "Recurso Estadual")):
        if abreviar_moeda(valor) != "-":
            linhas.append(f"{abreviar_moeda(valor)} · {rotulo}")
    return linhas


def _desenhar_recursos_a_contratar(s, m, y, tam_valor, alt_valor):
    if abreviar_moeda(m["total_a_contratar"]) == "-":
        return
    x, largura = X_RECURSOS_A_CONTRATAR, LARG_RECURSOS_A_CONTRATAR
    add_text(s, x, y, largura, 0.5, ROTULO_RECURSOS_A_CONTRATAR, 18, CINZA_TEXTO, negrito=True,
              alinhar=PP_ALIGN.RIGHT)
    add_text(s, x, y + 0.35, largura, alt_valor + 0.1, abreviar_moeda(m["total_a_contratar"]), tam_valor, LARANJA,
              negrito=True, alinhar=PP_ALIGN.RIGHT)
    y_linha = y + 0.35 + alt_valor
    for linha in _linhas_recursos_a_contratar(m):
        add_text(s, x, y_linha, largura, ALT_LINHA_RECURSOS, linha, 22, LARANJA, negrito=True,
                  alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
        y_linha += ALT_LINHA_RECURSOS


def _texto_barra_termo(longo, curto, valor, largura_pol):
    # Texto do segmento da barra; encurta se não couber (recortes com poucos
    # contratos) e some se o valor for zero. A mesma decisão vale pro PPTX e o HTML.
    moeda = abreviar_moeda(valor)
    if moeda == "-":
        return ""
    for texto in (f"{longo} · {moeda}", f"{curto} · {moeda}", moeda):
        if len(texto) * 0.135 + 0.3 <= largura_pol:
            return texto
    return ""


def _barra_termo_compromisso(m):
    # ((texto, percentual, fração da barra) do firmado, (idem) do não firmado)
    total = m["total_investimento"]
    fc = m["total_contratado"] / total if total else 0
    return (
        (_texto_barra_termo("TERMO DE COMPROMISSO FIRMADO", "T.C. FIRMADO", m["total_contratado"], LARG_BARRA_VISAO_GERAL * fc),
         pct(m["total_contratado"], total), fc),
        (_texto_barra_termo("TERMO DE COMPROMISSO NÃO FIRMADO", "T.C. NÃO FIRMADO", m["total_a_contratar"],
                            LARG_BARRA_VISAO_GERAL * (1 - fc)),
         pct(m["total_a_contratar"], total), 1 - fc),
    )


def _desenhar_panorama(s, m, y, compacto=False):
    # Investimento total, barra contratado x a contratar e cards de status. A
    # versão compacta é a do slide Visão Geral, abaixo do quadro por fase, com
    # o bloco Recursos a Contratar e a barra em termos de Termo de Compromisso.
    tam_valor, alt_valor = (44, 0.85) if compacto else (60, 1.2)
    add_text(s, 0.9, y, 10, 0.5, "INVESTIMENTO TOTAL DA CARTEIRA", 18 if compacto else 16, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, y + 0.35, 10, alt_valor + 0.1, abreviar_moeda(m["total_investimento"]), tam_valor, VERDE,
              negrito=True)
    if compacto:
        _desenhar_recursos_a_contratar(s, m, y, tam_valor, alt_valor)
        y_barra = y + ALT_TOPO_VISAO_GERAL
    else:
        y_texto = y + 0.35 + alt_valor
        add_text(s, 0.9, y_texto, 10, 0.4, f"{m['qtd_acoes']} ações do PAC acompanhadas pelo Estado da Bahia", 14,
                  CINZA_TEXTO)
        # barra empilhada contratado x a contratar
        y_barra = y_texto + 0.85
        add_text(s, 0.9, y_barra - 0.45, 6, 0.35, "T.C. FIRMADO  x  T.C. NÃO FIRMADO", 17, NAVY, negrito=True)
    largura_barra = LARG_BARRA_VISAO_GERAL if compacto else 17.0
    x0 = 0.9
    frac_contratado = m["total_contratado"] / m["total_investimento"] if m["total_investimento"] else 0
    largura_1 = largura_barra * frac_contratado
    add_rect(s, x0, y_barra, largura_1, 0.9, fill=VERDE)
    add_rect(s, x0 + largura_1, y_barra, largura_barra * (1 - frac_contratado), 0.9, fill=LARANJA)
    if compacto:
        (texto_1, pct_1, _), (texto_2, pct_2, frac_2) = _barra_termo_compromisso(m)
        larg_texto_1, larg_texto_2 = largura_1 - 0.2, largura_barra * frac_2 - 0.2
    else:
        texto_1, pct_1 = f"T.C. FIRMADO · {abreviar_moeda(m['total_contratado'])}", pct(m["total_contratado"], m["total_investimento"])
        texto_2, pct_2 = f"T.C. NÃO FIRMADO · {abreviar_moeda(m['total_a_contratar'])}", pct(m["total_a_contratar"], m["total_investimento"])
        larg_texto_1, larg_texto_2 = 6, 7
    for x_texto, largura_texto, texto, percentual in ((x0 + 0.15, larg_texto_1, texto_1, pct_1),
                                                      (x0 + largura_1 + 0.15, larg_texto_2, texto_2, pct_2)):
        if texto:
            add_text(s, x_texto, y_barra + 0.1, largura_texto, 0.35, texto, 17, BRANCO, negrito=True)
            add_text(s, x_texto, y_barra + 0.5, largura_texto, 0.3, percentual, 15, BRANCO)

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
LEGENDA_FASES = "tom forte = T.C. FIRMADO · tom claro = T.C. NÃO FIRMADO"


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
        ("NÃO DISCRIMINADO (T.C. firmado sem fonte informada)", "nao_discriminado", COR_NAO_DISCRIMINADO,
         pac.clarear_hex(COR_NAO_DISCRIMINADO).lstrip("#"), nd, 0.0),
    ]


def _parcelas(m):
    # Cards do slide 2: Federal x Estadual x Não discriminado (pedido do
    # usuário em 03/11/2026) — os três somam o investimento total. Card
    # zerado no recorte não aparece (pedido do usuário em 09/11/2026).
    nd = next(f[4] for f in _fontes_por_fase(m) if f[1] == "nao_discriminado")
    cards = [
        ("PARCELA FEDERAL (OGU)", m["total_parcela_federal"], AZUL),
        ("PARCELA ESTADUAL (Financiamento + Contrapartida + Complementar)*", m["total_parcela_estadual"], LARANJA),
        ("NÃO DISCRIMINADO (T.C. firmado sem fonte informada)", nd, COR_NAO_DISCRIMINADO),
    ]
    return [c for c in cards if abs(c[1]) >= 0.5]


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
    # Sem Não discriminado no recorte, a categoria some dos detalhamentos e a
    # nota que a explica também (pedido do usuário em 09/11/2026)
    if not any(abs(c) >= 0.5 for c, _ in m["fonte_fase"]["nao_discriminado"].values()):
        return ""
    return ("* Não discriminado: valor com T.C. firmado sem detalhamento por fonte de recurso na base, "
            "mostrado como categoria própria para a composição somar o investimento total."
            + (f" Ver slide {_slide_execucao(m)}." if com_remissao and _slide_execucao(m) else ""))


def _cabecalho_tabela(s, y):
    add_text(s, X_COL_CONTRATADO, y, LARG_COL_CONTRATADO, 0.28, "T.C. FIRMADO", 11, CINZA_TEXTO, negrito=True,
              alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
    add_text(s, X_COL_A_CONTRATAR, y, LARG_COL_A_CONTRATAR, 0.28, "T.C. NÃO FIRMADO", 11, CINZA_TEXTO, negrito=True,
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
LEGENDA_COMPOSICAO_FASE = "cinza escuro / tom forte = T.C. FIRMADO · cinza claro / tom claro = T.C. NÃO FIRMADO"
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
    add_text(s, 0.9, y, 10.6, 0.4, "INVESTIMENTO POR FONTE E FASE (T.C. firmado + T.C. não firmado)*", 17, CINZA_TEXTO,
              negrito=True)
    add_text(s, 9.9, y, 8.6, 0.4, LEGENDA_FASES, 12, CINZA_TEXTO, alinhar=PP_ALIGN.RIGHT,
              ancorar=MSO_ANCHOR.MIDDLE)
    y = _cabecalho_tabela(s, y + 0.5)
    for rotulo, chave, cheia, clara, c_val, a_val in _fontes_por_fase(m):
        if not _linha_com_valor(c_val, a_val):
            continue
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


# Fonte de recurso totalmente zerada num detalhamento some dele — linha de
# gráfico, barra, card ou coluna de tabela — em vez de aparecer só com "-"
# (pedido do usuário em 09/11/2026). Vale pro PPTX e pro HTML.
def _sem_colunas_zeradas(cabecalho, linhas, larguras=None):
    # Tira as colunas do meio (a 1ª é o rótulo, a última o TOTAL) em que toda
    # linha, inclusive a de total, está zerada ("-"). Com `larguras`, a largura
    # das colunas que saem é redistribuída nas que ficam, mesma largura total.
    n = len(cabecalho)
    manter = [i for i in range(n)
              if i in (0, n - 1) or any(str(linha[i]).strip() not in ("-", "") for linha in linhas)]
    cab = [cabecalho[i] for i in manter]
    lin = [[linha[i] for i in manter] for linha in linhas]
    if larguras is None:
        return cab, lin
    fator = sum(larguras) / sum(larguras[i] for i in manter)
    return cab, lin, [larguras[i] * fator for i in manter]


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
    nota = (_nota_nao_discriminado(m) + " Linhas sem valor são omitidas.").strip()
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
        add_text(s, 0.9, y, 10.6, 0.4, rotulo_secao, 17, CINZA_TEXTO, negrito=True)
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
ROTULO_GRAFICO_FASE = "INVESTIMENTO TOTAL, POR FASE E FONTE (T.C. firmado + T.C. não firmado)*"
COLUNAS_FONTES_OBJETOS = ["OGU (APOIADO)", "FINANCIAMENTO", "RECURSO ESTADUAL", "NÃO DISCRIMINADO*", "TOTAL"]
LARGURAS_OBJETOS = [4.6, 2.6, 2.6, 2.6, 2.6, 2.6]
# (título da 1ª coluna, índice em (contratado, a contratar), cor do cabeçalho =
# a da barra Contratado x A Contratar do panorama)
BLOCOS_OBJETOS = [("OBJETO COM T.C. NÃO FIRMADO", 1, LARANJA), ("OBJETO COM T.C. FIRMADO", 0, VERDE)]
ALT_LINHA_OBJETOS = 0.36
ESPACO_BLOCO = 0.25
Y_INICIO_FASE = 1.4
Y_APOS_GRAFICO = 4.05
Y_LIMITE_TABELAS = 10.2   # abaixo disso, só a nota e o rodapé
ROTULO_GRAFICO_ORIGEM = "INVESTIMENTO TOTAL, POR FONTE (T.C. firmado + T.C. não firmado)"
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


def _tabela_objetos(m, fase, titulo, indice):
    # (cabeçalho, linhas) da tabela de objetos, sem as fontes zeradas nela
    return _sem_colunas_zeradas([titulo] + COLUNAS_FONTES_OBJETOS, _linhas_objetos(m, fase, indice))


def _nota_objetos(m, fase, cabecalhos):
    # Nota do Não discriminado só quando a coluna dele ficou em alguma tabela
    if not any("NÃO DISCRIMINADO*" in cab for cab in cabecalhos):
        return ""
    nota = _nota_nao_discriminado(m)
    if any(fontes["nao_discriminado"][0] <= -0.5 for fontes in m["fase_objeto"][fase].values()):
        nota += " Valor negativo: detalhamento por fonte maior que o Valor Contratado na base."
    return nota


def _larguras_objetos(cabecalho):
    # 1ª coluna fixa; o resto da largura (17,6") dividido igual entre as fontes que ficaram + TOTAL
    resto = sum(LARGURAS_OBJETOS) - LARGURAS_OBJETOS[0]
    return [LARGURAS_OBJETOS[0]] + [resto / (len(cabecalho) - 1)] * (len(cabecalho) - 1)


def _paginar_blocos_objetos(blocos):
    # blocos: [(cabeçalho, linhas, cor), ...] (cabeçalho já sem as colunas zeradas)
    # [[(cabeçalho, linhas, tem_total, cor), ...], ...]: uma tabela só começa num slide
    # se couber um pedaço útil, e a linha de TOTAL nunca fica sozinha no seguinte
    paginas = [[]]
    livre = Y_LIMITE_TABELAS - Y_APOS_GRAFICO
    for cabecalho_tabela, linhas, cor_bloco in blocos:
        titulo = cabecalho_tabela[0]
        resto, continuacao = list(linhas), False
        while resto:
            cabem = int(livre / ALT_LINHA_OBJETOS + 1e-9) - 1   # -1 do cabeçalho
            if cabem < min(3, len(resto)):
                paginas.append([])
                livre = Y_LIMITE_TABELAS - Y_INICIO_FASE
                continue
            n = cabem - 1 if len(resto) - cabem == 1 else min(cabem, len(resto))
            parte, resto = resto[:n], resto[n:]
            cabecalho = [titulo + (" (cont.)" if continuacao else "")] + cabecalho_tabela[1:]
            paginas[-1].append((cabecalho, parte, not resto, cor_bloco))
            livre -= ALT_LINHA_OBJETOS * (len(parte) + 1) + ESPACO_BLOCO
            continuacao = True
    return paginas


def _paginas_composicao_fase(m):
    # [(tema do slide, (fase, c, a, fontes) só no 1º slide da fase, blocos, nota), ...]
    saida = []
    for fase, grafico in zip(m["fases_ordem"], _fases_por_fonte(m)):
        tabelas = [(*_tabela_objetos(m, fase, titulo, i), c) for titulo, i, c in BLOCOS_OBJETOS]
        tabelas = [t for t in tabelas if t[1]]
        paginas = _paginar_blocos_objetos(tabelas)
        nota = _nota_objetos(m, fase, [cab for cab, _, _ in tabelas])
        for i, pagina in enumerate(paginas, 1):
            tema = _tema_paginado(f"COMPOSIÇÃO POR FASE - {grafico[0].upper()}", i, len(paginas))
            saida.append((tema, grafico if i == 1 else None, pagina, nota))
    return saida


def _linha_grafico_fase(s, nivel, y, rotulo, c_val, a_val, base, cheia, clara, altura, tam_rotulo, tam_valor,
                        ordem_tc=False):
    x = 0.9 if nivel == 1 else 1.35
    if ordem_tc:
        # Composição por Fase (pedido do usuário em 09/11/2026): T.C. FIRMADO |
        # barra | T.C. NÃO FIRMADO | TOTAL, nas mesmas posições das tabelas de
        # status/eixo (ver _linha_tabela)
        add_text(s, x, y, X_COL_CONTRATADO - 0.1 - x, altura, rotulo, tam_rotulo, NAVY, negrito=True,
                  ancorar=MSO_ANCHOR.MIDDLE)
        add_text(s, X_COL_CONTRATADO, y, LARG_COL_CONTRATADO, altura, abreviar_moeda(c_val), tam_valor, CINZA_TEXTO,
                  alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
        _desenhar_barra_proporcional(s, X_BARRA_TABELA, y + altura * 0.18, LARG_BARRA_TABELA, altura * 0.64,
                                     c_val, a_val, base, cheia, clara)
        add_text(s, X_COL_A_CONTRATAR, y, LARG_COL_A_CONTRATAR, altura, abreviar_moeda(a_val), tam_valor, CINZA_TEXTO,
                  ancorar=MSO_ANCHOR.MIDDLE)
        add_text(s, X_COL_TOTAL, y, LARG_COL_TOTAL, altura, abreviar_moeda(c_val + a_val), tam_valor, NAVY,
                  negrito=True, alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
        return
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


def _desenhar_grafico_fase(s, m, y, grafico, rotulo_secao=ROTULO_GRAFICO_FASE, legenda=None, medidas=MEDIDAS_GRAFICO_FASE,
                           ordem_tc=False):
    # Devolve o y logo abaixo da última linha. ordem_tc=True: colunas T.C.
    # FIRMADO | barra | T.C. NÃO FIRMADO | TOTAL (Composição por Fase)
    rotulo, c_val, a_val, fontes = grafico
    alt_1, tam_1, alt_2, tam_2 = medidas
    if legenda:
        add_text(s, 0.9, y, 10.6, 0.4, rotulo_secao, 17, CINZA_TEXTO, negrito=True)
        add_text(s, 9.9, y, 8.6, 0.4, legenda, 12, CINZA_TEXTO, alinhar=PP_ALIGN.RIGHT, ancorar=MSO_ANCHOR.MIDDLE)
    else:
        add_text(s, 0.9, y, 17.6, 0.4, rotulo_secao, 17, CINZA_TEXTO, negrito=True)
    y += 0.42
    if ordem_tc:
        for x_valor, largura, titulo, alinhar in (
                (X_COL_CONTRATADO, LARG_COL_CONTRATADO, "T.C. FIRMADO", PP_ALIGN.RIGHT),
                (X_COL_A_CONTRATAR, LARG_COL_A_CONTRATAR, "T.C. NÃO FIRMADO", PP_ALIGN.LEFT),
                (X_COL_TOTAL, LARG_COL_TOTAL, "TOTAL", PP_ALIGN.RIGHT)):
            add_text(s, x_valor, y, largura, 0.28, titulo, 11, CINZA_TEXTO, negrito=True, alinhar=alinhar,
                      ancorar=MSO_ANCHOR.MIDDLE)
    else:
        for x_valor, titulo in ((X_GRAFICO_TOTAL, "TOTAL"), (X_GRAFICO_CONTRATADO, "T.C. FIRMADO"),
                                (X_GRAFICO_A_CONTRATAR, "T.C. NÃO FIRMADO")):
            add_text(s, x_valor, y, LARG_GRAFICO_VALOR, 0.28, titulo, 11, CINZA_TEXTO, negrito=True,
                      ancorar=MSO_ANCHOR.MIDDLE)
    add_rect(s, 0.9, y + 0.32, 17.6, 0.012, fill=CINZA_BORDA)
    y += 0.4
    _linha_grafico_fase(s, 1, y, rotulo, c_val, a_val, m["total_investimento"], CINZA_FASE_CONTRATADO,
                        CINZA_FASE_A_CONTRATAR, alt_1, *tam_1, ordem_tc=ordem_tc)
    y += alt_1 + AVANCO_NIVEL_1 - ALT_NIVEL_1
    for rotulo_fonte, c_fonte, a_fonte, cheia, clara in fontes:
        if not _linha_com_valor(c_fonte, a_fonte):
            continue
        _linha_grafico_fase(s, 2, y, rotulo_fonte, c_fonte, a_fonte, c_val + a_val, cheia, clara, alt_2, *tam_2,
                            ordem_tc=ordem_tc)
        y += alt_2
    return y


def slide_composicao_fase(prs, m):
    for tema, grafico, blocos, nota in _paginas_composicao_fase(m):
        s = novo_slide(prs, tema)
        y = Y_INICIO_FASE
        if grafico:
            _desenhar_grafico_fase(s, m, y, grafico, ordem_tc=True)
            y = Y_APOS_GRAFICO
        for cabecalho, linhas, tem_total, cor_bloco in blocos:
            altura = ALT_LINHA_OBJETOS * (len(linhas) + 1)
            tabela_nativa(s, 0.9, y, 17.6, altura, cabecalho, linhas, larguras=_larguras_objetos(cabecalho),
                          tam_fonte=14, linha_bold_indices={len(linhas) - 1} if tem_total else None,
                          cor_cabecalho=cor_bloco)
            y += altura + ESPACO_BLOCO
        if nota:
            add_text(s, 0.9, y - ESPACO_BLOCO + 0.1, 17.6, 0.5, nota, 12, CINZA_TEXTO)


def slide_eixo_fonte(prs, m):
    _slides_tabela_eixo(prs, m, "INVESTIMENTO POR EIXO E FONTE",
                        "INVESTIMENTO TOTAL, POR EIXO E FONTE (T.C. firmado + T.C. não firmado)*", LEGENDA_EIXO_FONTE,
                        _grupos_eixo_fonte(m))


def slide_fonte_eixo(prs, m):
    _slides_tabela_eixo(prs, m, "INVESTIMENTO POR FONTE E EIXO",
                        "INVESTIMENTO TOTAL, POR FONTE E EIXO (T.C. firmado + T.C. não firmado)*", LEGENDA_FASES,
                        _grupos_fonte_eixo(m))


def slide_3_execucao_contratado(prs, m):
    s = novo_slide(prs, "EXECUÇÃO DO T.C. FIRMADO, POR FONTE")
    add_text(s, 0.9, 1.5, 10, 0.4, "TOTAL COM T.C. FIRMADO", 17, CINZA_TEXTO, negrito=True)
    add_text(s, 0.9, 1.85, 9, 1.0, abreviar_moeda(m["total_contratado"]), 44, VERDE, negrito=True)
    add_text(s, 10.2, 2.15, 8.8, 0.7,
              f"Detalhamento por fonte disponível para {pct(m['total_contratado']*m['cobertura_detalhamento_pct'], m['total_contratado'])} "
              "do valor com T.C. firmado.", 15, CINZA_TEXTO)

    y = 3.3
    total_c = m["total_contratado"]
    segmentos = [
        ("APOIADO (OGU) - T.C. FIRMADO", m["contratado_ogu"], COR_OGU),
        ("FINANCIAMENTO - T.C. FIRMADO", m["contratado_financiamento"], COR_FINANCIAMENTO),
        ("RECURSO ESTADUAL - T.C. FIRMADO", m["contratado_estadual"], COR_ESTADUAL),
        ("NÃO DISCRIMINADO", m["contratado_nao_discriminado"], COR_NAO_DISCRIMINADO),
    ]
    for rotulo, valor, c in segmentos:
        if abs(valor) < 0.5:
            continue
        barra_horizontal(s, 0.9, y, 17.6, 0.75, rotulo, valor, total_c, c, largura_rotulo=5.4)
        y += 1.05

    add_pendencia_tag(s, 0.9, y + 0.2, w=2.6)
    add_text(s, 3.7, y + 0.2, 14.0, 0.85,
              f"{m['qtd_sem_detalhamento']} ação sem nenhuma categoria de origem preenchida e mais ações com "
              "divergência entre Valor Contratado e o detalhamento — já sinalizadas no Controle de Qualidade "
              "do painel, aguardando correção de cadastro.", 14, VERMELHO)
    return s


def slide_4_a_contratar(prs, m):
    s = novo_slide(prs, "RECURSOS COM T.C. NÃO FIRMADO")
    add_text(s, 0.9, 1.5, 12, 0.4, "VALOR TOTAL COM T.C. NÃO FIRMADO", 17, CINZA_TEXTO, negrito=True)
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
        if abs(valor) < 0.5:
            continue
        barra_horizontal(s, 0.9, y2, 8.6, 0.7, rotulo, valor, m["total_a_contratar"], c, largura_rotulo=2.9,
                          tam_rotulo=14, largura_valor=2.5, tam_valor=14)
        y2 += 0.95

    add_text(s, 9.9, y, 8.6, 0.35, "POR EIXO (top 5 em valor com T.C. não firmado)", 17, NAVY, negrito=True)
    por_eixo_ord = m["por_eixo"].sort_values("A_CONTRATAR", ascending=False).head(5)
    y3 = y + 0.6
    for eixo, linha in por_eixo_ord.iterrows():
        barra_horizontal(s, 9.9, y3, 8.6, 0.75, eixo, float(linha["A_CONTRATAR"]), m["total_a_contratar"],
                          TEAL, largura_rotulo=3.7, tam_rotulo=13, largura_valor=2.3, tam_valor=13)
        y3 += 1.0
    return s


def slide_5_tabela_resumo_fonte(prs, m):
    s = novo_slide(prs, "TABELA-RESUMO POR FONTE")
    cabecalho = ["FONTE", "T.C. FIRMADO", "T.C. NÃO FIRMADO", "TOTAL", "% DO TOTAL"]
    total_geral = m["total_investimento"]
    linhas_dados = [
        ("OGU (Apoiado)", m["contratado_ogu"], m["a_contratar_ogu"]),
        ("Financiamento", m["contratado_financiamento"], m["a_contratar_financiamento"]),
        ("Recurso Estadual", m["contratado_estadual"], m["a_contratar_estadual"]),
        ("Não discriminado (T.C. firmado)", m["contratado_nao_discriminado"], 0.0),
    ]
    linhas_dados = [d for d in linhas_dados if _linha_com_valor(d[1], d[2])]
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
                  tam_fonte=16, linha_bold_indices={len(linhas) - 3, len(linhas) - 2, len(linhas) - 1})
    if any(d[0].startswith("Não discriminado") for d in linhas_dados):
        add_text(s, 0.9, 7.9, 17, 0.4,
                  "\"Não discriminado\" = T.C. firmado sem detalhamento por fonte preenchido na base"
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
SUBTITULO_QUADRO_RESUMO = "INVESTIMENTO POR FASE E FONTE (T.C. firmado + T.C. não firmado)*"
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
    add_text(s, 0.9, y, 10.6, 0.4, SUBTITULO_QUADRO_RESUMO, 17, CINZA_TEXTO, negrito=True)
    add_text(s, 9.9, y, 8.6, 0.4, legenda, tam_nota, CINZA_TEXTO, alinhar=PP_ALIGN.RIGHT,
              ancorar=MSO_ANCHOR.MIDDLE)
    cabecalho, linhas, larguras = _sem_colunas_zeradas(CABECALHO_QUADRO_RESUMO, _linhas_quadro_resumo(m, formatar),
                                                       [3.2, 2.9, 3.0, 3.1, 2.4, 3.0])
    alt_tabela = alt_linha * (len(linhas) + 1)
    tabela_nativa(s, 0.9, y + folga_tabela, 17.6, alt_tabela, cabecalho, linhas,
                  larguras=larguras, tam_fonte=tam_tabela, linha_bold_indices={len(linhas) - 1})
    y_nota = y + folga_tabela + alt_tabela + 0.2
    if "NÃO DISCRIMINADO" in cabecalho:
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
    cabecalho = ["EIXO", "Nº AÇÕES", "INVESTIMENTO", "T.C. FIRMADO", "T.C. NÃO FIRMADO"]
    linhas = []
    for eixo, linha in m["por_eixo"].iterrows():
        linhas.append([eixo, int(linha["QTD"]), abreviar_moeda(linha["INVESTIMENTO"]),
                        abreviar_moeda(linha["CONTRATADO"]), abreviar_moeda(linha["A_CONTRATAR"])])
    tabela_nativa(s, 0.9, 1.9, 18.2, 6.6, cabecalho, linhas,
                  larguras=[6.5, 2.2, 3.17, 3.17, 3.16], tam_fonte=15)
    return s


def slide_7_top10(prs, m):
    s = novo_slide(prs, "MAIORES AÇÕES COM T.C. NÃO FIRMADO")
    cabecalho = ["AÇÃO", "MUNICÍPIO", "T.C. NÃO FIRMADO", "STATUS", "FASE"]
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
                  "ainda com T.C. não firmado na base — verificar cadastro antes de divulgar.",
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
          "impede identificar a fonte do valor com T.C. firmado"),
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
        (abreviar_moeda(m["total_a_contratar"]), "ainda estão com T.C. não firmado na carteira do PAC-BA"),
        (abreviar_moeda(m["a_contratar_parcela_federal"]), "desse valor dependem de fonte Federal (OGU)"),
        (f"{m['gestao_federal_qtd']}", "ações sob execução direta de órgãos federais na carteira"),
        (abreviar_moeda(m["valor_clausula_vigente"]), f"com T.C. firmado têm cláusula suspensiva vigente "
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
        {"FONTE": "Não discriminado (T.C. firmado)", "CONTRATADO": m["contratado_nao_discriminado"],
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
                   for f, c, t in ((fc, cheia, "T.C. firmado"), (fa, clara, "T.C. não firmado")) if f > 0)


def _html_linha_tabela(rotulo, c_val, a_val, base, cheia, clara):
    # Igual ao PPTX, ver _linha_tabela, mas a barra segue a mesma lógica de
    # cor do T.C. firmado/não firmado (verde/laranja) em vez da cor por fonte
    # (cheia/clara seguem recebidos só pra manter a assinatura e o PPTX
    # intocado) — pedido do usuário em 08/11/2026.
    return f'''<div class="barra-linha">
      <div class="barra-rotulo" style="width:340px">{_html_esc(rotulo)}</div>
      <div class="barra-col-valor col-contratado">{_html_esc(abreviar_moeda(c_val))}</div>
      <div class="barra-trilho">{_html_segmentos_barra(c_val, a_val, base, VERDE, LARANJA)}</div>
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
    # Linha do gráfico dos slides por fase (igual ao PPTX, ver _linha_grafico_fase).
    # T.C. FIRMADO/NÃO FIRMADO nas mesmas cores da Visão Geral (verde/laranja),
    # tanto no texto quanto na barra agora (cheia/clara seguem recebidos só
    # pra manter a assinatura e o PPTX intocado) — pedido do usuário em
    # 08/11/2026.
    return f'''<div class="grafico-fase-linha nivel-{nivel}">
      <div class="gf-rotulo">{_html_esc(rotulo)}</div>
      <div class="barra-trilho">{_html_segmentos_barra(c_val, a_val, base, VERDE, LARANJA)}</div>
      <div class="gf-total">{_html_esc(abreviar_moeda(c_val + a_val))}</div>
      <div style="color:#{VERDE};font-weight:700">{_html_esc(abreviar_moeda(c_val))}</div>
      <div style="color:#{LARANJA};font-weight:700">{_html_esc(abreviar_moeda(a_val))}</div>
    </div>'''


def _html_grafico_fase(m, grafico, rotulo_secao=ROTULO_GRAFICO_FASE, legenda=None):
    rotulo, c_val, a_val, fontes = grafico
    html_legenda = (f'<span style="font-size:13px;font-weight:400;color:#5B6576;text-transform:none">{_html_esc(legenda)}</span>'
                    if legenda else "")
    # rotulo_secao vazio/None = sem o subtítulo (usado quando ele já apareceu
    # uma vez antes, fora do gráfico — ver slides por fase em
    # montar_html_apresentacao)
    titulo_html = (f'''<div class="rotulo-secao" style="display:flex;justify-content:space-between;align-items:baseline">
        <span>{_html_esc(rotulo_secao)}</span>{html_legenda}
      </div>''' if rotulo_secao else "")
    return f'''{titulo_html}
      <div class="grafico-fase-linha gf-cabecalho"><div></div><div></div><div>TOTAL</div><div>T.C. FIRMADO</div><div>T.C. NÃO FIRMADO</div></div>
      {_html_linha_grafico_fase(1, rotulo, c_val, a_val, m["total_investimento"], CINZA_FASE_CONTRATADO, CINZA_FASE_A_CONTRATAR)}
      {"".join(_html_linha_grafico_fase(2, r, c, a, c_val + a_val, cheia, clara) for r, c, a, cheia, clara in fontes
               if _linha_com_valor(c, a))}'''


def _html_cabecalho_tabela():
    return ('<div class="barra-linha cabecalho-tabela"><div class="barra-rotulo" style="width:340px"></div>'
            '<div class="barra-col-valor col-contratado">T.C. FIRMADO</div>'
            '<div class="barra-trilho" style="background:none;height:auto"></div>'
            '<div class="barra-col-valor">T.C. NÃO FIRMADO</div><div class="barra-col-valor">TOTAL</div></div>')


def _html_tabela(cabecalho, linhas, linhas_bold=(), cor_cabecalho=None):
    estilo = f' style="background:#{cor_cabecalho}"' if cor_cabecalho else ""
    ths = "".join(f"<th{estilo}>{_html_esc(c)}</th>" for c in cabecalho)
    trs = []
    for i, linha in enumerate(linhas):
        classe = " class=\"linha-bold\"" if i in linhas_bold else ""
        tds = "".join(f"<td>{_html_esc(v)}</td>" for v in linha)
        trs.append(f"<tr{classe}>{tds}</tr>")
    return f'<table class="pac-tabela"><thead><tr>{ths}</tr></thead><tbody>{"".join(trs)}</tbody></table>'


def _html_cabecalho_rodape_base(tema, pagina, total_paginas, conteudo_html, recorte="", classe_extra=""):
    nome_planilha = os.path.basename(pac.CAMINHO_PLANILHA) if hasattr(pac, "CAMINHO_PLANILHA") else \
        "PANORAMA - PAC ORIGINAL - PAC SELEÇÕES - 2026"
    posicao = pac.ultima_atualizacao.strftime("%m/%Y") if hasattr(pac, "ultima_atualizacao") else date.today().strftime("%m/%Y")
    return f'''<section class="slide{(" " + classe_extra) if classe_extra else ""}">
      <div class="cabecalho">
        <div class="titulo">PAC - {_html_esc(tema)}</div>
        <div class="subtitulo">COORDENAÇÃO GERAL DE ACOMPANHAMENTO DE POLÍTICAS ESTRATÉGICAS</div>
        <div class="base-atualizada">{_html_esc(texto_base_atualizada())}</div>
        <div class="governo"><div>GOVERNO DO ESTADO DA BAHIA</div><div>CASA CIVIL</div></div>
      </div>
      <div class="conteudo">{conteudo_html}</div>
      <div class="rodape">
        <span>Fonte: {_html_esc(nome_planilha)} · posição em {_html_esc(posicao)}{(" · Filtros: " + _html_esc(recorte)) if recorte else ""}</span>
        <span><a href="{LINK_CANVA}" target="_blank" rel="noopener">{LINK_CANVA}</a>&nbsp;&nbsp;&nbsp;{pagina}/{total_paginas}</span>
      </div>
    </section>'''


def montar_html_apresentacao(df, filtros=None, secoes=None):
    m = calcular_metricas(df)
    # manter_ordem=True: o painel de seleção do HTML permite reordenar os
    # cartões (inclusive intercalando as páginas do Canva), então a ordem
    # final é a que vier em `secoes`, não a ordem fixa do catálogo do PPTX
    # — pedido do usuário em 08/11/2026.
    paginas_canva = _paginas_canva_disponiveis()
    chaves_validas_html = {s["chave"] for s in SLIDES_APRESENTACAO} | {_chave_canva(n) for n, _ in paginas_canva}
    escolhidas = _preparar_plano(m, secoes, contagem_paginas=_paginas_por_chave_html(m),
                                 manter_ordem=True, chaves_validas=chaves_validas_html)
    recorte = descrever_filtros(filtros)

    # todo slide abaixo leva o mesmo recorte no rodapé; o número da página é
    # preenchido no fim, quando já se sabe quais slides entram
    def _html_cabecalho_rodape(tema, conteudo, classe_extra=""):
        return _html_cabecalho_rodape_base(tema, "§PAG§", m["total_slides"], conteudo, recorte=recorte,
                                           classe_extra=classe_extra)

    total_investimento = m["total_investimento"]

    # --- slide 1 ---
    cards_status = [
        ("CONCLUÍDO / INAUGURADO", m["qtd_concluido"], "2E7D32"),
        ("EM OBRAS / ANDAMENTO", m["qtd_andamento"], "0B5CAD"),
        ("SELEÇÃO / PREPARATÓRIA / LICITAÇÃO", m["qtd_preparatoria"], AZUL_CLARO_MARCA),
        ("PARALISADO", m["qtd_paralisado"], "C62828"),
    ]
    frac_contratado = (m["total_contratado"] / total_investimento * 100) if total_investimento else 0
    cartoes_status_html = f'''<div class="cartoes-grade">
        {"".join(f'<div class="cartao-status"><div class="cs-barra" style="background:#{c}"></div><div class="cs-num" style="color:#{c}">{qtd}</div><div class="cs-rotulo">{_html_esc(rot)}</div></div>' for rot, qtd, c in cards_status if qtd > 0)}
      </div>'''
    # panorama_html/slide1 são montados mais abaixo, depois de _html_quadro()
    # e do bloco de T.C. não firmado por fonte (R$12,5Bi) — página única com
    # a Visão Geral, pedido do usuário em 08/11/2026.

    # --- slide 2 ---
    slide2 = _html_cabecalho_rodape("COMPOSIÇÃO POR ORIGEM DO RECURSO", f'''
      <div class="grafico-origem">{_html_grafico_fase(m, _grafico_total_fontes(m), ROTULO_GRAFICO_ORIGEM)}</div>
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

    # --- Composição por Fase: as 4 fases juntas, atrás de um único botão de
    # detalhamento na página do Panorama (mesma lógica do botão de
    # "Investimento por Fase e Fonte") — pedido do usuário em 08/11/2026.
    # Antes era 1 slide por fase (07/11/2026), depois virou página própria
    # com as 4 juntas; agora é mais um bloco dentro do Panorama.
    blocos_fases_html = []
    for idx_fase, (fase, (rotulo_f, c_f, a_f, fontes_f)) in enumerate(zip(m["fases_ordem"], _fases_por_fonte(m)), 1):
        grafico = (rotulo_f, c_f, a_f, fontes_f)
        blocos = []
        for titulo, indice, cor in BLOCOS_OBJETOS:
            cabecalho_obj, linhas = _tabela_objetos(m, fase, titulo, indice)
            if linhas:
                blocos.append((cabecalho_obj, linhas, cor))
        nota = _nota_objetos(m, fase, [cab for cab, _, _ in blocos])
        id_detalhe = f"detalhe-fase-{idx_fase}"
        detalhamento_html = ""
        if blocos:
            tabelas_html = "".join(
                _html_tabela(cabecalho, linhas, linhas_bold={len(linhas) - 1}, cor_cabecalho=cor)
                for cabecalho, linhas, cor in blocos
            )
            detalhamento_html = f'''
        <button type="button" class="btn-detalhamento"
                onclick="alternarDetalhamento(this, '{id_detalhe}')">Ver detalhamento por objeto ▾</button>
        <div class="detalhamento-fase oculto" id="{id_detalhe}">
          {tabelas_html}
          {f'<div class="nota-rodape" style="margin-top:6px">{_html_esc(nota)}</div>' if nota else ""}
        </div>'''
        separador = '<hr class="separador-fase">' if idx_fase > 1 else ""
        blocos_fases_html.append(f'''
      <div class="composicao-fase">
        {separador}
        <div class="grafico-fase-bloco gf-ordem-tc">{_html_grafico_fase(m, grafico, rotulo_secao="")}</div>
        {detalhamento_html}
      </div>''')
    composicao_fase_html = "".join(blocos_fases_html)

    # --- slide 4 (novo): status por fase e origem do recurso ---
    blocos_fase = []
    for rotulo, chave, cheia, clara, c_val, a_val in _fontes_por_fase(m):
        if not _linha_com_valor(c_val, a_val):
            continue
        linhas_fase = "".join(_html_linha_tabela(rotulo_fase(f), *m["fonte_fase"][chave][f], c_val + a_val, cheia, clara)
                              for f in m["fases_ordem"])
        blocos_fase.append(f'''<div class="bloco-fonte-fase compacto">
          {_html_linha_tabela(rotulo, c_val, a_val, total_investimento, cheia, clara)}
          <div class="nivel-2">{linhas_fase}</div>
        </div>''')
    slide2b = _html_cabecalho_rodape("STATUS POR FASE E ORIGEM DO RECURSO", f'''
      <div class="rotulo-secao" style="display:flex;justify-content:space-between;align-items:baseline">
        <span>INVESTIMENTO POR FONTE E FASE (T.C. firmado + T.C. não firmado)*</span>
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
                                          "INVESTIMENTO TOTAL, POR EIXO E FONTE (T.C. firmado + T.C. não firmado)*",
                                          LEGENDA_EIXO_FONTE, _grupos_eixo_fonte(m))
    slides_fonte_eixo = _html_slides_eixo("INVESTIMENTO POR FONTE E EIXO",
                                          "INVESTIMENTO TOTAL, POR FONTE E EIXO (T.C. firmado + T.C. não firmado)*",
                                          LEGENDA_FASES, _grupos_fonte_eixo(m))

    # --- slide 3 ---
    total_contratado = m["total_contratado"]
    slide3 = _html_cabecalho_rodape("EXECUÇÃO DO T.C. FIRMADO, POR FONTE", f'''
      <div class="duas-colunas-desiguais">
        <div>
          <div class="rotulo-secao">TOTAL COM T.C. FIRMADO</div>
          <div class="valor-principal" style="color:#2E7D32">{_html_esc(abreviar_moeda(total_contratado))}</div>
        </div>
        <div class="texto-secundario" style="align-self:center">Detalhamento por fonte disponível para
        {_html_esc(pct(total_contratado * m["cobertura_detalhamento_pct"], total_contratado))} do valor com T.C. firmado.</div>
      </div>
      <div class="bloco-barras" style="margin-top:20px">
        {"".join(_html_barra(rot, valor, total_contratado, cor_b, 340) for rot, valor, cor_b in (
            ("APOIADO (OGU) - T.C. FIRMADO", m["contratado_ogu"], COR_OGU),
            ("FINANCIAMENTO - T.C. FIRMADO", m["contratado_financiamento"], COR_FINANCIAMENTO),
            ("RECURSO ESTADUAL - T.C. FIRMADO", m["contratado_estadual"], COR_ESTADUAL),
            ("NÃO DISCRIMINADO", m["contratado_nao_discriminado"], COR_NAO_DISCRIMINADO),
        ) if abs(valor) >= 0.5)}
      </div>
      <div class="caixa-pendencia">⚠ PENDÊNCIA — {m['qtd_sem_detalhamento']} ação sem nenhuma categoria de origem
      preenchida e mais ações com divergência entre Valor Contratado e o detalhamento — já sinalizadas no
      Controle de Qualidade do painel, aguardando correção de cadastro.</div>
    ''')

    # --- slide 4 ---
    por_eixo_ord = m["por_eixo"].sort_values("A_CONTRATAR", ascending=False).head(5)
    slide4 = _html_cabecalho_rodape("RECURSOS COM T.C. NÃO FIRMADO", f'''
      <div class="rotulo-secao">VALOR TOTAL COM T.C. NÃO FIRMADO</div>
      <div class="valor-principal" style="color:#E68A00">{_html_esc(abreviar_moeda(m["total_a_contratar"]))}</div>
      <div class="texto-secundario">{_html_esc(pct(m["total_a_contratar"], total_investimento))} do investimento
      total da carteira ainda não tem instrumento assinado</div>
      <div class="duas-colunas" style="margin-top:24px; align-items:start">
        <div>
          <div class="rotulo-secao" style="margin-bottom:14px">POR FONTE</div>
          <div class="bloco-barras">
            {"".join(_html_barra(rot, valor, m["total_a_contratar"], cor_b, 170) for rot, valor, cor_b in (
                ("OGU (Apoiado)", m["a_contratar_ogu"], COR_OGU_CLARO),
                ("Financiamento", m["a_contratar_financiamento"], COR_FINANCIAMENTO_CLARO),
                ("Recurso Estadual", m["a_contratar_estadual"], COR_ESTADUAL_CLARO),
            ) if abs(valor) >= 0.5)}
          </div>
        </div>
        <div>
          <div class="rotulo-secao" style="margin-bottom:14px">POR EIXO (top 5 em valor com T.C. não firmado)</div>
          <div class="bloco-barras">
            {"".join(_html_barra(eixo, float(linha["A_CONTRATAR"]), m["total_a_contratar"], "2F7F92", 230) for eixo, linha in por_eixo_ord.iterrows())}
          </div>
        </div>
      </div>
    ''')

    # --- cláusulas suspensivas: cards por ação, agrupados por eixo (pedido do
    # usuário em 10/11/2026) — ver _calcular_clausulas ---
    cores_prazo = {"vencida": VERMELHO, "proxima": LARANJA, "no_prazo": VERDE, "sem_data": COR_NAO_DISCRIMINADO}

    def _texto_prazo(acao):
        dias = acao["dias"]
        if dias is None:
            return "sem data"
        if dias < 0:
            return f"vencida há {-dias} dia{'s' if dias != -1 else ''}"
        if dias == 0:
            return "vence hoje"
        return f"vence em {dias} dia{'s' if dias != 1 else ''}"

    def _bloco_vigencia(acao, cor_p):
        # Vigência em destaque (pedido do usuário em 10/11/2026): faixa
        # própria com a data grande e o prazo restante num selo colorido.
        data = acao["vigencia"].strftime("%d/%m/%Y") if acao["vigencia"] else "Não informada"
        return f'''<div class="cl-vig" style="border-color:#{cor_p};background:#{cor_p}14">
              <div><div class="cl-vig-rotulo">Vigência da cláusula</div>
                <div class="cl-vig-data" style="color:#{cor_p}">{_html_esc(data)}</div></div>
              <span class="cl-vig-selo" style="background:#{cor_p}">{_html_esc(_texto_prazo(acao))}</span>
            </div>'''

    def _card_clausula(acao):
        cor_p = cores_prazo[acao["prazo"]]
        chips = "".join(f'<span class="cl-chip">{_html_esc(t)}</span>'
                        for t in (acao["situacao"], acao["fase"], acao["status"]) if t)
        motivo = (f'<div class="cl-motivo-rotulo">Motivo / pendência</div>'
                  f'<div class="cl-motivo" title="{_html_esc(acao["motivo"])}">{_html_esc(acao["motivo"])}</div>'
                  if acao["motivo"] else '<div class="cl-motivo cl-sem-motivo">Motivo não informado na planilha</div>')
        return f'''<div class="cl-card">
          <div class="cl-faixa" style="background:#{cor_p}"></div>
          <div class="cl-corpo">
            <div class="cl-topo">
              <div>
                <div class="cl-objeto">{_html_esc(acao["objeto"])}</div>
                <div class="cl-municipio">{_html_esc(acao["municipio"] or "Município não informado")}</div>
              </div>
              <div class="cl-valor">{_html_esc(abreviar_moeda(acao["valor"]))}</div>
            </div>
            <div class="cl-chips">{chips}</div>
            {_bloco_vigencia(acao, cor_p)}
            {motivo}
          </div>
        </div>'''

    prazos = m["clausulas_prazo"]
    kpis_clausula = [
        ("CLÁUSULA SUSPENSIVA NÃO RETIRADA", str(m["clausulas_qtd"]),
         f'{abreviar_moeda(m["clausulas_valor"])} em investimento', NAVY),
        ("VIGÊNCIA VENCIDA", str(prazos["vencida"]), "ações com prazo já expirado", VERMELHO),
        (f"VENCE EM ATÉ {DIAS_ALERTA_VIGENCIA} DIAS", str(prazos["proxima"]), "ações que pedem prioridade", LARANJA),
        ("CLÁUSULA RETIRADA (TOTAL)", str(m["qtd_clausula_retirada_total"]), "ações já liberadas", VERDE),
    ]
    kpis_html = "".join(f'''<div class="cartao-status"><div class="cs-barra" style="background:#{c}"></div>
        <div class="cs-num" style="color:#{c}">{_html_esc(n)}</div>
        <div class="cs-rotulo">{_html_esc(rot)}<div class="cl-kpi-sub">{_html_esc(sub)}</div></div></div>'''
                        for rot, n, sub, c in kpis_clausula)
    def _resumo_prazos_eixo(acoes_eixo):
        # Mini-contagem por prazo no cabeçalho do eixo — dá pra ver a
        # urgência mesmo com os cards recolhidos.
        partes = []
        for chave, rot in (("vencida", "vencida"), ("proxima", f"≤ {DIAS_ALERTA_VIGENCIA} dias"),
                           ("no_prazo", "no prazo"), ("sem_data", "sem data")):
            n = sum(1 for a in acoes_eixo if a["prazo"] == chave)
            if n:
                partes.append(f'<span class="cl-eixo-prazo" style="color:#{cores_prazo[chave]}">● {n} {rot}</span>')
        return "".join(partes)

    if m["clausulas_por_eixo"]:
        # Cada eixo com o seu botão de detalhamento (pedido do usuário em
        # 10/11/2026): cards recolhidos por padrão, mesmo "+" das outras
        # seções; o "Gerar PDF" já expande tudo antes de imprimir.
        grupos_clausula_html = "".join(f'''<div class="cl-eixo">
          <div class="cl-eixo-cabecalho">
            <span class="cl-eixo-titulo">
              <button type="button" class="btn-icone-detalhamento" title="Ver detalhamento do eixo"
                      onclick="alternarDetalhamento(this, 'detalhe-clausula-eixo-{i}')">+</button>
              <span class="cl-eixo-nome">{_html_esc(nome_eixo)}</span>
              {_resumo_prazos_eixo(acoes_eixo)}
            </span>
            <span class="cl-eixo-resumo">{len(acoes_eixo)} aç{"ões" if len(acoes_eixo) != 1 else "ão"} ·
              {_html_esc(abreviar_moeda(sum(a["valor"] for a in acoes_eixo)))}</span>
          </div>
          <div class="cl-grade oculto" id="detalhe-clausula-eixo-{i}">{"".join(_card_clausula(a) for a in acoes_eixo)}</div>
        </div>''' for i, (nome_eixo, acoes_eixo) in enumerate(m["clausulas_por_eixo"], 1))
    else:
        grupos_clausula_html = ('<div class="texto-secundario" style="margin-top:30px">Nenhuma ação com cláusula '
                                'suspensiva pendente neste recorte.</div>')
    legenda_prazo = "".join(f'<span><i style="background:#{cores_prazo[k]}"></i>{rot}</span>' for k, rot in (
        ("vencida", "Vigência vencida"), ("proxima", f"Vence em até {DIAS_ALERTA_VIGENCIA} dias"),
        ("no_prazo", "No prazo"), ("sem_data", "Sem data")))
    slide_clausulas = _html_cabecalho_rodape("CLÁUSULAS SUSPENSIVAS", f'''
      <div class="cartoes-grade cl-kpis">{kpis_html}</div>
      <div class="cl-legenda">{legenda_prazo}</div>
      {grupos_clausula_html}
      <div class="nota-rodape">* Valor = investimento total da ação (T.C. firmado + não firmado). Cards ordenados
      pela data de vigência da cláusula (mais urgente primeiro); eixos do maior para o menor valor.</div>
    ''', classe_extra="slide-continua")

    cabecalho_quadro, linhas_quadro = _sem_colunas_zeradas(CABECALHO_QUADRO_RESUMO,
                                                           _linhas_quadro_resumo(m, _mi_bi_2_casas))

    # --- panorama + visão geral: página única (pedido do usuário em
    # 08/11/2026; antes eram 2 páginas com investimento total, barra e cards
    # duplicados — agora o valor de T.C. não firmado por fonte e a tabela de
    # investimento por fase entram na página do Panorama) ---
    recursos_a_contratar = "" if abreviar_moeda(m["total_a_contratar"]) == "-" else f'''<div class="vg-recursos">
          <div class="rotulo-secao">{_html_esc(ROTULO_RECURSOS_A_CONTRATAR)}</div>
          <div class="valor-principal" style="color:#{LARANJA}">{_html_esc(abreviar_moeda(m["total_a_contratar"]))}</div>
          <div class="linha-icone-detalhamento">
            <span class="label-detalhamento">Detalhamento por fonte</span>
            <button type="button" class="btn-icone-detalhamento" title="Ver detalhamento por fonte"
                    onclick="alternarDetalhamento(this, 'detalhe-fontes-recurso')">+</button>
          </div>
          <div class="detalhamento-fontes oculto" id="detalhe-fontes-recurso">
            {"".join(f'<div class="vg-recurso-linha">{_html_esc(linha)}</div>' for linha in _linhas_recursos_a_contratar(m))}
          </div>
        </div>'''
    panorama_html = f'''<div class="visao-geral">
      <div class="vg-topo">
        <div>
          <div class="rotulo-secao">INVESTIMENTO TOTAL DA CARTEIRA</div>
          <div class="valor-principal" style="color:#0B1F3F">{_html_esc(abreviar_moeda(total_investimento))}</div>
          <div class="texto-secundario">{m["qtd_acoes"]} ações do PAC acompanhadas pelo Estado da Bahia</div>
        </div>
        {recursos_a_contratar}
      </div>
      <div class="barra-empilhada">
        <div class="barra-empilhada-seg" style="width:{frac_contratado:.2f}%;background:#2E7D32">
          <div class="be-titulo">T.C. FIRMADO · {_html_esc(abreviar_moeda(m["total_contratado"]))}</div>
          <div class="be-sub">{_html_esc(pct(m["total_contratado"], total_investimento))}</div>
        </div>
        <div class="barra-empilhada-seg" style="width:{100 - frac_contratado:.2f}%;background:#E68A00">
          <div class="be-titulo">T.C. NÃO FIRMADO · {_html_esc(abreviar_moeda(m["total_a_contratar"]))}</div>
          <div class="be-sub">{_html_esc(pct(m["total_a_contratar"], total_investimento))}</div>
        </div>
      </div>
      {cartoes_status_html}
      <div style="margin-top:40px">
        <div class="rotulo-secao" style="display:flex;align-items:baseline;gap:10px;margin-bottom:18px">
          <span>INVESTIMENTO POR FASE E FONTE</span>
          <button type="button" class="btn-icone-detalhamento" title="Ver tabela de investimento por fase"
                  onclick="alternarDetalhamento(this, 'detalhe-tabela-fase')">+</button>
        </div>
        <div class="oculto" id="detalhe-tabela-fase">
          {_html_tabela([c + "*" if c == "NÃO DISCRIMINADO" else c for c in cabecalho_quadro],
                        linhas_quadro, linhas_bold={len(linhas_quadro) - 1})}
          {f'<div class="nota-rodape">{_html_esc(_nota_nao_discriminado(m, com_remissao=False))}</div>'
           if "NÃO DISCRIMINADO" in cabecalho_quadro else ""}
        </div>
      </div>
      <div style="margin-top:40px">
        <div class="rotulo-secao" style="display:flex;align-items:baseline;gap:10px;margin-bottom:18px">
          <span>COMPOSIÇÃO POR FASE</span>
          <button type="button" class="btn-icone-detalhamento" title="Ver composição por fase"
                  onclick="alternarDetalhamento(this, 'detalhe-composicao-fase')">+</button>
        </div>
        <div class="oculto" id="detalhe-composicao-fase">
          {composicao_fase_html}
        </div>
      </div>
    </div>'''
    # slide-continua: a página tem conteúdo variável (botões de detalhamento),
    # então a altura precisa se ajustar em vez de ficar fixa e sobrepor o
    # rodapé quando os detalhamentos estão abertos — pedido do usuário em
    # 08/11/2026.
    slide1 = _html_cabecalho_rodape("PANORAMA DA CARTEIRA", panorama_html, classe_extra="slide-continua")

    css = '''
    * { box-sizing: border-box; }
    body { margin:0; padding:0; background:#e9ebef; font-family:'Roboto','Segoe UI',Arial,sans-serif; }
    /* Navegação em carrossel (1 slide por vez, com ‹/›, teclado e tela
       cheia) — mesmo recurso do projeto Relatório SSP, pedido do usuário em
       08/11/2026. Fora da tela (display:none), o slide some; o ativo fica
       centralizado e limitado ao viewport, preservando a proporção pelas
       mesmas regras de width/aspect-ratio já usadas na impressão. Na
       impressão (@media print mais abaixo) isso tudo é desfeito e todos os
       slides voltam a aparecer, um por folha, como já funcionava antes. */
    .slide { display:none; width:1480px; max-width:100%; aspect-ratio:20/11.25; margin:0 auto 40px; background:#fff;
      box-shadow:0 4px 24px rgba(0,0,0,.2); position:relative; overflow:hidden; border-radius:6px; }
    .slide.ativo { display:block; position:fixed; inset:0; margin:auto; max-width:100vw; max-height:100vh; }
    .slide-continua.ativo { max-height:100vh; overflow-y:auto; }
    /* O rodapé fica preso ao fundo da caixa do slide (position:absolute),
       mas no carrossel o slide-continua passa a rolar por dentro — o
       rodapé ficaria sempre por cima do conteúdo, em vez de abaixo dele.
       Some só na tela; a impressão continua mostrando normalmente — pedido
       do usuário em 08/11/2026. */
    .slide-continua.ativo .rodape { display:none; }
    /* Espaçamentos do cabeçalho reduzidos pra caber a linha "Base atualizada
       em" sem aumentar a altura total (o conteúdo dos slides não se mexe) —
       pedido do usuário em 09/11/2026. */
    .cabecalho { background:#0B1F3F; color:#fff; padding:14px 40px 8px; position:relative; }
    .cabecalho .titulo { font-size:30px; font-weight:700; }
    .cabecalho .subtitulo { font-size:14px; font-weight:700; color:#C9D6EA; margin-top:4px; }
    .cabecalho .base-atualizada { font-size:12px; line-height:1.3; color:#C9D6EA; margin-top:3px; }
    .cabecalho .governo { position:absolute; top:14px; right:40px; text-align:right; font-weight:700;
      font-size:17px; color:#4FA3E0; line-height:1.7; }
    .conteudo { padding: 28px 40px; }
    .rodape { position:absolute; bottom:14px; left:40px; right:40px; display:flex; justify-content:space-between;
      font-size:12px; color:#5B6576; }
    .rodape a { color:#5B6576; }
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
    .grafico-fase-linha { display:grid; grid-template-columns:400px 1fr 175px 175px 145px; column-gap:14px;
      align-items:center; color:#5B6576; font-size:14px; margin-bottom:2px; }
    .grafico-fase-linha .barra-trilho { flex:none; height:13px; }
    .grafico-fase-linha.nivel-1 { font-size:17px; margin-bottom:4px; }
    .grafico-fase-linha.nivel-1 .barra-trilho { height:20px; }
    .grafico-fase-linha > div:nth-child(3), .grafico-fase-linha > div:nth-child(4),
    .grafico-fase-linha > div:nth-child(5) { text-align:right; }
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
    /* Composição por Fase: colunas na ordem T.C. FIRMADO | gráfico | T.C. NÃO
       FIRMADO | TOTAL, igual às tabelas de status/eixo — pedido do usuário em
       09/11/2026. Só reordena pela grade (filhos: 1 rótulo, 2 barra, 3 total,
       4 firmado, 5 não firmado); a Composição por Origem segue como estava. */
    .gf-ordem-tc .grafico-fase-linha { grid-template-columns:400px 175px 1fr 175px 145px; }
    .gf-ordem-tc .grafico-fase-linha > div:nth-child(1) { order:1; }
    .gf-ordem-tc .grafico-fase-linha > div:nth-child(4) { order:2; text-align:right; }
    .gf-ordem-tc .grafico-fase-linha > div:nth-child(2) { order:3; }
    .gf-ordem-tc .grafico-fase-linha > div:nth-child(5) { order:4; text-align:left; }
    .gf-ordem-tc .grafico-fase-linha > div:nth-child(3) { order:5; text-align:right; }
    .composicao-fase table.pac-tabela { font-size:15px; margin-top:14px; table-layout:fixed; }
    .composicao-fase table.pac-tabela th { padding:3px 14px; font-size:14px; }
    .composicao-fase table.pac-tabela td { padding:2px 14px; }
    .composicao-fase table.pac-tabela th:first-child { width:26%; }
    .visao-geral .rotulo-secao { margin:0 0 8px !important; font-size:16px; }
    .visao-geral .nota-rodape { margin:8px 0 14px; font-size:12px; }
    .visao-geral table.pac-tabela { font-size:15px; margin-top:0; }
    .visao-geral table.pac-tabela th { padding:7px 14px; font-size:14px; }
    .visao-geral table.pac-tabela td { padding:5px 14px; }
    .visao-geral .valor-principal { font-size:66px; margin-top:0; line-height:1.1; }
    .visao-geral .barra-empilhada { height:64px; margin-top:6px; }
    .vg-topo { display:flex; justify-content:space-between; }
    .vg-recursos { text-align:right; }
    .vg-recurso-linha { font-size:22px; font-weight:700; color:#E68A00; line-height:1.3; text-align:right; }
    .visao-geral .barra-empilhada-seg { white-space:nowrap; overflow:hidden; }
    .visao-geral .barra-empilhada-seg { padding:8px 16px; }
    .visao-geral .be-titulo { font-size:16px; }
    .visao-geral .be-sub { font-size:14px; margin-top:2px; }
    .visao-geral .cartoes-grade { margin-top:28px; }
    .visao-geral .cs-barra { height:6px; }
    .visao-geral .cs-num { font-size:28px; padding:2px 16px 0; }
    .visao-geral .cs-rotulo { font-size:13px; padding:0 16px 5px; }
    .barra-col-valor { width:130px; flex-shrink:0; font-size:15px; color:#5B6576; }
    .col-contratado { text-align:right; }
    .col-a-contratar { font-weight:700; color:#0B1F3F; text-align:right; }
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

    /* Cláusulas suspensivas: cards por ação agrupados por eixo — pedido do
       usuário em 10/11/2026. */
    .cl-kpis { margin-top:0; }
    .cl-kpis .cs-num { font-size:40px; padding:10px 18px 0; }
    .cl-kpis .cs-rotulo { font-size:14px; padding:6px 18px 14px; }
    .cl-kpi-sub { font-size:13px; font-weight:400; margin-top:3px; }
    .cl-legenda { display:flex; gap:22px; font-size:13px; color:#5B6576; margin:18px 0 6px; }
    .cl-legenda i { display:inline-block; width:12px; height:12px; border-radius:3px; margin-right:6px;
      vertical-align:-1px; }
    .cl-eixo { margin-top:26px; }
    .cl-eixo-cabecalho { display:flex; justify-content:space-between; align-items:baseline;
      border-bottom:2px solid #0B1F3F; padding-bottom:6px; margin-bottom:14px; }
    .cl-eixo-titulo { display:flex; align-items:center; gap:12px; flex-wrap:wrap; }
    .cl-eixo-nome { font-size:20px; font-weight:700; color:#0B1F3F; text-transform:uppercase; }
    .cl-eixo-prazo { font-size:13px; font-weight:700; }
    .cl-grade.oculto { display:none; }
    .cl-vig { display:flex; justify-content:space-between; align-items:center; gap:10px; margin-top:12px;
      border:1px solid; border-left-width:4px; border-radius:8px; padding:8px 12px; }
    .cl-vig-rotulo { font-size:11px; font-weight:700; color:#5B6576; text-transform:uppercase; }
    .cl-vig-data { font-size:26px; font-weight:900; line-height:1.1; }
    .cl-vig-selo { color:#fff; font-size:13px; font-weight:700; border-radius:14px; padding:5px 12px; white-space:nowrap; }
    .cl-eixo-resumo { font-size:15px; font-weight:700; color:#5B6576; }
    .cl-grade { display:grid; grid-template-columns:repeat(3, 1fr); gap:16px; }
    .cl-card { display:flex; background:#fff; border:1px solid #D8DCE3; border-radius:10px; overflow:hidden; }
    .cl-faixa { width:6px; flex-shrink:0; }
    .cl-corpo { padding:14px 16px; flex:1; min-width:0; }
    .cl-topo { display:flex; justify-content:space-between; gap:12px; align-items:flex-start; }
    .cl-objeto { font-size:16px; font-weight:700; color:#0B1F3F; }
    .cl-municipio { font-size:14px; color:#5B6576; margin-top:2px; }
    .cl-valor { font-size:18px; font-weight:900; color:#0B1F3F; white-space:nowrap; }
    .cl-chips { display:flex; flex-wrap:wrap; gap:6px; margin-top:10px; }
    .cl-chip { font-size:11px; font-weight:700; color:#5B6576; background:#F2F4F7; border-radius:10px;
      padding:3px 9px; text-transform:uppercase; }
    .cl-vigencia { font-size:13px; font-weight:700; margin-top:10px; }
    .cl-motivo-rotulo { font-size:11px; font-weight:700; color:#5B6576; text-transform:uppercase; margin-top:10px; }
    .cl-motivo { font-size:13px; color:#0B1F3F; margin-top:3px; line-height:1.35; white-space:pre-line;
      display:-webkit-box; -webkit-line-clamp:4; -webkit-box-orient:vertical; overflow:hidden; }
    .cl-sem-motivo { color:#8C8C8C; font-style:italic; margin-top:10px; }

    /* Composição por Fase: página contínua (altura livre, sem recortar
       conteúdo) e detalhamento por objeto escondido atrás de um botão. */
    .slide-continua { aspect-ratio:unset; height:auto; overflow:visible; }
    .slide-continua .conteudo { padding-bottom:40px; }
    .oculto { display:none; }
    .btn-detalhamento { margin-top:20px; padding:10px 20px; font-size:14px; font-weight:700; color:#0B1F3F;
      background:#F2F4F7; border:1px solid #D8DCE3; border-radius:8px; cursor:pointer; font-family:inherit; }
    .btn-detalhamento:hover { background:#E6E9EE; }
    .detalhamento-fase { margin-top:18px; }
    .separador-fase { border:none; border-top:1px solid #D8DCE3; margin:28px 0 24px; }

    /* Panorama + Visão Geral: botões de detalhamento (em ícone) do valor de
       T.C. não firmado por fonte e da tabela de investimento por fase —
       pedido do usuário em 08/11/2026. */
    .detalhamento-fontes { margin-top:10px; }
    .vg-recursos .linha-icone-detalhamento, .vg-recursos .detalhamento-fontes { text-align:right; }
    .linha-icone-detalhamento { display:flex; align-items:center; gap:10px; margin-top:10px; }
    .vg-recursos .linha-icone-detalhamento { justify-content:flex-end; }
    .label-detalhamento { font-size:13px; font-weight:700; color:#5B6576; text-transform:uppercase; }
    .btn-icone-detalhamento { width:28px; height:28px; border-radius:50%; border:1px solid #D8DCE3;
      background:#F2F4F7; color:#0B1F3F; font-size:17px; font-weight:700; line-height:1; cursor:pointer;
      display:inline-flex; align-items:center; justify-content:center; padding:0; font-family:inherit; flex-shrink:0; }
    .btn-icone-detalhamento:hover { background:#E6E9EE; }

    /* Botão flutuante "Gerar PDF" — expande todos os detalhamentos ocultos
       antes de chamar a impressão do navegador — pedido do usuário em
       08/11/2026. */
    .btn-gerar-pdf { position:fixed; top:18px; right:24px; z-index:1000; padding:13px 24px; font-size:15px;
      font-weight:700; color:#fff; background:#0B5CAD; border:none; border-radius:8px; cursor:pointer;
      box-shadow:0 4px 16px rgba(0,0,0,.3); font-family:inherit; }
    .btn-gerar-pdf:hover { background:#0A4E92; }
    body.tela-cheia .btn-gerar-pdf { display:none; }

    /* Barra de navegação do carrossel (‹/›, contador, dica de tela cheia) —
       mesmo recurso do projeto Relatório SSP, pedido do usuário em
       08/11/2026. Some sozinha depois de 2,5s sem mexer o mouse. */
    #nav-carrossel { position:fixed; left:14px; bottom:14px; z-index:1000; color:#C9D6EA; font-size:12px;
      display:flex; gap:10px; align-items:center; user-select:none; background:rgba(11,31,63,.85);
      padding:6px 10px; border-radius:8px; opacity:0; transition:opacity .3s; }
    body.mostra-nav #nav-carrossel, #nav-carrossel:hover { opacity:1; }
    #nav-carrossel button { background:rgba(255,255,255,.1); color:#fff; border:none; border-radius:6px;
      width:30px; height:26px; cursor:pointer; font-size:14px; font-family:inherit; }
    #nav-carrossel button:hover { background:rgba(255,255,255,.22); }

    /* Página do Canva (imagem exportada) — pedido do usuário em 08/11/2026.
       Sem cabeçalho/rodapé do dash; a imagem já vem na mesma proporção
       16:9 do slide, então só precisa preencher a caixa. O display:flex só
       entra junto com .ativo (mesma especificidade de ".slide.ativo", e
       vem depois dela no CSS) — sem isso, "display:flex" sozinho tinha a
       mesma especificidade do "display:none" padrão do carrossel e vencia
       por ordem, deixando TODAS as páginas do Canva visíveis ao mesmo
       tempo, sobrepostas — bug corrigido em 08/11/2026. */
    .slide-canva.ativo { display:flex; padding:0; }
    .img-canva { width:100%; height:100%; object-fit:contain; background:#fff; }

    /* Impressão ("Gerar PDF"): página 16:9 do tamanho exato do slide, sem
       margem, pra exibição em tela (projetor/TV) — pedido do usuário em
       09/11/2026; antes era A4 paisagem com 10mm de margem e zoom 0,705.
       1480 x 832,5px é 16:9 exato; a altura vai arredondada pra 833px pra
       sobra de arredondamento nunca jogar 1px numa página em branco. Sem
       margem, o navegador também deixa de imprimir data, título e caminho
       do arquivo no topo/rodapé de cada página. */
    @page { size: 1480px 833px; margin: 0; }
    @media print {
      body { padding:0; background:#fff; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
      /* Desfaz o carrossel na impressão: o slide ativo usa position:fixed e
         limites de viewport só pra tela — aqui tudo volta ao fluxo normal,
         um slide por folha, como já funcionava antes do carrossel existir
         — pedido do usuário em 08/11/2026. */
      .slide, .slide.ativo { display:block !important; position:static !important; inset:auto !important;
        width:1480px; max-width:none !important; max-height:none !important; zoom:1; margin:0;
        box-shadow:none; border-radius:0;
        break-after:page; page-break-after:always; break-inside:avoid; page-break-inside:avoid; }
      .slide:last-of-type { break-after:auto; page-break-after:auto; }
      .slide-continua.ativo { overflow:visible !important; }
      #nav-carrossel { display:none !important; }
      /* Slides de altura variável (Panorama e Composição por Fase) podem
         ficar maiores que uma folha quando os detalhamentos são expandidos
         pelo botão "Gerar PDF" — sem isso, break-inside:avoid força o
         navegador a inserir uma página em branco antes deles. Em vez disso,
         evita-se cortar apenas as peças menores (cartões, linhas de gráfico,
         linhas de tabela) — pedido do usuário em 08/11/2026. */
      .slide-continua { break-inside:auto; page-break-inside:auto; }
      .cl-card, .cl-eixo-cabecalho { break-inside:avoid; page-break-inside:avoid; }
      .cl-eixo-cabecalho { break-after:avoid; page-break-after:avoid; }
      .cl-motivo { -webkit-line-clamp:unset; display:block; }
      .cartao-status, .cartao-destaque, .cartao-pendencia, .grafico-fase-linha,
      .bloco-fonte-fase, .barra-linha, .pac-tabela tr {
        break-inside:avoid; page-break-inside:avoid;
      }
      /* Gráfico completo de uma fase (TOTAL + linhas por fonte) e cada
         tabela de detalhamento por objeto não podem ser separados entre
         páginas — pedido do usuário em 08/11/2026. A fase inteira (gráfico +
         detalhamento) também evita quebrar, pra ela nunca ficar sozinha no
         rodapé de uma folha com o resto empurrado pra próxima — se não
         couber de jeito nenhum numa página, as regras acima garantem que ao
         menos a quebra cai numa borda limpa (entre gráfico e tabela, ou
         entre tabelas), nunca no meio de uma — pedido do usuário em
         08/11/2026. */
      .grafico-fase-bloco, .pac-tabela, .composicao-fase { break-inside:avoid; page-break-inside:avoid; }
      .rotulo-secao { break-after:avoid; page-break-after:avoid; }
      .btn-gerar-pdf { display:none !important; }
    }
    '''

    # só os slides escolhidos, numerados na ordem
    por_chave = {"PANORAMA": [slide1], "VISAO_GERAL": [], "COMP_ORIGEM": [slide2], "COMP_FASE": [],
                 "STATUS_FASE": [slide2b],
                 "EIXO_FONTE": slides_eixo_fonte, "FONTE_EIXO": slides_fonte_eixo, "EXECUCAO": [slide3],
                 "A_CONTRATAR": [slide4], "CLAUSULAS": [slide_clausulas], "TABELA_RESUMO": [], "QUADRO_FASE": [],
                 "DETALHE_EIXO": [],
                 "MAIORES_ACOES": [], "PENDENCIAS": [], "SINTESE": []}

    # --- páginas extras: imagens exportadas do site do Canva, uma por
    # página, só geradas pras que realmente foram escolhidas (evita ler e
    # recomprimir imagem de página que nem entrou na apresentação) — pedido
    # do usuário em 08/11/2026.
    for numero, nome_arquivo in paginas_canva:
        chave = _chave_canva(numero)
        if chave not in escolhidas:
            continue
        uri = _data_uri_imagem(os.path.join(PASTA_SLIDES_CANVA, nome_arquivo))
        por_chave[chave] = [f'''<div class="slide slide-canva">
      <img class="img-canva" src="{uri}" alt="Apresentação Canva — página {numero}">
    </div>''']

    corpo = "".join(html_slide.replace("§PAG§", str(numero)) for numero, html_slide in enumerate(
        (h for chave in escolhidas for h in por_chave[chave]), 1))

    return f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Apresentação CGAPE</title>
<style>{css}</style>
</head>
<body>
<button type="button" class="btn-gerar-pdf" onclick="gerarPdfCompleto()">Gerar PDF</button>
<div id="nav-carrossel">
  <button type="button" id="btn-slide-ant" title="Slide anterior (←)">‹</button>
  <span id="contador-slide"></span>
  <button type="button" id="btn-slide-prox" title="Próximo slide (→)">›</button>
  <span>F: tela cheia</span>
</div>
{corpo}
<script>
function alternarDetalhamento(botao, idAlvo) {{
  var alvo = document.getElementById(idAlvo);
  if (!alvo) return;
  var vaiAbrir = alvo.classList.contains('oculto');
  alvo.classList.toggle('oculto', !vaiAbrir);
  if (botao.classList.contains('btn-icone-detalhamento')) {{
    botao.textContent = vaiAbrir ? '–' : '+';
    botao.title = botao.title.indexOf('Ver') === 0 ? botao.title.replace('Ver', 'Ocultar')
                                                    : botao.title.replace('Ocultar', 'Ver');
  }} else {{
    botao.textContent = vaiAbrir ? 'Ocultar detalhamento por objeto ▴' : 'Ver detalhamento por objeto ▾';
  }}
}}

function gerarPdfCompleto() {{
  // Expande todo detalhamento oculto para sair no PDF e devolve ao estado
  // anterior depois que a caixa de impressão do navegador é fechada.
  var ocultos = Array.prototype.slice.call(document.querySelectorAll('.oculto'));
  ocultos.forEach(function (el) {{ el.classList.remove('oculto'); }});
  document.querySelectorAll('.btn-detalhamento').forEach(function (b) {{
    b.textContent = 'Ocultar detalhamento por objeto ▴';
  }});
  document.querySelectorAll('.btn-icone-detalhamento').forEach(function (b) {{
    b.textContent = '–';
    if (b.title.indexOf('Ver') === 0) b.title = b.title.replace('Ver', 'Ocultar');
  }});

  function restaurar() {{
    ocultos.forEach(function (el) {{ el.classList.add('oculto'); }});
    document.querySelectorAll('.btn-detalhamento').forEach(function (b) {{
      b.textContent = 'Ver detalhamento por objeto ▾';
    }});
    document.querySelectorAll('.btn-icone-detalhamento').forEach(function (b) {{
      b.textContent = '+';
      if (b.title.indexOf('Ocultar') === 0) b.title = b.title.replace('Ocultar', 'Ver');
    }});
    window.removeEventListener('afterprint', restaurar);
  }}
  window.addEventListener('afterprint', restaurar);
  window.print();
}}

// Carrossel (1 slide por vez, ‹/›, teclado, tela cheia) — mesmo recurso do
// projeto Relatório SSP, pedido do usuário em 08/11/2026.
(function () {{
  var slides = Array.prototype.slice.call(document.querySelectorAll('.slide'));
  if (!slides.length) return;
  var atual = Math.max(0, Math.min(slides.length - 1, (parseInt(location.hash.slice(1), 10) || 1) - 1));
  var contador = document.getElementById('contador-slide');

  function mostrar(indice) {{
    atual = Math.max(0, Math.min(slides.length - 1, indice));
    slides.forEach(function (s, k) {{ s.classList.toggle('ativo', k === atual); }});
    if (contador) contador.textContent = (atual + 1) + '/' + slides.length;
    history.replaceState(null, '', '#' + (atual + 1));
  }}

  function telaCheia() {{
    if (document.fullscreenElement) document.exitFullscreen();
    else if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen();
  }}

  // Esconde o botão "Gerar PDF" durante a tela cheia — pedido do usuário em
  // 08/11/2026.
  document.addEventListener('fullscreenchange', function () {{
    document.body.classList.toggle('tela-cheia', !!document.fullscreenElement);
  }});

  var navTimer = null;
  document.addEventListener('mousemove', function () {{
    document.body.classList.add('mostra-nav');
    clearTimeout(navTimer);
    navTimer = setTimeout(function () {{ document.body.classList.remove('mostra-nav'); }}, 2500);
  }});

  document.addEventListener('keydown', function (e) {{
    // Espaço/Enter ficam de fora de propósito (diferente do Relatório SSP):
    // nos nossos slides eles ativam botões como "Ver detalhamento", e usar
    // a mesma tecla pra também virar o slide causaria duplo efeito.
    if (['ArrowRight', 'PageDown'].indexOf(e.key) !== -1) {{ mostrar(atual + 1); e.preventDefault(); }}
    else if (['ArrowLeft', 'PageUp', 'Backspace'].indexOf(e.key) !== -1) {{ mostrar(atual - 1); e.preventDefault(); }}
    else if (e.key === 'Home') mostrar(0);
    else if (e.key === 'End') mostrar(slides.length - 1);
    else if (e.key === 'f' || e.key === 'F') telaCheia();
  }});

  var btnAnt = document.getElementById('btn-slide-ant');
  var btnProx = document.getElementById('btn-slide-prox');
  if (btnAnt) btnAnt.onclick = function () {{ mostrar(atual - 1); }};
  if (btnProx) btnProx.onclick = function () {{ mostrar(atual + 1); }};

  mostrar(atual);
}})();
</script>
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
    {"chave": "EXECUCAO", "titulo": "Execução do T.C. Firmado, por Fonte", "layout": "barras"},
    {"chave": "A_CONTRATAR", "titulo": "Recursos com T.C. Não Firmado", "layout": "barras"},
    {"chave": "CLAUSULAS", "titulo": "Cláusulas Suspensivas", "layout": "cards_tabela"},
    {"chave": "TABELA_RESUMO", "titulo": "Tabela-Resumo por Fonte", "layout": "tabela"},
    {"chave": "QUADRO_FASE", "titulo": "Quadro Resumo de Investimentos por Fase", "layout": "tabela"},
    {"chave": "DETALHE_EIXO", "titulo": "Detalhe por Eixo do PAC", "layout": "tabela"},
    {"chave": "MAIORES_ACOES", "titulo": "Maiores Ações com T.C. Não Firmado", "layout": "tabela"},
    {"chave": "PENDENCIAS", "titulo": "Pendências que Condicionam a Contratação", "layout": "cards_tabela"},
    {"chave": "SINTESE", "titulo": "Síntese e Pleito", "layout": "texto"},
]

# Removidas da apresentação (HTML e PPTX) a pedido do usuário em 10/11/2026.
# As funções de desenho continuam no arquivo: para trazer alguma de volta,
# basta tirar a chave deste conjunto.
SLIDES_REMOVIDOS = {"COMP_ORIGEM", "STATUS_FASE", "EIXO_FONTE", "FONTE_EIXO", "EXECUCAO", "A_CONTRATAR"}
SLIDES_APRESENTACAO = [s for s in SLIDES_APRESENTACAO if s["chave"] not in SLIDES_REMOVIDOS]


# Chaves que só existem como slide/página no PPTX — no HTML, VISAO_GERAL e
# COMP_FASE foram mescladas na página do Panorama e as demais saíram da
# apresentação — pedido do usuário em 08/11/2026.
SLIDES_OCULTOS_NO_HTML = {"VISAO_GERAL", "COMP_FASE", "TABELA_RESUMO", "QUADRO_FASE", "DETALHE_EIXO",
                          "MAIORES_ACOES", "PENDENCIAS", "SINTESE"}
# E o contrário: seções que só existem no HTML (cards de cláusula suspensiva,
# pedido do usuário em 10/11/2026) — não entram no PPTX nem no seu catálogo.
SLIDES_SO_NO_HTML = {"CLAUSULAS"}


def _paginas_por_chave(m):
    return {"CLAUSULAS": 0,
            "COMP_FASE": len(_paginas_composicao_fase(m)),
            "EIXO_FONTE": len(_paginar_grupos(_grupos_eixo_fonte(m))),
            "FONTE_EIXO": len(_paginar_grupos(_grupos_fonte_eixo(m)))}


def _paginas_por_chave_html(m):
    # VISAO_GERAL e COMP_FASE foram incorporadas à página do PANORAMA, e
    # TABELA_RESUMO/QUADRO_FASE/DETALHE_EIXO/MAIORES_ACOES/PENDENCIAS/SINTESE
    # saíram do HTML (pedido do usuário em 08/11/2026) — nenhuma dessas
    # ocupa página própria no HTML; todas continuam existindo como slides
    # do PPTX.
    return dict(_paginas_por_chave(m), **{chave: 0 for chave in SLIDES_OCULTOS_NO_HTML},
                **{chave: 1 for chave in SLIDES_SO_NO_HTML})


def _preparar_plano(m, secoes, contagem_paginas=None, manter_ordem=False, chaves_validas=None):
    # Decide quais slides entram e em que ordem, e guarda em m a numeração,
    # usada nas remissões e nos rodapés.
    # manter_ordem=False (PPTX e HTML sem o painel novo): None/vazio = todos,
    # na ordem fixa do catálogo; chave desconhecida é ignorada.
    # manter_ordem=True (HTML, pedido do usuário em 08/11/2026): respeita a
    # ordem de `secoes` tal como veio (o painel de seleção agora deixa
    # reordenar os cartões, inclusive intercalando a apresentação Canva, que
    # não existe no catálogo do PPTX) — por isso precisa de `chaves_validas`
    # própria em vez de só olhar SLIDES_APRESENTACAO.
    if manter_ordem:
        validas = chaves_validas if chaves_validas is not None else {s["chave"] for s in SLIDES_APRESENTACAO}
        vistas = set()
        escolhidas = []
        for chave in (secoes or []):
            if chave in validas and chave not in vistas:
                escolhidas.append(chave)
                vistas.add(chave)
        escolhidas = escolhidas or [s["chave"] for s in SLIDES_APRESENTACAO]
    else:
        pedidas = set(secoes or [])
        escolhidas = [s["chave"] for s in SLIDES_APRESENTACAO if s["chave"] in pedidas]
        escolhidas = escolhidas or [s["chave"] for s in SLIDES_APRESENTACAO]
    paginas = contagem_paginas if contagem_paginas is not None else _paginas_por_chave(m)
    m["num_slide"], proximo = {}, 1
    for chave in escolhidas:
        m["num_slide"][chave] = proximo
        proximo += paginas.get(chave, 1)
    m["total_slides"] = proximo - 1
    return escolhidas


def listar_slides(df, formato=None):
    # "Visão Geral da Carteira" e "Composição por Fase" foram mescladas no
    # Panorama, e as páginas de tabela/listas que seguem (Tabela-Resumo,
    # Quadro Resumo por Fase, Detalhe por Eixo, Maiores Ações, Pendências,
    # Síntese) saíram do HTML — ambas só valem pro HTML, no PPTX tudo
    # continua existindo como slide próprio. Somem do catálogo só quando o
    # painel pede a lista pra montar o HTML, porque marcar/desmarcar essas
    # opções não teria efeito nenhum nesse formato — pedido do usuário em
    # 08/11/2026.
    m = calcular_metricas(df)
    paginas = _paginas_por_chave_html(m) if formato == "html" else _paginas_por_chave(m)
    catalogo = [s for s in SLIDES_APRESENTACAO
                if not (formato == "html" and s["chave"] in SLIDES_OCULTOS_NO_HTML)
                and not (formato != "html" and s["chave"] in SLIDES_SO_NO_HTML)]
    if formato == "html":
        # Uma página do catálogo por imagem do Canva disponível em
        # PASTA_SLIDES_CANVA — só existe no HTML (não dá pra desenhar numa
        # slide de PowerPoint algo que não faz parte do painel). Entram no
        # fim da lista por padrão, mas o painel deixa reordenar pra
        # intercalar com as seções do dash — pedido do usuário em
        # 08/11/2026.
        catalogo = catalogo + [{"chave": _chave_canva(numero), "titulo": f"Canva — página {numero}",
                                "layout": "texto"} for numero, _ in _paginas_canva_disponiveis()]
    return [dict(s, paginas=paginas.get(s["chave"], 1), marcada=True) for s in catalogo]


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
        if chave in desenhar:  # seções só do HTML (SLIDES_SO_NO_HTML) não têm slide
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

    caminho_pptx = os.path.join(PASTA_BASE, "APRESENTACAO CGAPE.pptx")
    prs.save(caminho_pptx)
    print("PPTX gerado:", caminho_pptx)

    caminho_xlsx = os.path.join(PASTA_BASE, "APOIO APRESENTACAO CGAPE.xlsx")
    gerar_xlsx_apoio(m, caminho_xlsx)
    print("XLSX de apoio gerado:", caminho_xlsx)

    caminho_html = os.path.join(PASTA_BASE, "APRESENTACAO CGAPE.html")
    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(montar_html_apresentacao(pac.df_original))
    print("HTML gerado:", caminho_html)


if __name__ == "__main__":
    main()
