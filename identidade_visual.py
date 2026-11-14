# -*- coding: utf-8 -*-
"""Identidade visual única do projeto CGAPE - BALANÇO PAC.

Fonte única dos tokens de cor e tipografia usados por TODAS as saídas do
projeto (painel desktop/web, apresentação HTML/PPTX da Análise Financeira,
Resumo de Ações, relatório PDF e seção MONITORAMENTO) — pedido do usuário
em 14/11/2026 ("Prompt — Identidade visual única + tela HOME").

A REFERÊNCIA é o HTML da apresentação (`montar_html_apresentacao`, em
gerar_pptx_pac.py): todos os tokens abaixo foram extraídos de lá. As demais
saídas passam a importar deste módulo em vez de repetir hex soltos ou usar
paletas próprias (AGF, no painel; ad-hoc, no relatório de monitoramento).

Este módulo não importa nada do projeto (fica no "fundo" da árvore de
imports) — só stdlib, e Pillow/reportlab/python-pptx são importados de
forma preguiçosa, dentro das funções que precisam deles, pra não forçar
essas dependências em quem só quer os hex/CSS.
"""
import base64
import os
import sys

if getattr(sys, "frozen", False):
    PASTA_BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
    PASTA_BASE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# Paleta (hex SEM "#" — pronto pra RGBColor.from_string do python-pptx; use
# hex_css()/cor_reportlab() abaixo quando precisar do "#" ou de um objeto
# colors.HexColor do ReportLab).
# ---------------------------------------------------------------------------
NAVY = "0B1F3F"
AZUL = "0B5CAD"
TEAL = "2F7F92"
LARANJA = "E68A00"
VERDE = "2E7D32"
VERMELHO = "C62828"
CINZA_TEXTO = "5B6576"
CINZA_CLARO = "F2F4F7"
CINZA_BORDA = "D8DCE3"
BRANCO = "FFFFFF"
AZUL_CLARO_TEXTO = "C9D6EA"   # subtítulo do cabeçalho, texto claro sobre navy
AZUL_CLARO_MARCA = "4FA3E0"   # "GOVERNO DO ESTADO DA BAHIA" no cabeçalho
PENDENCIA_FUNDO = "FDECEC"
COR_NAO_DISCRIMINADO = "8C8C8C"

# Cores semânticas por fonte de recurso — IDÊNTICAS em todas as saídas
# (painel, apresentação, resumo de ações, relatório).
COR_OGU = AZUL
COR_FINANCIAMENTO = TEAL
COR_ESTADUAL = LARANJA

# Tipografia e forma.
FONTE_CSS = "'Roboto','Segoe UI',Arial,sans-serif"
# Mesma fonte do PPTX de referência (gerar_pptx_pac.py: FONTE = "Roboto").
# Sem um peso "Black" equivalente de Roboto estabelecido no projeto,
# FONTE_PPTX_BLACK é a mesma Roboto — quem precisar de destaque usa
# negrito (fonte.bold), não uma família de fonte à parte.
FONTE_PPTX = "Roboto"
FONTE_PPTX_BLACK = "Roboto"
RAIO_CARD = 10
RAIO_CHIP = 12
RAIO_BOTAO_ICONE = "50%"
SOMBRA_HOVER = "0 4px 14px rgba(11,31,63,.15)"
SOMBRA_MODAL = "0 20px 60px rgba(0,0,0,.35)"

# ---------------------------------------------------------------------------
# Escala categórica (muitas categorias distintas, ex.: os ~14 status do
# relatório de MONITORAMENTO) — derivada dos tons acima (não é uma paleta
# solta nova): tons escuros/claros de cada matiz de referência, em vez de
# cores arbitrárias. Da entrega mais consolidada (verde) à mais incipiente
# (navy claro) — pedido do usuário em 14/11/2026.
CORES_CATEGORICAS = [
    "1B5E20",  # verde escuro — concluída/entregue
    VERDE,     # verde — inaugurada
    AZUL,      # azul — em andamento
    "5B8DB8",  # azul claro — emissão de O.S.
    LARANJA,   # laranja — em licitação/contratação
    "C9A227",  # dourado — processamento de licitação
    TEAL,      # teal — autorizado
    "7FB3AC",  # teal claro — aguardando autorizo
    "6D4C41",  # marrom — desenvolvimento de projetos
    "8E24AA",  # roxo — em estudo
    "C77DC9",  # roxo claro — verificação de viabilidade
    VERMELHO,  # vermelho — vencida/atrasada
    CINZA_TEXTO,   # cinza — cancelada
    "AEB4BD",  # cinza claro — sem informação
]


def hex_css(hex6):
    """'0B1F3F' -> '#0B1F3F' (CSS/HTML)."""
    return f"#{hex6}"


def cor_reportlab(hex6):
    """'0B1F3F' -> reportlab.lib.colors.HexColor('#0B1F3F')."""
    from reportlab.lib import colors
    return colors.HexColor(hex_css(hex6))


def cor_pptx(hex6):
    """'0B1F3F' -> pptx.dml.color.RGBColor (python-pptx)."""
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(hex6)


def caminho_recurso(nome_arquivo):
    return os.path.join(PASTA_BASE, nome_arquivo)


def data_uri_logo(nome_arquivo):
    """Base64 do PNG tal como está (sem recomprimir) — usado pelos logos
    PAC.png/GOVERNO.png no cabeçalho de qualquer HTML. Diferente da
    recompressão agressiva usada pras páginas do Canva em gerar_pptx_pac.py
    (_data_uri_imagem): logo é pequeno e PRECISA manter a transparência do
    PNG, então não passa por conversão pra JPEG."""
    caminho = caminho_recurso(nome_arquivo)
    with open(caminho, "rb") as f:
        dados = f.read()
    mime = "image/png" if caminho.lower().endswith(".png") else "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(dados).decode('ascii')}"


# ---------------------------------------------------------------------------
# CSS compartilhado — variáveis + estilos-base (cabeçalho, card, botão,
# tabela, chip) que qualquer HTML do projeto pode injetar. Cada página ainda
# define seus próprios estilos específicos por cima disso; isto cobre só o
# que deve ser IGUAL em todo lugar.
# ---------------------------------------------------------------------------
CSS_TOKENS = f"""
  :root {{
    --ivp-navy: #{NAVY};
    --ivp-azul: #{AZUL};
    --ivp-teal: #{TEAL};
    --ivp-laranja: #{LARANJA};
    --ivp-verde: #{VERDE};
    --ivp-vermelho: #{VERMELHO};
    --ivp-cinza-texto: #{CINZA_TEXTO};
    --ivp-cinza-claro: #{CINZA_CLARO};
    --ivp-cinza-borda: #{CINZA_BORDA};
    --ivp-branco: #{BRANCO};
    --ivp-azul-claro-texto: #{AZUL_CLARO_TEXTO};
    --ivp-azul-claro-marca: #{AZUL_CLARO_MARCA};
    --ivp-pendencia-fundo: #{PENDENCIA_FUNDO};
    --ivp-nao-discriminado: #{COR_NAO_DISCRIMINADO};
    --ivp-raio-card: {RAIO_CARD}px;
    --ivp-raio-chip: {RAIO_CHIP}px;
    --ivp-sombra-hover: {SOMBRA_HOVER};
    --ivp-sombra-modal: {SOMBRA_MODAL};
    --ivp-fonte: {FONTE_CSS};
  }}
  .ivp-cabecalho {{ background:var(--ivp-navy); color:#fff; padding:14px 40px 8px; position:relative; }}
  .ivp-cabecalho .ivp-titulo {{ font-size:30px; font-weight:700; }}
  .ivp-cabecalho .ivp-subtitulo {{ font-size:14px; font-weight:700; color:var(--ivp-azul-claro-texto); margin-top:4px; }}
  .ivp-cabecalho .ivp-governo {{ position:absolute; top:14px; right:40px; text-align:right; font-weight:700;
    font-size:17px; color:var(--ivp-azul-claro-marca); line-height:1.7; }}
  .ivp-card {{ background:#fff; border:1px solid var(--ivp-cinza-borda); border-radius:var(--ivp-raio-card); }}
  .ivp-card:hover, .ivp-card:focus {{ border-color:var(--ivp-navy); box-shadow:var(--ivp-sombra-hover); outline:none; }}
  .ivp-chip {{ font-size:11px; font-weight:700; color:var(--ivp-cinza-texto); background:var(--ivp-cinza-claro);
    border-radius:var(--ivp-raio-chip); padding:3px 9px; text-transform:uppercase; }}
  .ivp-botao {{ font-weight:700; color:#fff; background:var(--ivp-azul); border:none; border-radius:8px;
    cursor:pointer; font-family:var(--ivp-fonte); padding:13px 24px; font-size:15px; }}
  .ivp-botao:hover {{ background:#0A4E92; }}
  .ivp-botao-icone {{ width:28px; height:28px; border-radius:50%; border:1px solid var(--ivp-cinza-borda);
    background:var(--ivp-cinza-claro); color:var(--ivp-navy); font-size:17px; font-weight:700; line-height:1;
    cursor:pointer; display:inline-flex; align-items:center; justify-content:center; padding:0;
    font-family:var(--ivp-fonte); flex-shrink:0; }}
  .ivp-botao-icone:hover {{ background:#E6E9EE; }}
  .ivp-rotulo-secao {{ font-size:18px; font-weight:700; color:var(--ivp-cinza-texto); text-transform:uppercase; }}
  .ivp-tabela {{ width:100%; border-collapse:collapse; font-family:var(--ivp-fonte); }}
  .ivp-tabela th {{ background:var(--ivp-navy); color:#fff; text-align:left; padding:8px 14px; font-size:14px; }}
  .ivp-tabela td {{ padding:6px 14px; border-bottom:1px solid var(--ivp-cinza-borda); font-size:14px; }}
  .ivp-tabela tr:nth-child(even) td {{ background:var(--ivp-cinza-claro); }}
"""
