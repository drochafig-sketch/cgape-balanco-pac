"""Importação automática da planilha "Ações do Cronograma.xlsx" — pedido do
usuário em 17/11/2026.

O site https://app.conder.ba.gov.br/cronogramaacoes/ (INFORMS/CONDER) monta
a planilha no próprio navegador (botão "Exportar para planilha") a partir do
serviço ArcGIS público abaixo, sem login. Este módulo faz a mesma coisa em
Python: consulta o serviço e grava a planilha com as MESMAS colunas, textos e
formatos do botão do site (ver gerarPlanilhaAcoes/criarObjAcao em
js/main.js do site), para que resumo_acoes.py e relatorio_monitoramento.py a
leiam sem mudança nenhuma.

Uso:
    python importar_cronograma.py              -> grava na pasta do programa
    python importar_cronograma.py destino.xlsx -> grava em outro lugar
ou, de dentro do programa, baixar_planilha_cronograma(caminho).

Se o site mudar as colunas da exportação, ajustar COLUNAS_GRID/MARCOS
conforme js/main.js (dicColsGrid e dadosMarcosList).
"""

import json
import math
import os
import sys
import tempfile
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

URL_CAMADA = ("https://mapas.conder.ba.gov.br/arcgis/rest/services/"
              "INTERVENCAO_OBRA/INTERVENCAO_OBRAS_GERAL/MapServer/0")
ARQUIVO = "Ações do Cronograma.xlsx"
# Bahia: UTC-3 o ano todo (o site converte as datas no fuso do navegador)
FUSO = timezone(timedelta(hours=-3))
LOTE = 2000

# (título da coluna, campo do ArcGIS, tipo) na ordem do site (dicColsGrid
# com incluirNaPlanilha, por "ordem")
COLUNAS_GRID = [
    ("Código", None, "codigo"),
    ("Ação", "INTERVENCAO", "txt"),
    ("Território Identidade", "TERRITORIO_IDENTIDADE", "txt"),
    ("Município", "MUNICIPIO", "txt"),
    ("Executor", "EXECUTOR", "txt"),
    ("Número do Contrato", "NUMERO_CONTRATO", "txt"),
    ("Programa de Transferência de Recursos", "PROGRAMA_TRANSFERENCIA_RECURSOS", "txt"),
    ("Contratada", "CONTRATADA", "txt"),
    ("Escopo", "ESCOPO", "txt"),
    ("Detalhamento Escopo", "DETALHAMENTO_ESCOPO", "txt"),
    ("Fotovoltaica", "FOTOVOLTAICA", "txt"),
    ("Objeto", "OBJETO", "txt"),
    ("Status", "STATUS", "txt"),
    ("Prazo Contratual", "PRAZO_INICIAL_CONTRATUAL", "txt"),
    ("Avanço", "AVANCO_FISICO", "txt"),
    ("Diagnóstico Energético", "PV_DIAGNOSTICO_ENERGETICO", "txt"),
    ("Projeto Executivo aprovado e Adequação SEs", "PV_PROJETO_EXECUTIVO", "txt"),
    ("Parecer Concessionária conexão mini ou micro geração", "PV_PARECER_ACESSO", "txt"),
    ("Emissão pedidos compra materiais", "PV_PEDIDOS_COMPRA_MATERIAIS", "txt"),
    ("Recebimento materiais", "PV_RECEBIMENTO_MATERIAIS", "txt"),
    ("Execução obras", "PV_EXECUCAO_OBRAS", "txt"),
    ("Comissionamento", "PV_COMISSIONAMENTO", "txt"),
    ("Usina fotovoltaica incorporada", "PV_USINA_FOTOVOLTAICA", "txt"),
    ("Treinamentos e Manuais", "PV_TREINAMENTOS_MANUAIS", "txt"),
    ("Eixo", "EIXO", "txt"),
    ("Responsável Informação", "RESPONSAVEL_INFORMACAO", "txt"),
    ("Valor do Orçamento", "VALOR_ORCAMENTO", "moeda"),
    ("Valor de Investimento", "INVESTIMENTO", "moeda"),
    ("Última Atualização", "DATA_ULTIMA_ATUALIZACAO", "datahora"),
]
# marcos (dadosMarcosList, ordem cronológica típica): data em dd/mm/aaaa
MARCOS = [
    ("Conclusão dos Estudos", "PRV_CONCLUSAO_ESTUDO"),
    ("Conclusão do Projeto", "PRV_CONCLUSAO_DESENV_PROJETO"),
    ("Aviso de Licitação", "PRV_AVISO_LICITACAO"),
    ("Licitação / Contratação", "PRV_CONTRATACAO"),
    ("Ordem de Serviço", "PRV_EMISSAO_OS"),
    ("Entrega Obra", "PRV_CONCLUSAO_OBRA"),
    ("Instalação Equipamentos", "DATA_EQUIPARACAO"),
    ("Funcionalidade", "PRV_FUNCIONAMENTO"),
]
DERIVADAS = ["Dias em Atraso (De Marco)", "Avanço Ideal", "Faixa de Atraso (De Avanço)"]
FAIXAS_ATRASO_AVANCO = [5, 10, 20]   # cfgAlertasAvancoFisico.nivelAtrasoList


def _post(url, campos, tempo=120):
    dados = urllib.parse.urlencode(campos).encode("utf-8")
    req = urllib.request.Request(url, data=dados, headers={
        "Content-Type": "application/x-www-form-urlencoded",
        "User-Agent": "CGAPE-BalancoPAC/1.0"})
    with urllib.request.urlopen(req, timeout=tempo) as r:
        resposta = json.loads(r.read().decode("utf-8"))
    if "error" in resposta:
        raise RuntimeError(f"Serviço da CONDER respondeu erro: {resposta['error']}")
    return resposta


def buscar_registros():
    """Todas as ações da camada (atributos crus do ArcGIS), em lotes."""
    total = _post(URL_CAMADA + "/query", {"where": "1=1", "returnCountOnly": "true",
                                          "f": "json"})["count"]
    registros, inicio = [], 0
    while inicio < total:
        r = _post(URL_CAMADA + "/query", {
            "where": "1=1", "outFields": "*", "returnGeometry": "false",
            "orderByFields": "OBJECTID", "resultOffset": inicio,
            "resultRecordCount": LOTE, "f": "json"})
        lote = [f["attributes"] for f in r.get("features", [])]
        if not lote:
            break
        registros += lote
        inicio += len(lote)
    if len(registros) != total:
        raise RuntimeError(f"Leitura incompleta do serviço: {len(registros)} de {total} ações.")
    return registros


# ---- formatação igual à do site --------------------------------------------
def _vazio(v):
    # tratarNull: 0 fica; '', null, 'null', '<Null>' viram vazio
    if v == 0 and not isinstance(v, bool) and v is not None and not isinstance(v, str):
        return False
    return v is None or v == "" or v == "null" or v == "<Null>" or (
        isinstance(v, float) and math.isnan(v))


def _data_local(ms):
    return datetime.fromtimestamp(ms / 1000, FUSO).replace(tzinfo=None)


def _moeda(v):
    # toLocaleString('pt-br', {style:'currency', currency:'BRL'})
    txt = f"{abs(v):,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")
    return ("-" if v < 0 else "") + "R$ " + txt


def _numero(v):
    return int(v) if isinstance(v, float) and v.is_integer() else v


def _pct(v):
    # convertToPercent: parseFloat(v).toFixed(2) com vírgula
    try:
        return f"{float(v):.2f}".replace(".", ",") + "%"
    except (TypeError, ValueError):
        return "NaN%"


def montar_linhas(registros, agora=None):
    """Cabeçalho + linhas da planilha, como gerarPlanilhaAcoes do site."""
    agora = agora or datetime.now(FUSO).replace(tzinfo=None)
    cab = [c[0] for c in COLUNAS_GRID] + [m[0] for m in MARCOS] + DERIVADAS
    linhas = []
    for a in registros:
        linha = []
        for _titulo, campo, tipo in COLUNAS_GRID:
            if tipo == "codigo":
                resp = a.get("RESPONSAVEL_INFORMACAO")
                cod = a.get("CODIGO_ITEM")
                v = f"{'null' if resp is None else resp}-{'null' if cod is None else _numero(cod)}"
            else:
                v = a.get(campo)
                if tipo == "moeda":
                    v = _moeda(v) if isinstance(v, (int, float)) and not _vazio(v) else None
                elif tipo == "datahora":
                    v = _data_local(v) if isinstance(v, (int, float)) and v else None
                elif isinstance(v, (int, float)):
                    v = _numero(v)
            linha.append(None if v is not None and not isinstance(v, datetime) and _vazio(v) else v)
        for _nome, campo in MARCOS:
            v = a.get(campo)
            linha.append(f"{_data_local(v):%d/%m/%Y}" if isinstance(v, (int, float)) and v else None)
        # Dias em atraso (de marco): fim do status (23:59:59) já passou
        fim = a.get("DATA_FIM_STATUS")
        dias = None
        if isinstance(fim, (int, float)) and fim:
            d_fim = _data_local(fim).replace(hour=23, minute=59, second=59, microsecond=0)
            if agora > d_fim:
                dias = f"{math.ceil((agora - d_fim).total_seconds() / 86400)} dias"
        linha.append(dias)
        # Avanço ideal e faixa de atraso de avanço (só com avanço físico)
        fisico, ideal = a.get("AVANCO_FISICO"), a.get("AVANCO_IDEAL")
        if isinstance(fisico, (int, float)) and not (isinstance(fisico, float) and math.isnan(fisico)):
            ideal_n = 0 if ideal is None else ideal          # '' >= 0 no JS
            linha.append(_pct(ideal) if ideal_n >= 0 else "Sem informação")
            diferenca = ideal_n - fisico
            faixa = None
            for minimo in FAIXAS_ATRASO_AVANCO:
                if diferenca >= minimo:
                    faixa = minimo
            linha.append(f"Superior a {faixa}%" if faixa is not None else None)
        else:
            linha += [None, None]
        linhas.append(linha)
    return cab, linhas


def gravar_xlsx(cab, linhas, caminho):
    """Grava no mesmo formato do site (aba "Dados", cabeçalho verde)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    wb.properties.title = "Ações do Cronograma"
    wb.properties.creator = "INFORMS - CONDER"
    ws.append(cab)
    fino = Side(style="thin")
    for cel in ws[1]:
        cel.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
        cel.fill = PatternFill("solid", fgColor="FF127622")
        cel.alignment = Alignment(horizontal="center")
        cel.border = Border(top=fino, right=fino, bottom=fino, left=fino)
    col_data = cab.index("Última Atualização") + 1
    for linha in linhas:
        ws.append(linha)
        cel = ws.cell(row=ws.max_row, column=col_data)
        if cel.value is not None:
            cel.number_format = "mm-dd-yy"
    wb.save(caminho)


def baixar_planilha_cronograma(caminho):
    """Baixa do serviço da CONDER e grava `caminho` (troca atômica: o arquivo
    antigo só é substituído se tudo deu certo). Devolve um resumo."""
    registros = buscar_registros()
    if not registros:
        raise RuntimeError("O serviço da CONDER não devolveu nenhuma ação.")
    cab, linhas = montar_linhas(registros)
    pasta = os.path.dirname(os.path.abspath(caminho))
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=pasta)
    os.close(fd)
    try:
        gravar_xlsx(cab, linhas, tmp)
        os.replace(tmp, caminho)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    return {"acoes": len(linhas), "caminho": caminho, "quando": datetime.now()}


def _valores(caminho):
    from openpyxl import load_workbook
    # cabeçalho + linhas em ordem fixa: a ordem das linhas não conta (o site
    # exporta na ordem da tela; aqui sai por OBJECTID)
    wb = load_workbook(caminho, read_only=True)
    try:
        linhas = [tuple(None if v == "" else v for v in linha)
                  for linha in wb.active.iter_rows(values_only=True)]
    finally:
        wb.close()
    return linhas[:1] + sorted(linhas[1:], key=repr)


def atualizar_se_mudou(caminho):
    """Baixa e só troca `caminho` se o CONTEÚDO mudou (o .xlsx muda de bytes
    a cada gravação, então a comparação é célula a célula) — usado pelo
    vigia (VIGIAR_GITHUB.ps1), para não gerar commit sem mudança."""
    cab, linhas = montar_linhas(buscar_registros())
    if not linhas:
        raise RuntimeError("O serviço da CONDER não devolveu nenhuma ação.")
    pasta = os.path.dirname(os.path.abspath(caminho))
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=pasta)
    os.close(fd)
    try:
        gravar_xlsx(cab, linhas, tmp)
        if os.path.exists(caminho) and _valores(tmp) == _valores(caminho):
            os.remove(tmp)
            return {"mudou": False, "acoes": len(linhas)}
        os.replace(tmp, caminho)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    return {"mudou": True, "acoes": len(linhas)}


if __name__ == "__main__":
    # python importar_cronograma.py [destino.xlsx]          -> baixa e grava
    # python importar_cronograma.py --se-mudou [destino]    -> vigia: saída
    #   "SEM_MUDANCA" (código 0), "ATUALIZADA" (código 3) ou erro (código 1)
    args = sys.argv[1:]
    se_mudou = "--se-mudou" in args
    args = [a for a in args if a != "--se-mudou"]
    destino = args[0] if args else os.path.join(os.path.dirname(os.path.abspath(__file__)), ARQUIVO)
    try:
        if se_mudou:
            r = atualizar_se_mudou(destino)
            print(("ATUALIZADA" if r["mudou"] else "SEM_MUDANCA") + f" {r['acoes']} acoes")
            sys.exit(3 if r["mudou"] else 0)
        r = baixar_planilha_cronograma(destino)
        print(f"{r['acoes']} ações gravadas em {r['caminho']}")
    except SystemExit:
        raise
    except Exception as exc:
        print(f"ERRO {exc}")
        sys.exit(1)
