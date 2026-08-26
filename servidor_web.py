"""Servidor web do painel PAC — abre no navegador, sem instalar nada.

Serve o MESMO painel de filtros HTML/CSS/JS que o app desktop
("CGAPE - BALANÇO PAC.py") já usa, só que por HTTP em vez de numa janela
nativa (pywebview). Toda a lógica de negócio (ETL da planilha, filtros,
geração de PDF) continua morando no arquivo original — este servidor só
importa esse arquivo como módulo e expõe as mesmas operações da ponte
"APIFiltros" como rotas Flask.

Rodar localmente:
    PAC_WEB_MODE=1 python servidor_web.py
    (no Windows/PowerShell: $env:PAC_WEB_MODE=1; python servidor_web.py)
    depois abrir http://localhost:5000

Em produção (Render, ver render.yaml): gunicorn servidor_web:app
"""

import importlib.util
import os
import sys
import tempfile

# Precisa ser definida ANTES de importar o arquivo original: ele lê
# PAC_WEB_MODE == "1" uma única vez, na primeira linha executada do módulo,
# pra decidir se pode chamar os.startfile() (só faz sentido numa máquina
# desktop) ou mostrar diálogos do tkinter em erros fatais de inicialização.
os.environ.setdefault("PAC_WEB_MODE", "1")

from flask import Flask, Response, jsonify, request

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
CAMINHO_SCRIPT_ORIGINAL = os.path.join(PASTA_BASE, "CGAPE - BALANÇO PAC.py")

# O nome do arquivo original tem espaço e acento — não dá para usar
# "import" normal. importlib.util carrega qualquer caminho de arquivo como
# módulo Python. Isso executa o arquivo inteiro (inclusive o carregamento e
# o ETL da planilha, que roda no nível do módulo) uma única vez, na
# inicialização do servidor — todas as requisições depois reaproveitam o
# mesmo `pac.df_original` já pronto em memória, sem reler a planilha a cada
# clique.
_spec = importlib.util.spec_from_file_location("cgape_balanco_pac", CAMINHO_SCRIPT_ORIGINAL)
pac = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = pac
_spec.loader.exec_module(pac)

app = Flask(__name__)


@app.route("/")
def painel():
    # Mesmo HTML/CSS/JS do painel desktop — o helper JS embutido nele
    # (chamarAPI/baixarPDF, ver montar_html_painel) detecta sozinho que não
    # existe "window.pywebview" aqui e passa a falar com as rotas /api/*
    # abaixo em vez da ponte nativa.
    return pac.montar_html_painel(pac.df_original)


def _argumentos_da_requisicao():
    # O JS do painel sempre manda a lista de argumentos como um array JSON
    # no corpo do POST (ver chamarAPI/baixarPDF no HTML) — corresponde 1:1
    # aos argumentos posicionais que cada método de APIFiltros já recebia
    # no modo desktop.
    corpo = request.get_json(force=True, silent=True)
    return corpo if isinstance(corpo, list) else []


def _rota_json(funcao):
    # Empacota uma função _api_* (já devolve um dict JSON-serializável, com
    # seu próprio try/except interno) como uma rota Flask. O try/except
    # aqui é só uma rede de segurança extra para erros de transporte (JSON
    # inválido no corpo, por exemplo) que a função em si nunca veria.
    def rota():
        try:
            resultado = funcao(*_argumentos_da_requisicao())
        except Exception as erro:
            return jsonify({"ok": False, "erro": str(erro)}), 500
        return jsonify(resultado)

    rota.__name__ = f"rota_{funcao.__name__}"
    return rota


def _rota_pdf(funcao):
    # Mesma ideia de _rota_json, mas para as duas operações que geram PDF:
    # quando dá certo, a resposta É o PDF (Content-Type: application/pdf +
    # Content-Disposition com o nome sugerido) — quando não, volta um JSON
    # {"ok": False, ...} normal, no mesmo formato de sempre. O baixarPDF()
    # do lado do JS decide o que fazer olhando o Content-Type da resposta.
    def rota():
        try:
            resultado = funcao(*_argumentos_da_requisicao())
        except Exception as erro:
            return jsonify({"ok": False, "erro": str(erro)}), 500

        if not resultado.get("ok"):
            return jsonify(resultado)

        resposta = Response(resultado["conteudo"], mimetype="application/pdf")
        nome_arquivo = resultado.get("nome_arquivo", "relatorio.pdf")
        resposta.headers["Content-Disposition"] = f'inline; filename="{nome_arquivo}"'
        return resposta

    rota.__name__ = f"rota_{funcao.__name__}"
    return rota


# Uma rota por método que a ponte APIFiltros já expunha no desktop (mesmos
# nomes — o JS do painel chama "/api/<nome do método>" sem precisar saber
# a diferença). compartilhar_whatsapp, salvar_visualizacao_html e
# fechar_por_inatividade ficaram só no lado do JS (ver montar_html_painel):
# são operações puramente de navegador (abrir link, criar um Blob local) ou
# exclusivas de janela desktop, sem contrapartida útil num servidor.
app.add_url_rule(
    "/api/pre_visualizar", view_func=_rota_json(pac._api_pre_visualizar), methods=["POST"]
)
app.add_url_rule(
    "/api/mapa_mental", view_func=_rota_json(pac._api_mapa_mental), methods=["POST"]
)
app.add_url_rule(
    "/api/iniciar_geracao", view_func=_rota_json(pac._api_iniciar_geracao), methods=["POST"]
)
app.add_url_rule(
    "/api/listar_paginas_relatorio",
    view_func=_rota_json(pac._api_listar_paginas_relatorio),
    methods=["POST"],
)
app.add_url_rule(
    "/api/buscar_ficha_acao", view_func=_rota_json(pac._api_buscar_ficha_acao), methods=["POST"]
)
app.add_url_rule(
    "/api/buscar_ficha_por_item",
    view_func=_rota_json(pac._api_buscar_ficha_por_item),
    methods=["POST"],
)
app.add_url_rule(
    "/api/escolher_local_e_gerar",
    view_func=_rota_pdf(pac._api_gerar_pdf_bytes),
    methods=["POST"],
)
app.add_url_rule(
    "/api/exportar_ficha_pdf",
    view_func=_rota_pdf(pac._api_exportar_ficha_pdf_bytes),
    methods=["POST"],
)


@app.route("/saude")
def saude():
    # Endpoint simples para checagem de "o servidor está de pé e a planilha
    # carregou" — útil pra healthcheck do Render e pra depuração manual.
    return jsonify(
        {
            "ok": True,
            "linhas_carregadas": int(len(pac.df_original)),
            "atualizado_em": pac.ultima_atualizacao.isoformat(),
        }
    )


if __name__ == "__main__":
    # Uso local (desenvolvimento/teste) — em produção quem sobe o app é o
    # gunicorn (ver render.yaml: "gunicorn servidor_web:app"), não este
    # bloco.
    porta = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=porta, debug=False)
