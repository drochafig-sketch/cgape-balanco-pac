"""Diário Oficial do Estado (DOE) no CGAPE — pedido do usuário em 26/11/2026.

O mesmo recurso DOE do projeto RELATÓRIO SSP, para enxergar as publicações de
TODAS as secretarias (não só da Segurança Pública), sem o alerta de
exoneração de uma pessoa específica, e com uma caixa de seleção múltipla de
secretarias que filtra o detalhamento do painel. Só para os usuários com
login (acesso.py, seção "DOE").

  - Pasta de dados na rede (para os três usuários verem o mesmo): PASTA_DOE,
    com os doe_AAAA-MM-DD[_suplemento]_completo.md, o manual.json (correções
    à mão herdadas do RELATÓRIO SSP) e o doe_dados.json lido pelo painel.
  - ATUALIZAR DOE: baixa do site oficial (dool.egba.ba.gov.br) as edições
    dos últimos dias que faltam na pasta e refaz a análise
    (doe_analisar.executar). O download é o mesmo do RELATÓRIO SSP.
  - HTML_JANELA: janela #doe-overlay (aberta pelo cartão da HOME e pelo
    botão DOE no cabeçalho dos dashes do PAC, MONITORAMENTO e PRAZOS).
"""
from __future__ import annotations

import json
import os
import re
import threading

PASTA_PADRAO = r"X:\COAPI\NOTAS TÉCNICAS COAPI\2024-2026\PANORAMA\2026\DOE"
ARQUIVO_DADOS = "doe_dados.json"
ARQUIVO_MANUAL = "manual.json"
DIAS_ATUALIZAR = 7
DOE_SITE = "https://dool.egba.ba.gov.br"


def pasta_doe() -> str:
    return os.environ.get("CGAPE_DOE_PASTA") or PASTA_PADRAO


def caminho_dados() -> str:
    return os.path.join(pasta_doe(), ARQUIVO_DADOS)


# ---------------------------------------------------------------- leitura

_CACHE = {"mtime": None, "dados": None}


def _tem_diarios(pasta):
    try:
        return any(n.startswith("doe_") and n.endswith(".md") for n in os.listdir(pasta))
    except OSError:
        return False


def analisar():
    """Refaz o doe_dados.json a partir dos .md da pasta."""
    import doe_analisar
    pasta = pasta_doe()
    return doe_analisar.executar(pasta, caminho_dados(), os.path.join(pasta, ARQUIVO_MANUAL))


def carregar():
    """Dados do painel (cache pela data do arquivo) ou None."""
    caminho = caminho_dados()
    try:
        mtime = os.path.getmtime(caminho)
    except OSError:
        if _tem_diarios(pasta_doe()):
            analisar()
            return carregar()
        return None
    if _CACHE["mtime"] != mtime:
        with open(caminho, encoding="utf-8") as f:
            _CACHE["dados"] = json.load(f)
        _CACHE["mtime"] = mtime
    return _CACHE["dados"]


def api_dados():
    try:
        if not os.path.isdir(pasta_doe()):
            return {"ok": False, "mensagem": f"Pasta do DOE não encontrada: {pasta_doe()} (rede desconectada?)."}
        dados = carregar()
        if not dados:
            return {"ok": False, "mensagem": "Ainda não há diários na pasta do DOE. Use ATUALIZAR DOE para baixar os "
                                             f"últimos {DIAS_ATUALIZAR} dias do site oficial."}
        return {"ok": True, "doe": dados, "pasta": pasta_doe()}
    except Exception as exc:
        return {"ok": False, "mensagem": f"Erro ao ler os dados do DOE: {exc}"}


# ---------------------------------------------------------------- atualização

_TRAVA = threading.Lock()


def _sem_secoes(caminho):
    """True se o .md não tem nenhum título de órgão (<mark>): conversão do
    PDF em texto corrido, que a análise não consegue ler."""
    try:
        with open(caminho, encoding="utf-8", errors="replace") as f:
            return "<mark>" not in f.read()
    except OSError:
        return False


def api_atualizar(dias=None):
    """Baixa do site oficial as edições dos últimos `dias` dias que ainda não
    estão na pasta e, havendo novidade, refaz a análise. Devolve
    {"ok", "mensagem", "doe"}."""
    from datetime import date, timedelta
    dias = int(dias or DIAS_ATUALIZAR)
    if not _TRAVA.acquire(blocking=False):
        return {"ok": False, "mensagem": "Já existe uma atualização do DOE em andamento."}
    try:
        pasta = pasta_doe()
        try:
            os.makedirs(pasta, exist_ok=True)
        except OSError as exc:
            return {"ok": False, "mensagem": f"Pasta do DOE inacessível ({pasta}): {exc}"}
        novos, erros, sem = [], [], []
        hoje = date.today()
        for n in range(dias - 1, -1, -1):
            dia = hoje - timedelta(n)
            try:
                resp = json.loads(_doe_texto(f"{DOE_SITE}/apifront/portal/edicoes/edicoes_from_data/{dia.isoformat()}"))
                edicoes = [] if resp.get("erro") else resp.get("itens", [])
            except Exception as exc:
                erros.append(f"{dia:%d/%m}: falha ao consultar o site ({exc})")
                continue
            if not edicoes:
                sem.append(f"{dia:%d/%m}")
            for ed in edicoes:
                destino = os.path.join(pasta, _doe_nome_arquivo(ed))
                if os.path.exists(destino) and not _sem_secoes(destino):
                    continue
                try:
                    md = _doe_montar_md(ed)
                    if os.path.exists(destino):
                        os.replace(destino, destino + ".bak")
                    tmp = destino + ".tmp"
                    with open(tmp, "w", encoding="utf-8") as f:
                        f.write(md)
                    os.replace(tmp, destino)
                    novos.append(ed["data"][:5] + (" (suplemento)" if _doe_suplemento(ed) else ""))
                except Exception as exc:
                    erros.append(f"{ed['data'][:5]}: {exc}")
        partes = []
        if novos or not os.path.exists(caminho_dados()):
            try:
                analisar()
            except Exception as exc:
                erros.append(f"a análise falhou: {exc}")
        if novos:
            partes.append(f"Diários baixados do site: {', '.join(novos)}.")
        else:
            partes.append(f"Nenhum diário novo nos últimos {dias} dias — a pasta já estava em dia.")
        if sem:
            partes.append(f"Sem edição no site: {', '.join(sem)}.")
        if erros:
            partes.append("Problemas: " + "; ".join(erros))
        return {"ok": not erros, "mensagem": "\n".join(partes), "doe": carregar()}
    finally:
        _TRAVA.release()


# ---------------------------------------------------------------- download (do RELATÓRIO SSP)
# Usa a versão HTML pública do diário (sem login): lista de edições do dia,
# sumário com a árvore de órgãos e o texto de cada matéria. Grava um
# doe_AAAA-MM-DD[_suplemento]_completo.md por edição. Não são endpoints
# documentados: se a EGBA mudar o site, o botão avisa o erro.

_DOE_UA = "Mozilla/5.0 (CGAPE - monitoramento DOE)"
_DOE_PARALELO = 6   # downloads simultâneos de matérias (gentil com o site)

# O servidor do DOE envia o intermediário errado (RapidSSL TLS RSA CA G1) em
# vez do que assinou o certificado (GeoTrust TLS RSA CA G1). Navegadores e o
# Windows completam a cadeia sozinhos; o Python não, e dá
# CERTIFICATE_VERIFY_FAILED. Juntamos aqui o intermediário certo (emitido pela
# DigiCert Global Root G2, válido até 02/11/2027) às raízes do sistema — a
# verificação continua completa, só a cadeia fica correta.
_DOE_INTERMEDIARIO = """-----BEGIN CERTIFICATE-----
MIIEjTCCA3WgAwIBAgIQDQd4KhM/xvmlcpbhMf/ReTANBgkqhkiG9w0BAQsFADBhMQswCQYDVQQG
EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSAw
HgYDVQQDExdEaWdpQ2VydCBHbG9iYWwgUm9vdCBHMjAeFw0xNzExMDIxMjIzMzdaFw0yNzExMDIx
MjIzMzdaMGAxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
dy5kaWdpY2VydC5jb20xHzAdBgNVBAMTFkdlb1RydXN0IFRMUyBSU0EgQ0EgRzEwggEiMA0GCSqG
SIb3DQEBAQUAA4IBDwAwggEKAoIBAQC+F+jsvikKy/65LWEx/TMkCDIuWegh1Ngwvm4QyISgP7oU
5d79eoySG3vOhC3w/3jEMuipoH1fBtp7m0tTpsYbAhch4XA7rfuD6whUgajeErLVxoiWMPkC/DnU
vbgi74BJmdBiuGHQSd7LwsuXpTEGG9fYXcbTVN5SATYqDfbexbYxTMwVJWoVb6lrBEgM3gBBqiiA
iy800xu1Nq07JdCIQkBsNpFtZbIZhsDSfzlGWP4wEmBQ3O67c+ZXkFr2DcrXBEtHam80Gp2SNhou
2U5U7UesDL/xgLK6/0d76TnEVMSUVJkZ8VeZr+IUIlvoLrtjLbqugb0T3OYXW+CQU0kBAgMBAAGj
ggFAMIIBPDAdBgNVHQ4EFgQUlE/UXYvkpOKmgP792PkA76O+AlcwHwYDVR0jBBgwFoAUTiJUIBiV
5uNu5g/6+rkS7QYXjzkwDgYDVR0PAQH/BAQDAgGGMB0GA1UdJQQWMBQGCCsGAQUFBwMBBggrBgEF
BQcDAjASBgNVHRMBAf8ECDAGAQH/AgEAMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
cDovL29jc3AuZGlnaWNlcnQuY29tMEIGA1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmwzLmRpZ2lj
ZXJ0LmNvbS9EaWdpQ2VydEdsb2JhbFJvb3RHMi5jcmwwPQYDVR0gBDYwNDAyBgRVHSAAMCowKAYI
KwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwDQYJKoZIhvcNAQELBQADggEB
AIIcBDqC6cWpyGUSXAjjAcYwsK4iiGF7KweG97i1RJz1kwZhRoo6orU1JtBYnjzBc4+/sXmnHJk3
mlPyL1xuIAt9sMeC7+vreRIF5wFBC0MCN5sbHwhNN1JzKbifNeP5ozpZdQFmkCo+neBiKR6HqIA+
LMTMCMMuv2khGGuPHmtDze4GmEGZtYLyF8EQpa5YjPuV6k2Cr/N3XxFpT3hRpt/3usU/Zb9wfKPt
WpoznZ4/44c1p9rzFcZYrWkj3A+7TNBJE0GmP2fhXhP1D/XVfIW/h0yCJGEiV9Glm/uGOa3DXHlm
bAcxSyCRraG+ZBkA7h4SeM6Y8l/7MBRpPCz6l8Y=
-----END CERTIFICATE-----
"""
_doe_ctx_cache = []


def _doe_contexto_ssl():
    """Contexto TLS com as raízes do sistema + o intermediário do DOE."""
    if not _doe_ctx_cache:
        import ssl
        ctx = ssl.create_default_context()
        try:
            ctx.load_verify_locations(cadata=_DOE_INTERMEDIARIO)
        except ssl.SSLError:
            pass
        _doe_ctx_cache.append(ctx)
    return _doe_ctx_cache[0]


def _doe_baixar(url, tentativas=3):
    """(bytes, content-type), com novas tentativas em falha de rede."""
    import time
    import urllib.request
    for n in range(tentativas):
        try:
            req = urllib.request.Request(url,
                                         headers={"User-Agent": _DOE_UA})
            with urllib.request.urlopen(req, timeout=60,
                                        context=_doe_contexto_ssl()) as r:
                return r.read(), r.headers.get("Content-Type") or ""
        except Exception:
            if n == tentativas - 1:
                raise
            time.sleep(2 * (n + 1))


def _doe_texto(url):
    return _doe_baixar(url)[0].decode("utf-8", errors="replace")


def _doe_maiuscula(s):
    letras = [c for c in s if c.isalpha()]
    return bool(letras) and all(c.isupper() for c in letras)


def _doe_suplemento(edicao):
    # a API devolve 0/1 ou o texto "Suplemento", conforme a edição
    return str(edicao.get("suplemento") or "0").strip() not in ("0", "")


def _doe_nome_arquivo(edicao):
    d, m, a = edicao["data"].split("/")
    sup = "_suplemento" if _doe_suplemento(edicao) else ""
    return f"doe_{a}-{m}-{d}{sup}_completo.md"


def _doe_sumario(id_edicao):
    """[(caminho_de_pastas, [[nome, tem_subpasta], ...], id, título)] lidos
    de /html/<id>.html (árvore de órgãos e categorias do diário)."""
    from html.parser import HTMLParser

    class Sumario(HTMLParser):
        def __init__(self):
            super().__init__()
            self.pilha, self.niveis_li, self.materias = [], [], []
            self.em_pasta, self.em_link, self.texto = False, None, ""

        def handle_starttag(self, tag, attrs):
            a = dict(attrs)
            if tag == "span" and a.get("class") == "folder":
                self.em_pasta, self.texto = True, ""
            elif tag == "a" and "linkMateria" in (a.get("class") or ""):
                self.em_link = (a.get("data-materia-id")
                                or a.get("identificador"))
                self.texto = ""
            elif tag == "li":
                self.niveis_li.append(len(self.pilha))

        def handle_endtag(self, tag):
            if tag == "span" and self.em_pasta:
                self.em_pasta = False
                if self.pilha:
                    self.pilha[-1][1] = True     # a pasta-mãe tem subpasta
                self.pilha.append([self.texto.strip(), False])
            elif tag == "a" and self.em_link:
                titulo = re.sub(r"^\s*#\d+\s*-\s*", "", self.texto).strip()
                self.materias.append((tuple(p[0] for p in self.pilha),
                                      list(self.pilha), self.em_link,
                                      titulo))
                self.em_link = None
            elif tag == "li" and self.niveis_li:
                del self.pilha[self.niveis_li.pop():]

        def handle_data(self, d):
            if self.em_pasta or self.em_link:
                self.texto += d

    sm = Sumario()
    sm.feed(_doe_texto(f"{DOE_SITE}/html/{id_edicao}.html"))
    return sm.materias


def _doe_html_para_md(texto_html):
    """HTML da matéria (exportado do Word) -> linhas de markdown:
    parágrafos separados por linha em branco, negrito como **, tabelas
    como |a|b| — o mesmo que o analisar.py encontrava no PDF convertido."""
    from html.parser import HTMLParser

    class Materia(HTMLParser):
        BLOCO = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li"}

        def __init__(self):
            super().__init__()
            self.saida, self.par, self.ignorar = [], "", 0
            self.tabela, self.celula = None, None

        def _fecha_par(self):
            t = re.sub(r"[ \t\r\n\xa0]+", " ", self.par).strip()
            t = re.sub(r"\*\*\s*\*\*", " ", t).strip()
            if t and t != "**":
                self.saida += [t, ""]
            self.par = ""

        def _alvo(self, s):
            if self.celula is not None:
                self.celula += s
            else:
                self.par += s

        def handle_starttag(self, tag, attrs):
            if tag in ("style", "script", "head", "title"):
                self.ignorar += 1
            elif tag == "table":
                self._fecha_par()
                self.tabela = []
            elif tag == "tr" and self.tabela is not None:
                self.tabela.append([])
            elif tag in ("td", "th") and self.tabela is not None:
                self.celula = ""
            elif tag in ("b", "strong"):
                self._alvo(" **")
            elif tag == "br":
                self._alvo(" " if self.celula is not None else "\n")
            elif tag in self.BLOCO and self.celula is None:
                self._fecha_par()
            elif tag in self.BLOCO:
                self.celula += " "

        def handle_endtag(self, tag):
            if tag in ("style", "script", "head", "title"):
                self.ignorar = max(0, self.ignorar - 1)
            elif tag in ("b", "strong"):
                self._alvo("** ")
            elif tag in ("td", "th") and self.celula is not None:
                c = re.sub(r"\s+", " ",
                           self.celula.replace("**", "")).strip()
                if self.tabela:
                    self.tabela[-1].append(c.replace("|", "/"))
                self.celula = None
            elif tag == "table" and self.tabela is not None:
                linhas = [l for l in self.tabela if any(l)]
                if linhas:
                    n = max(len(l) for l in linhas)
                    for i, l in enumerate(linhas):
                        l = l + [""] * (n - len(l))
                        self.saida.append("|" + "|".join(l) + "|")
                        if i == 0:
                            self.saida.append("|" + "|".join(["---"] * n)
                                              + "|")
                    self.saida.append("")
                self.tabela = None
            elif tag in self.BLOCO and self.celula is None:
                self._fecha_par()

        def handle_data(self, d):
            if not self.ignorar:
                self._alvo(d)

    p = Materia()
    try:
        p.feed(texto_html)
        p._fecha_par()
    except Exception:
        pass   # HTML malformado: fica com o que já foi lido
    return p.saida


def _doe_pdf_para_md(dados):
    """Matéria publicada como PDF anexado: mesma conversão do fluxo antigo
    (pymupdf4llm); na falta dela, texto puro do pymupdf."""
    try:
        import pymupdf
        doc = pymupdf.open(stream=dados, filetype="pdf")
    except Exception:
        return []
    try:
        import pymupdf4llm
        md = pymupdf4llm.to_markdown(doc)
    except Exception:
        md = "\n\n".join(pg.get_text() for pg in doc)
    return md.splitlines() + [""]


def _doe_materia(id_materia):
    """Linhas de markdown da matéria; [] se não der para ler (o título do
    sumário entra no lugar)."""
    try:
        dados, tipo = _doe_baixar(f"{DOE_SITE}/apifront/portal/edicoes/"
                                  f"publicacoes_ver_conteudo/{id_materia}")
    except Exception as exc:
        print(f"  [aviso] matéria {id_materia}: {exc}")
        return []
    if "pdf" in tipo.lower() or dados[:5] == b"%PDF-":
        return _doe_pdf_para_md(dados)
    return _doe_html_para_md(dados.decode("utf-8", errors="replace"))


def _doe_montar_md(edicao):
    """Baixa sumário + matérias de uma edição e devolve o markdown."""
    from concurrent.futures import ThreadPoolExecutor
    materias = _doe_sumario(edicao["id"])
    with ThreadPoolExecutor(_DOE_PARALELO) as ex:
        textos = list(ex.map(_doe_materia, [m[2] for m in materias]))

    o = [f"SALVADOR, {edicao['data']} - Nº {edicao['numero']}"
         f"{' - SUPLEMENTO' if _doe_suplemento(edicao) else ''}", ""]
    anterior = ()
    for (caminho, pastas, _id, titulo), linhas in zip(materias, textos):
        # Títulos de seção como no PDF: só as pastas que são órgãos. O 1º
        # nível (EXECUTIVO, LICITAÇÕES...) não entra; pastas-folha em caixa
        # mista (Portarias, Resumo De Contrato...) são categorias e também
        # não. Secretarias (MAIÚSCULAS) abrem seção, como no PDF.
        comum = 0
        while (comum < min(len(caminho), len(anterior))
               and caminho[comum] == anterior[comum]):
            comum += 1
        for nivel in range(max(comum, 1), len(caminho)):
            nome, tem_sub = pastas[nivel]
            if not tem_sub and not _doe_maiuscula(nome):
                continue
            marca = "###" if _doe_maiuscula(nome) else "#####"
            o += [f"{marca} **<mark>{nome}</mark>**", ""]
        anterior = caminho
        o += linhas or [titulo, ""]
    return "\n".join(o) + "\n"



# ---------------------------------------------------------------- janela do painel
# Injetada no painel (montar_html_painel) só para quem tem a seção "DOE".
# Tokens de cor do painel (--cor-card etc.), tema claro/escuro incluso.
HTML_JANELA = r"""
<div id="doe-overlay" aria-hidden="true">
  <div id="doe-janela" role="dialog" aria-label="Diário Oficial do Estado">
    <div class="doe-cabecalho">
      <div class="doe-titulo">DIÁRIO OFICIAL DO ESTADO (DOE)
        <small id="doeMeta">Publicações de todas as secretarias</small></div>
      <div class="doe-botoes">
        <button type="button" id="doeAtualizar" title="Baixa do site oficial (dool.egba.ba.gov.br) os diários dos últimos 7 dias que ainda não estão na pasta e refaz a análise">ATUALIZAR DOE</button>
        <button type="button" id="doeFechar" class="doe-fechar" title="Fechar (Esc)" aria-label="Fechar">×</button>
      </div>
    </div>
    <div class="doe-status" id="doeStatus" hidden></div>
    <div class="doe-controles">
      <label class="doe-ctl"><span>Diário</span><select id="doeDia"></select></label>
      <div class="doe-ctl doe-sec">
        <span>Secretarias</span>
        <button type="button" id="doeSecBtn" aria-haspopup="true" aria-expanded="false">Todas as secretarias</button>
        <div class="doe-sec-menu" id="doeSecMenu" hidden>
          <input type="search" id="doeSecBusca" placeholder="Procurar secretaria…" autocomplete="off">
          <div class="doe-sec-acoes"><button type="button" id="doeSecTodas">Marcar todas</button><button type="button" id="doeSecLimpar">Limpar (todas)</button></div>
          <div class="doe-sec-lista" id="doeSecLista"></div>
        </div>
      </div>
      <label class="doe-ctl doe-ctl-busca"><span>Buscar</span><input type="search" id="doeBusca" autocomplete="off"
        placeholder="Nome, cargo, objeto, empresa, processo, ementa…"></label>
    </div>
    <div class="doe-kpis" id="doeKpis"></div>
    <div class="doe-abas" id="doeAbas" role="tablist"></div>
    <div class="doe-conteudo" id="doeConteudo"><div class="doe-vazio">Carregando o Diário Oficial…</div></div>
  </div>
</div>
<style>
  #doe-overlay { display:none; position:fixed; inset:0; z-index:1100; background:rgba(5,12,24,.55);
    align-items:center; justify-content:center; padding:18px; font-family:var(--fonte, 'Segoe UI', Roboto, Arial, sans-serif); }
  #doe-overlay.aberto { display:flex; }
  #doe-janela { background:var(--cor-fundo); color:var(--cor-texto-primario); width:min(1500px, 100%); height:100%;
    border-radius:12px; display:flex; flex-direction:column; overflow:hidden; box-shadow:0 20px 60px rgba(0,0,0,.4);
    border-top:4px solid var(--cor-acento-teal); }
  .doe-cabecalho { display:flex; align-items:center; justify-content:space-between; gap:16px; padding:14px 20px;
    background:var(--cor-card); border-bottom:1px solid var(--cor-card-elevado); }
  .doe-titulo { font-size:18px; font-weight:900; letter-spacing:.3px; }
  .doe-titulo small { display:block; font-size:12px; font-weight:400; letter-spacing:0; color:var(--cor-texto-secundario); margin-top:2px; }
  .doe-botoes { display:flex; align-items:center; gap:10px; }
  #doeAtualizar { font:inherit; font-size:12.5px; font-weight:700; cursor:pointer; border-radius:8px; padding:8px 14px;
    background:transparent; color:var(--cor-acento-teal-texto); border:1px solid var(--cor-acento-teal); }
  #doeAtualizar:hover:not(:disabled) { background:var(--cor-acento-teal); color:#fff; }
  #doeAtualizar:disabled { opacity:.6; cursor:wait; }
  .doe-fechar { font-size:28px; line-height:1; border:none; background:none; cursor:pointer; color:var(--cor-texto-secundario); padding:0 4px; }
  .doe-fechar:hover { color:var(--cor-texto-primario); }
  .doe-status { margin:10px 20px 0; padding:10px 14px; border-radius:8px; font-size:12.5px; white-space:pre-line;
    background:var(--cor-card); border-left:4px solid var(--cor-acento-teal); }
  .doe-status.erro { border-left-color:#D64545; }
  .doe-controles { display:flex; flex-wrap:wrap; gap:12px; align-items:flex-end; padding:12px 20px 4px; }
  .doe-ctl { display:flex; flex-direction:column; gap:4px; font-size:11px; font-weight:700; text-transform:uppercase;
    letter-spacing:.4px; color:var(--cor-texto-secundario); position:relative; }
  .doe-ctl select, .doe-ctl input[type=search], #doeSecBtn { font:inherit; font-size:13px; font-weight:400; text-transform:none; letter-spacing:0;
    color:var(--cor-texto-primario); background:var(--cor-card); border:1px solid var(--cor-card-elevado); border-radius:8px;
    padding:7px 10px; min-width:220px; box-sizing:border-box; }
  .doe-ctl-busca { flex:1; min-width:260px; }
  .doe-ctl-busca input[type=search] { width:100%; }
  #doeSecBtn { cursor:pointer; text-align:left; min-width:320px; max-width:460px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  #doeSecBtn.ativo, .doe-ctl select:focus, .doe-ctl input[type=search]:focus { border-color:var(--cor-acento-teal); outline:none; }
  .doe-sec-menu { position:absolute; top:100%; left:0; margin-top:4px; z-index:5; width:460px; max-width:90vw;
    background:var(--cor-card); border:1px solid var(--cor-card-elevado); border-radius:10px; box-shadow:0 12px 32px rgba(0,0,0,.25);
    padding:10px; text-transform:none; letter-spacing:0; font-weight:400; }
  .doe-sec-menu input[type=search] { width:100%; min-width:0; margin-bottom:8px; }
  .doe-sec-acoes { display:flex; gap:8px; margin-bottom:6px; }
  .doe-sec-acoes button { font:inherit; font-size:12px; font-weight:700; cursor:pointer; background:none; border:none;
    color:var(--cor-acento-teal-texto); padding:2px 0; }
  /* lista de secretarias (26/11/2026): caixa | nome (à esquerda, quebra em
     linhas) | quantidade, em grade — isolada dos estilos de label/input do painel */
  .doe-sec-lista { max-height:340px; overflow:auto; border-top:1px solid var(--cor-card-elevado); padding-top:4px; }
  #doe-overlay .doe-sec-lista label { display:grid; grid-template-columns:18px minmax(0, 1fr) auto; align-items:center;
    column-gap:10px; width:100%; box-sizing:border-box; margin:0; padding:6px 8px; text-align:left;
    font-size:13px; font-weight:400; text-transform:none; letter-spacing:0; line-height:1.3;
    color:var(--cor-texto-primario); border-radius:6px; cursor:pointer; }
  #doe-overlay .doe-sec-lista label:hover { background:var(--cor-card-elevado); }
  #doe-overlay .doe-sec-lista input[type=checkbox] { width:16px; height:16px; min-width:0; margin:0; padding:0;
    accent-color:var(--cor-acento-teal); cursor:pointer; }
  #doe-overlay .doe-sec-lista span { text-align:left; }
  #doe-overlay .doe-sec-lista small { justify-self:end; font-size:11.5px; font-weight:700; min-width:28px; text-align:right;
    color:var(--cor-texto-terciario); white-space:nowrap; }
  #doe-overlay .doe-sec-lista input:checked ~ span { font-weight:700; }
  .doe-kpis { display:grid; grid-template-columns:repeat(6, minmax(0, 1fr)); gap:10px; padding:10px 20px; }
  .doe-kpi { background:var(--cor-card); border-radius:10px; padding:10px 14px; border-left:4px solid var(--cor-acento-teal); }
  .doe-kpi b { display:block; font-size:22px; font-weight:900; }
  .doe-kpi span { font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.3px; color:var(--cor-texto-secundario); }
  .doe-abas { display:flex; gap:4px; flex-wrap:wrap; padding:0 20px; border-bottom:1px solid var(--cor-card-elevado); }
  .doe-abas button { font:inherit; font-size:13px; font-weight:700; cursor:pointer; background:none; border:none;
    border-bottom:3px solid transparent; color:var(--cor-texto-secundario); padding:9px 12px; }
  .doe-abas button.ativa { color:var(--cor-acento-teal-texto); border-bottom-color:var(--cor-acento-teal); }
  .doe-abas small { font-weight:400; color:var(--cor-texto-terciario); }
  .doe-conteudo { flex:1; overflow:auto; padding:12px 20px 20px; }
  .doe-aviso { font-size:12px; color:var(--cor-texto-secundario); margin-bottom:8px; }
  .doe-vazio { padding:30px 4px; color:var(--cor-texto-secundario); font-size:14px; }
  .doe-tabela { width:100%; border-collapse:collapse; font-size:12.5px; background:var(--cor-card); border-radius:8px; }
  .doe-tabela, .doe-status, #doeMeta { user-select:text; }
  .doe-tabela th { position:sticky; top:0; z-index:1; text-align:left; background:var(--cor-card-elevado); color:var(--cor-texto-secundario);
    font-size:11px; text-transform:uppercase; letter-spacing:.3px; padding:8px; cursor:pointer; white-space:nowrap; }
  .doe-tabela th.ordenado { color:var(--cor-acento-teal-texto); }
  .doe-tabela td { padding:7px 8px; border-bottom:1px solid var(--cor-card-elevado); vertical-align:top; }
  .doe-tabela td.num { text-align:right; white-space:nowrap; }
  .doe-tabela td.nowrap { white-space:nowrap; }
  .doe-tabela tr.grupo td { background:var(--cor-fundo); font-weight:800; color:var(--cor-acento-teal-texto); font-size:12px; padding-top:12px; }
  .doe-tag { display:inline-block; font-size:11px; font-weight:700; padding:2px 8px; border-radius:999px; white-space:nowrap; }
  .doe-tag.nom { background:rgba(46,125,50,.15); color:#2E7D32; }
  .doe-tag.exo { background:rgba(214,69,69,.15); color:#D64545; }
  html[data-tema="escuro"] .doe-tag.nom { color:#81C784; }
  html[data-tema="escuro"] .doe-tag.exo { color:#FF8A80; }
  .doe-lista { margin:0; padding-left:18px; font-size:13px; line-height:1.6; }
  .doe-btn-cab { font-weight:800; }
  /* cartão DOE na barra de botões dos dashes: mesma classe do vizinho */
  @media (max-width:900px) { .doe-kpis { grid-template-columns:repeat(3, minmax(0, 1fr)); } }
  @media print { #doe-overlay { display:none !important; } }
</style>
<script>
(function () {
  var overlay = document.getElementById("doe-overlay");
  if (!overlay) return;
  document.body.appendChild(overlay);
  var DOE = null, carregando = null, deHome = false;
  var LS_SEC = "cgapeDoeSecretariasV1";
  var E = {dia: "", busca: "", aba: "atos", ordem: {}, sec: lerSec()};
  var LIMITE = 1500;

  function lerSec() { try { var v = JSON.parse(localStorage.getItem(LS_SEC) || "[]"); return Array.isArray(v) ? v : []; } catch (e) { return []; } }
  function gravarSec() { try { localStorage.setItem(LS_SEC, JSON.stringify(E.sec)); } catch (e) {} }
  async function api(nome) {
    var args = Array.prototype.slice.call(arguments, 1);
    if (window.pywebview && !window.pywebview.api) {   // desktop: ponte ainda subindo
      await new Promise(function (ok) { window.addEventListener("pywebviewready", ok, {once: true}); });
    }
    if (window.pywebview && window.pywebview.api && window.pywebview.api[nome]) {
      return await window.pywebview.api[nome].apply(null, args);
    }
    var r = await fetch("/api/" + nome, {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(args)});
    return await r.json();
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c];
    });
  }
  function norm(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase(); }
  function br(iso) { return iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : ""; }
  function chaveDia(d) { return d.data + (d.suplemento ? "_s" : ""); }
  function rotDia(d) { return br(d.data) + (d.suplemento ? " (suplemento)" : ""); }
  function valorNum(v) { var m = String(v || "").match(/-?R\$\s?([\d.]+,\d{2})/); if (!m) return null;
    var n = parseFloat(m[1].replace(/\./g, "").replace(",", ".")); return /^-/.test(String(v).trim()) ? -n : n; }
  function n(x) { return Number(x || 0).toLocaleString("pt-BR"); }
  function status(txt, erro) {
    var el = document.getElementById("doeStatus");
    el.hidden = !txt; el.textContent = txt || ""; el.classList.toggle("erro", !!erro);
  }

  // ---- dados
  function carregar() {
    if (DOE) return Promise.resolve(DOE);
    if (carregando) return carregando;
    carregando = api("doe_dados").then(function (r) {
      carregando = null;
      if (!r || !r.ok) { mostrarErro((r && r.mensagem) || "Não foi possível ler o DOE."); return null; }
      DOE = r.doe; montarControles(); return DOE;
    }, function (e) { carregando = null; mostrarErro("Erro ao ler o DOE: " + e); return null; });
    return carregando;
  }
  function mostrarErro(msg) {
    document.getElementById("doeConteudo").innerHTML = '<div class="doe-vazio">' + esc(msg) + "</div>";
    document.getElementById("doeKpis").innerHTML = ""; document.getElementById("doeAbas").innerHTML = "";
  }

  // ---- controles
  function montarControles() {
    var ordem = DOE.diarios.slice().sort(function (a, b) { return chaveDia(b) < chaveDia(a) ? -1 : 1; });
    var sel = document.getElementById("doeDia");
    sel.innerHTML = '<option value="">Todos os diários (' + DOE.diarios.length + ")</option>" +
      ordem.map(function (d) { return '<option value="' + chaveDia(d) + '">' + rotDia(d) + "</option>"; }).join("");
    if (E.dia && !DOE.diarios.some(function (d) { return chaveDia(d) === E.dia; })) E.dia = "";
    sel.value = E.dia;
    E.sec = E.sec.filter(function (s) { return DOE.secretarias.indexOf(s) !== -1; });
    document.getElementById("doeMeta").textContent = "Publicações de todas as secretarias · " +
      (DOE.periodo.length ? br(DOE.periodo[0]) + " a " + br(DOE.periodo[1]) + " · " : "") +
      DOE.diarios.length + " diários (com suplementos) · dados gerados em " + DOE.gerado_em;
    montarListaSec();
  }
  function contagemPorSec() {
    var c = {};
    diasFiltrados().forEach(function (d) {
      d.atos.forEach(function (a) { c[a.secretaria] = (c[a.secretaria] || 0) + 1; });
      d.licitacoes.forEach(function (a) { c[a.secretaria] = (c[a.secretaria] || 0) + 1; });
    });
    return c;
  }
  function montarListaSec() {
    var termo = norm(document.getElementById("doeSecBusca").value), c = contagemPorSec();
    document.getElementById("doeSecLista").innerHTML = DOE.secretarias.filter(function (s) { return !termo || norm(s).indexOf(termo) !== -1; })
      .map(function (s) {
        return '<label><input type="checkbox" value="' + esc(s) + '"' + (E.sec.indexOf(s) !== -1 ? " checked" : "") + "><span>" +
          esc(s) + "</span><small>" + n(c[s]) + "</small></label>";
      }).join("") || '<div class="doe-vazio">Nenhuma secretaria com esse nome.</div>';
    var b = document.getElementById("doeSecBtn");
    b.textContent = !E.sec.length ? "Todas as secretarias" : E.sec.length === 1 ? E.sec[0] : E.sec.length + " secretarias selecionadas";
    b.title = E.sec.length ? E.sec.join("\n") : "Todas as secretarias";
    b.classList.toggle("ativo", !!E.sec.length);
  }

  // ---- filtros
  function diasFiltrados() {
    return DOE.diarios.filter(function (d) { return !E.dia || chaveDia(d) === E.dia; });
  }
  function passaSec(x) { return !E.sec.length || E.sec.indexOf(x.secretaria) !== -1; }
  function passaBusca(campos) {
    var termos = norm(E.busca).split(/\s+/).filter(Boolean);
    if (!termos.length) return true;
    var t = norm(campos.join(" "));
    return termos.every(function (p) { return t.indexOf(p) !== -1; });
  }
  function linhas(aba) {
    var out = [];
    diasFiltrados().forEach(function (d) {
      var base = {_data: d.data, _dia: rotDia(d)};
      if (aba === "atos") d.atos.forEach(function (a) {
        if (passaSec(a) && passaBusca([a.secretaria, a.orgao, a.nome, a.tipo, a.cargo, a.simbolo, a.lotacao])) out.push(Object.assign({}, base, a));
      });
      if (aba === "licitacoes") d.licitacoes.forEach(function (a) {
        if (passaSec(a) && passaBusca([a.secretaria, a.tipo, a.sei, a.objeto, a.empresa, a.valor, a.data])) out.push(Object.assign({}, base, a));
      });
      if (aba === "decretos") d.decretos.forEach(function (a) {
        if (passaBusca([a.tipo, a.numero, a.data, a.ementa])) out.push(Object.assign({}, base, a));
      });
      if (aba === "casa_civil") d.casa_civil.forEach(function (a) {
        if (passaBusca([a.numero, a.texto])) out.push(Object.assign({}, base, a));
      });
      if (aba === "colegiados") (d.colegiados || []).forEach(function (a) {
        if (passaBusca([a.orgao, a.colegiado, a.membros])) out.push(Object.assign({}, base, a));
      });
    });
    return out;
  }

  var ABAS = [
    ["atos", "Nomeações e exonerações"], ["licitacoes", "Licitações e contratos"], ["decretos", "Decretos"],
    ["casa_civil", "Portarias da Casa Civil"], ["colegiados", "Conselhos e comissões"], ["sobre", "Sobre os dados"]
  ];
  var COLS = {
    atos: [["_data", "Diário"], ["secretaria", "Secretaria"], ["nome", "Nome"], ["tipo", "Ato"], ["cargo", "Cargo"],
           ["simbolo", "Símbolo"], ["lotacao", "Lotação"], ["orgao", "Órgão (no ato)"]],
    licitacoes: [["_data", "Diário"], ["secretaria", "Secretaria"], ["tipo", "Tipo"], ["objeto", "Objeto / resumo"],
                 ["empresa", "Empresa"], ["valor", "Valor"], ["sei", "Processo SEI"], ["data", "Data"]],
    decretos: [["_data", "Diário"], ["tipo", "Tipo"], ["numero", "Número"], ["data", "Data"], ["ementa", "Resumo (ementa)"]],
    casa_civil: [["_data", "Diário"], ["numero", "Número / data"], ["texto", "Resumo"]],
    colegiados: [["_data", "Diário"], ["orgao", "Órgão"], ["colegiado", "Colegiado / encargo"], ["membros", "Nomeados"]]
  };

  function render() {
    if (!DOE) return;
    var cont = {};
    ABAS.forEach(function (a) { if (a[0] !== "sobre") cont[a[0]] = linhas(a[0]); });
    var nom = cont.atos.filter(function (a) { return a.tipo === "Nomeação"; }).length;
    var soma = cont.licitacoes.reduce(function (s, a) { var v = valorNum(a.valor); return s + (v || 0); }, 0);
    var kpi = function (rot, val, tit) { return '<div class="doe-kpi"' + (tit ? ' title="' + esc(tit) + '"' : "") + "><b>" + val + "</b><span>" + rot + "</span></div>"; };
    document.getElementById("doeKpis").innerHTML =
      kpi("Diários", n(diasFiltrados().length)) + kpi("Nomeações", n(nom)) + kpi("Exonerações", n(cont.atos.length - nom)) +
      kpi("Licitações e contratos", n(cont.licitacoes.length), "Soma dos valores encontrados: R$ " + soma.toLocaleString("pt-BR", {minimumFractionDigits: 2, maximumFractionDigits: 2})) +
      kpi("Decretos", n(cont.decretos.length)) + kpi("Portarias da Casa Civil", n(cont.casa_civil.length));
    document.getElementById("doeAbas").innerHTML = ABAS.map(function (a) {
      return '<button type="button" role="tab" data-aba="' + a[0] + '"' + (E.aba === a[0] ? ' class="ativa"' : "") + ">" + a[1] +
        (cont[a[0]] ? " <small>(" + n(cont[a[0]].length) + ")</small>" : "") + "</button>";
    }).join("");
    var alvo = document.getElementById("doeConteudo");
    if (E.aba === "sobre") {
      alvo.innerHTML = '<ul class="doe-lista">' + DOE.limitacoes.map(function (l) { return "<li>" + esc(l) + "</li>"; }).join("") +
        "<li>Dias de terça a sábado sem diário na pasta: " + esc(DOE.sem_arquivo.join(", ") || "nenhum") + ".</li></ul>";
      return;
    }
    var lista = cont[E.aba], cols = COLS[E.aba].filter(function (c) { return !(E.dia && c[0] === "_data"); });
    var o = E.ordem[E.aba];
    if (o) {
      lista = lista.slice().sort(function (a, b) {
        var x = a[o.col], y = b[o.col];
        if (o.col === "valor") { x = valorNum(x); y = valorNum(y); x = x == null ? -Infinity : x; y = y == null ? -Infinity : y; return (x - y) * o.dir; }
        return norm(x).localeCompare(norm(y), "pt-BR", {numeric: true}) * o.dir;
      });
    } else if (!E.dia) {
      lista = lista.slice().sort(function (a, b) { return a._data < b._data ? 1 : a._data > b._data ? -1 : 0; });
    }
    var aviso = "";
    if ((E.aba === "decretos" || E.aba === "casa_civil" || E.aba === "colegiados") && E.sec.length) {
      aviso = '<div class="doe-aviso">O filtro de secretarias não se aplica a esta aba.</div>';
    }
    if (!lista.length) { alvo.innerHTML = aviso + '<div class="doe-vazio">Nenhuma publicação encontrada com os filtros atuais.</div>'; return; }
    var corte = lista.length > LIMITE;
    if (corte) aviso += '<div class="doe-aviso">Mostrando as ' + n(LIMITE) + " primeiras de " + n(lista.length) +
      " publicações — escolha um diário, secretarias ou digite na busca para ver as demais.</div>";
    var html = aviso + '<table class="doe-tabela"><thead><tr>' + cols.map(function (c) {
      var ord = o && o.col === c[0];
      return '<th data-col="' + c[0] + '"' + (ord ? ' class="ordenado"' : "") + ' title="Ordenar">' + c[1] + (ord ? (o.dir > 0 ? " ▲" : " ▼") : "") + "</th>";
    }).join("") + "</tr></thead><tbody>";
    html += lista.slice(0, LIMITE).map(function (a) {
      return "<tr>" + cols.map(function (c) {
        var v = a[c[0]];
        if (c[0] === "_data") return '<td class="nowrap">' + esc(a._dia) + "</td>";
        if (c[0] === "tipo" && E.aba === "atos") return '<td><span class="doe-tag ' + (v === "Nomeação" ? "nom" : "exo") + '">' + esc(v) + "</span></td>";
        if (c[0] === "valor") return '<td class="num">' + esc(v) + "</td>";
        if (c[0] === "data" || c[0] === "sei" || c[0] === "simbolo" || c[0] === "numero") return '<td class="nowrap">' + esc(v) + "</td>";
        return "<td>" + esc(v) + "</td>";
      }).join("") + "</tr>";
    }).join("") + "</tbody></table>";
    alvo.innerHTML = html;
  }

  // ---- eventos
  document.getElementById("doeDia").onchange = function () { E.dia = this.value; montarListaSec(); render(); };
  var tBusca = null;
  document.getElementById("doeBusca").oninput = function () {
    var v = this.value; clearTimeout(tBusca); tBusca = setTimeout(function () { E.busca = v; render(); }, 180);
  };
  document.getElementById("doeAbas").onclick = function (ev) {
    var b = ev.target.closest("button[data-aba]"); if (!b) return; E.aba = b.dataset.aba; render();
  };
  document.getElementById("doeConteudo").onclick = function (ev) {
    var th = ev.target.closest("th[data-col]"); if (!th) return;
    var o = E.ordem[E.aba], col = th.dataset.col;
    E.ordem[E.aba] = o && o.col === col ? (o.dir > 0 ? {col: col, dir: -1} : null) : {col: col, dir: col === "valor" ? -1 : 1};
    render();
  };
  var menu = document.getElementById("doeSecMenu"), secBtn = document.getElementById("doeSecBtn");
  function menuAberto(sim) { menu.hidden = !sim; secBtn.setAttribute("aria-expanded", sim ? "true" : "false");
    if (sim) { document.getElementById("doeSecBusca").value = ""; montarListaSec(); document.getElementById("doeSecBusca").focus(); } }
  secBtn.onclick = function () { menuAberto(menu.hidden); };
  document.addEventListener("mousedown", function (ev) {
    if (!menu.hidden && !menu.contains(ev.target) && ev.target !== secBtn) menuAberto(false);
  });
  document.getElementById("doeSecBusca").oninput = montarListaSec;
  document.getElementById("doeSecLista").onchange = function (ev) {
    var cb = ev.target; if (!cb || cb.type !== "checkbox") return;
    var i = E.sec.indexOf(cb.value);
    if (cb.checked && i === -1) E.sec.push(cb.value);
    if (!cb.checked && i !== -1) E.sec.splice(i, 1);
    gravarSec(); montarListaSec(); render();
  };
  document.getElementById("doeSecTodas").onclick = function () {
    var termo = norm(document.getElementById("doeSecBusca").value);
    DOE.secretarias.forEach(function (s) { if ((!termo || norm(s).indexOf(termo) !== -1) && E.sec.indexOf(s) === -1) E.sec.push(s); });
    if (E.sec.length === DOE.secretarias.length) E.sec = [];   // todas = sem filtro
    gravarSec(); montarListaSec(); render();
  };
  document.getElementById("doeSecLimpar").onclick = function () { E.sec = []; gravarSec(); montarListaSec(); render(); };

  document.getElementById("doeAtualizar").onclick = function () {
    var b = this; b.disabled = true; b.textContent = "ATUALIZANDO…";
    status("Baixando do site oficial os diários dos últimos dias que faltam e refazendo a análise… pode levar alguns minutos.");
    api("doe_atualizar").then(function (r) {
      if (r && r.doe) { DOE = r.doe; montarControles(); render(); }
      status((r && r.mensagem) || "Sem resposta.", !(r && r.ok));
    }, function (e) { status("Erro ao atualizar o DOE: " + e, true); })
      .then(function () { b.disabled = false; b.textContent = "ATUALIZAR DOE"; });
  };
  function fechar() {
    overlay.classList.remove("aberto"); overlay.setAttribute("aria-hidden", "true"); menuAberto(false);
    if (deHome && window.irParaHome) window.irParaHome();
  }
  document.getElementById("doeFechar").onclick = fechar;
  document.addEventListener("keydown", function (ev) {
    if (ev.key !== "Escape" || !overlay.classList.contains("aberto")) return;
    if (!menu.hidden) { menuAberto(false); return; }
    fechar();
  });

  window.doeAbrir = function (opcoes) {
    deHome = !!(opcoes && opcoes.deHome);
    overlay.classList.add("aberto"); overlay.setAttribute("aria-hidden", "false");
    carregar().then(function (d) { if (d) render(); });
  };

  // Botão DOE no cabeçalho dos dashes (PAC, MONITORAMENTO e PRAZOS), logo
  // depois da busca geral, com a mesma classe do botão vizinho
  function botaoNoCabecalho(busca) {
    if (!busca || !busca.parentNode || busca.parentNode.querySelector(".doe-btn-cab")) return;
    var viz = busca.nextElementSibling;
    var b = document.createElement("button");
    b.type = "button"; b.className = ((viz && viz.tagName === "BUTTON" && viz.className) || "btn") + " doe-btn-cab";
    b.textContent = "DOE"; b.title = "Diário Oficial do Estado — publicações de todas as secretarias";
    b.onclick = function () { window.doeAbrir(); };
    busca.parentNode.insertBefore(b, viz);
  }
  [document.getElementById("dash-busca-geral")].concat(Array.prototype.slice.call(document.querySelectorAll(".ra-botoes .ra-busca-geral")))
    .forEach(botaoNoCabecalho);
})();
</script>
"""
