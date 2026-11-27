"""GERAR RELATÓRIO do CONTROLE DE PRAZOS — pedido do usuário em 26/11/2026.

Gera HTML (página única, autossuficiente), PDF (reportlab, A4 paisagem) e
PPTX (python-pptx, 16:9) com as ações do RECORTE ATUAL do dash (filtros da
grade + busca — a tela manda os ids), com:
  - indicadores: total, um por alerta e investimento total;
  - resumo por órgão/ministério (quantidade por alerta + investimento);
  - as ações agrupadas por alerta (Vencido → Concluído), ordenadas pelo prazo.

A regra de alerta é a mesma do dash e do Excel (controle_prazos.py):
Concluído pelo status; sem data → "Sem data"; dias < 0 Vencido; ≤ 7 Crítico;
≤ 30 Atenção; senão OK.

Entrada: montar(itens, formato, recorte, atualizado_em) -> (bytes, extensão, mimetype).
"""
from __future__ import annotations

import datetime as dt
import html
import io
import unicodedata

ALERTAS = ["Vencido", "Crítico", "Atenção", "OK", "Sem data", "Concluído"]
ALERTA_COR = {"Vencido": "C62828", "Crítico": "E65100", "Atenção": "C99400",
              "OK": "2E7D32", "Sem data": "8C97A6", "Concluído": "0B5CAD"}
ALERTA_DESC = {"Vencido": "prazo já passou", "Crítico": "vence em até 7 dias",
               "Atenção": "vence em 8 a 30 dias", "OK": "vence em mais de 30 dias",
               "Sem data": "prazo a definir", "Concluído": "ações concluídas"}
VERMELHO = "C62828"
NAVY = "0B1F3F"
CINZA = "5B6576"
TITULO = "CONTROLE DE PRAZOS"
SUBTITULO = "Pauta do Governador · Portos e Transporte"

MIMES = {
    "html": "text/html; charset=utf-8",
    "pdf": "application/pdf",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}


# ---------------------------------------------------------------- dados ----
def _norm(t):
    t = unicodedata.normalize("NFD", str(t or "")).encode("ascii", "ignore").decode()
    return t.lower().strip()


def _data(iso):
    try:
        return dt.date.fromisoformat(str(iso)[:10]) if iso else None
    except ValueError:
        return None


def moeda(v, curta=False):
    if v in (None, ""):
        return ""
    n = float(v)
    if curta:
        for lim, suf in ((1e9, "BI"), (1e6, "MI"), (1e3, "MIL")):
            if n >= lim:
                return f"R$ {n / lim:.1f} {suf}".replace(".", ",")
    s = f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "R$ " + s


def preparar(itens, hoje=None):
    """Itens do cadastro -> linhas com alerta, dias e textos prontos."""
    hoje = hoje or dt.date.today()
    linhas = []
    for r in itens:
        d = _data(r.get("data_prazo"))
        concl = _norm(r.get("status")).startswith("conclu")
        dias = (d - hoje).days if d else None
        alerta = ("Concluído" if concl else "Sem data" if d is None else "Vencido" if dias < 0
                  else "Crítico" if dias <= 7 else "Atenção" if dias <= 30 else "OK")
        if alerta == "Concluído":
            situacao = "concluído"
        elif alerta == "Sem data":
            situacao = "sem prazo definido"
        elif dias < 0:
            situacao = f"{-dias} {'dia' if dias == -1 else 'dias'} de atraso"
        elif dias == 0:
            situacao = "vence hoje"
        else:
            situacao = f"faltam {dias} {'dia' if dias == 1 else 'dias'}"
        inf = str(r.get("prazo_informado") or "").strip()
        if _norm(inf) == "a definir":
            inf = ""
        inv = r.get("investimento")
        try:
            inv = float(inv) if inv not in (None, "") else None
        except (TypeError, ValueError):
            inv = None
        linhas.append({
            "id": str(r.get("id") or ""), "descricao": str(r.get("descricao") or "").strip() or "(sem descrição)",
            "projeto": str(r.get("projeto") or "").strip(), "ministerio": str(r.get("ministerio") or "").strip(),
            "eixo": str(r.get("eixo") or "").strip(), "status": str(r.get("status") or "").strip(),
            "responsavel": str(r.get("responsavel") or "").strip() or "Não informado",
            "observacoes": str(r.get("observacoes") or "").strip(),
            "data": d, "data_txt": d.strftime("%d/%m/%Y") if d else "a definir",
            "prazo_informado": inf, "dias": dias, "alerta": alerta, "situacao": situacao,
            "investimento": inv,
        })
    linhas.sort(key=lambda x: (ALERTAS.index(x["alerta"]), x["data"] or dt.date.max, _norm(x["descricao"])))
    return linhas


def resumo(linhas):
    cont = {a: 0 for a in ALERTAS}
    for x in linhas:
        cont[x["alerta"]] += 1
    invs = [x["investimento"] for x in linhas if x["investimento"] is not None]
    por_orgao = {}
    for x in linhas:
        o = por_orgao.setdefault(x["ministerio"] or "Não informado",
                                 {"total": 0, "inv": 0.0, "n_inv": 0, **{a: 0 for a in ALERTAS}})
        o["total"] += 1
        o[x["alerta"]] += 1
        if x["investimento"] is not None:
            o["inv"] += x["investimento"]
            o["n_inv"] += 1
    orgaos = sorted(por_orgao.items(), key=lambda kv: (-kv[1]["total"], _norm(kv[0])))
    return {"total": len(linhas), "cont": cont, "inv_total": sum(invs), "n_inv": len(invs), "orgaos": orgaos}


def grupos(linhas):
    return [(a, [x for x in linhas if x["alerta"] == a]) for a in ALERTAS
            if any(x["alerta"] == a for x in linhas)]


def nome_arquivo(formato, agora=None):
    agora = agora or dt.datetime.now()
    return f"Relatório Controle de Prazos - {agora:%d-%m-%Y %Hh%M}.{formato}"


def montar(itens, formato, recorte="", atualizado_em=""):
    formato = str(formato or "").lower().strip(".")
    if formato not in MIMES:
        raise ValueError("formato inválido (use html, pdf ou pptx)")
    linhas = preparar(itens)
    info = {"recorte": str(recorte or "").strip() or "todas as ações cadastradas",
            "atualizado": str(atualizado_em or ""), "gerado": dt.datetime.now().strftime("%d/%m/%Y às %Hh%M")}
    fn = {"html": _html, "pdf": _pdf, "pptx": _pptx}[formato]
    return fn(linhas, resumo(linhas), info), formato, MIMES[formato]


# ----------------------------------------------------------------- HTML ----
# Página única e autossuficiente, pensada para três jeitos de abrir (pedido de
# 27/11/2026): tela 16:9 (ações em duas colunas), celular em retrato (tudo
# empilhado, resumo por órgão em cartões, barra de busca/filtro fixa no topo) e
# celular em paisagem (cabeçalho compacto, indicadores numa linha só, ações em
# duas colunas). Cada ação tem o seu botão de compartilhar no WhatsApp, com o
# texto da ação já montado (wa.me abre a escolha do contato).
EMOJI_ALERTA = {"Vencido": "\U0001F534", "Crítico": "\U0001F7E0", "Atenção": "\U0001F7E1",
                "OK": "\U0001F7E2", "Sem data": "⚪", "Concluído": "\U0001F535"}

_ICONE_WHATS = ('<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12.04 2C6.58 2 2.13 6.45 '
                '2.13 11.91c0 1.75.46 3.45 1.32 4.95L2.05 22l5.25-1.38c1.45.79 3.08 1.21 4.74 1.21 5.46 0 9.91-4.45 '
                '9.91-9.91C21.95 6.45 17.5 2 12.04 2Zm5.8 14.02c-.24.68-1.42 1.31-1.96 1.36-.5.05-.97.24-3.27-.68-2.77'
                '-1.09-4.52-3.92-4.66-4.1-.13-.18-1.11-1.48-1.11-2.82 0-1.34.7-2 .95-2.27.25-.27.54-.34.72-.34h.52c.17 '
                '0 .4-.06.62.47.24.56.8 1.94.87 2.08.07.14.12.3.02.48-.09.18-.14.3-.27.46-.14.16-.29.36-.41.48-.14.14'
                '-.28.29-.12.56.16.27.71 1.17 1.52 1.9 1.05.93 1.93 1.22 2.2 1.36.27.14.43.12.59-.07.16-.18.68-.8.86'
                '-1.07.18-.27.36-.23.61-.14.25.09 1.58.75 1.85.88.27.14.45.2.52.32.07.11.07.66-.17 1.34Z"/></svg>')


def _texto_whats(x, info):
    a = x["alerta"]
    prazo = x["data_txt"] + (f' ({x["prazo_informado"]})' if x["prazo_informado"] else "")
    partes = [f"*CONTROLE DE PRAZOS · CGAPE*",
              f'{EMOJI_ALERTA[a]} *{a.upper()}* — {x["situacao"]}',
              "",
              f'*Ação:* {x["descricao"]}']
    for rot, chave in (("Órgão/Ministério", "ministerio"), ("Eixo", "eixo"), ("Projeto", "projeto")):
        if x[chave]:
            partes.append(f"*{rot}:* {x[chave]}")
    partes += [f"*Prazo:* {prazo}", f'*Status:* {x["status"] or "—"}', f'*Responsável:* {x["responsavel"]}']
    if x["investimento"] is not None:
        partes.append(f'*Investimento:* {moeda(x["investimento"])}')
    if x["observacoes"]:
        partes.append(f'*Observações:* {x["observacoes"]}')
    partes += ["", f'_Relatório gerado em {info["gerado"]} · ação #{x["id"]}_']
    return "\n".join(partes)


def _link_whats(x, info):
    from urllib.parse import quote
    return "https://wa.me/?text=" + quote(_texto_whats(x, info), safe="")


def _slug(a):
    return _norm(a).replace(" ", "-")


_CSS_HTML = r"""
:root { --fundo:#F2F4F7; --card:#fff; --borda:#D8DCE3; --txt:#0B1F3F; --txt2:#5B6576; --txt3:#8A93A3;
  --acento:#C62828; --sombra:0 1px 3px rgba(11,31,63,.08); --sombra2:0 6px 18px rgba(11,31,63,.14); --barra:rgba(242,244,247,.92); }
@media (prefers-color-scheme: dark) { :root:not([data-tema="claro"]) { --fundo:#0E1726; --card:#16233A; --borda:#26344D;
  --txt:#EAF0F8; --txt2:#AEB9CA; --txt3:#7D8AA0; --acento:#EF5350; --sombra:0 1px 3px rgba(0,0,0,.3);
  --sombra2:0 6px 18px rgba(0,0,0,.45); --barra:rgba(14,23,38,.92); } }
:root[data-tema="escuro"] { --fundo:#0E1726; --card:#16233A; --borda:#26344D; --txt:#EAF0F8; --txt2:#AEB9CA;
  --txt3:#7D8AA0; --acento:#EF5350; --sombra:0 1px 3px rgba(0,0,0,.3); --sombra2:0 6px 18px rgba(0,0,0,.45); --barra:rgba(14,23,38,.92); }
* { box-sizing:border-box; }
html { scroll-behavior:smooth; -webkit-text-size-adjust:100%; }
body { margin:0; background:var(--fundo); color:var(--txt); font-family:'Roboto','Segoe UI',Arial,sans-serif; font-size:14px;
  overflow-x:hidden; }

/* cabeçalho */
header { background:#0B1F3F; color:#fff; border-bottom:5px solid #C62828;
  padding:20px max(28px, env(safe-area-inset-right)) 16px max(28px, env(safe-area-inset-left)); }
.cab { max-width:1680px; margin:0 auto; display:flex; gap:16px; align-items:flex-end; justify-content:space-between; }
header .gov { font-size:11px; letter-spacing:1.5px; color:#4FA3E0; font-weight:700; }
header h1 { margin:4px 0 2px; font-size:26px; letter-spacing:.5px; line-height:1.15; }
header .sub { color:#C9D6EA; font-size:13px; }
#tema { flex:none; font:inherit; font-size:12px; font-weight:700; color:#fff; background:rgba(255,255,255,.1);
  border:1px solid rgba(255,255,255,.3); border-radius:999px; padding:7px 14px; cursor:pointer; min-height:36px; }

main { max-width:1680px; margin:0 auto; padding:20px max(20px, env(safe-area-inset-right)) 56px max(20px, env(safe-area-inset-left)); }

/* indicadores */
.kpis { display:grid; grid-template-columns:repeat(10, minmax(0, 1fr)); gap:12px; margin-bottom:24px; }
.kpi { background:var(--card); border-radius:12px; padding:12px 14px; border-left:5px solid var(--c, var(--acento));
  box-shadow:var(--sombra); text-decoration:none; color:inherit; display:block; min-width:0; }
a.kpi { transition:box-shadow .15s, transform .15s; } a.kpi:hover { box-shadow:var(--sombra2); transform:translateY(-1px); }
.kpi.zero { opacity:.45; }
.kpi span { display:block; font-size:11px; font-weight:700; letter-spacing:.4px; color:var(--txt2); white-space:nowrap;
  overflow:hidden; text-overflow:ellipsis; }
.kpi b { display:block; font-size:32px; font-weight:900; color:var(--c, var(--txt)); line-height:1.2; font-variant-numeric:tabular-nums;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.kpi small { display:block; font-size:11.5px; color:var(--txt3); line-height:1.3; }
.kpi.total { grid-column:span 2; --c:var(--txt); } .kpi.total b { font-size:40px; }
.kpi.inv { grid-column:span 2; --c:var(--acento); } .kpi.inv b.ni { font-size:20px; }

/* resumo por órgão */
h2.tit { font-size:15px; letter-spacing:.5px; margin:0 0 10px; color:var(--txt2); text-transform:uppercase; }
.tabela { background:var(--card); border-radius:12px; overflow:auto; box-shadow:var(--sombra); }
table { border-collapse:collapse; width:100%; font-size:13px; }
th { background:#0B1F3F; color:#fff; font-size:11.5px; font-weight:700; padding:9px 10px; text-align:center;
  border-bottom:3px solid var(--c, #C62828); white-space:nowrap; }
th:first-child { text-align:left; }
td { padding:8px 10px; border-bottom:1px solid var(--borda); }
tbody tr:last-child td { border-bottom:0; }
td.org { font-weight:600; } td.n { text-align:center; font-weight:800; color:var(--c, var(--txt)); font-variant-numeric:tabular-nums; }
td.z { color:var(--txt3); font-weight:400; } td.tot { color:var(--txt); } td.r { text-align:right; white-space:nowrap; font-weight:700; }
.orgs-m { display:none; }
.org-c { background:var(--card); border-radius:12px; box-shadow:var(--sombra); padding:12px 14px; }
.org-c .t { display:flex; justify-content:space-between; gap:10px; align-items:baseline; font-weight:700; }
.org-c .t b { font-size:20px; font-weight:900; font-variant-numeric:tabular-nums; }
.org-c .chips { display:flex; flex-wrap:wrap; gap:6px; margin-top:8px; }
.org-c .inv { margin-top:6px; font-size:12.5px; color:var(--txt2); } .org-c .inv b { color:var(--acento); }

/* selo de alerta (tom claro, legível nos dois temas) */
.chip { display:inline-flex; align-items:center; gap:6px; padding:3px 10px 3px 8px; border-radius:999px; font-size:12px;
  font-weight:800; white-space:nowrap; color:color-mix(in srgb, var(--c) 72%, #000);
  background:color-mix(in srgb, var(--c) 13%, #fff); border:1px solid color-mix(in srgb, var(--c) 45%, transparent); }
.chip::before { content:""; width:7px; height:7px; border-radius:50%; background:var(--c); flex:none; }
@media (prefers-color-scheme: dark) { :root:not([data-tema="claro"]) .chip { color:var(--c);
  background:color-mix(in srgb, var(--c) 18%, transparent); border-color:color-mix(in srgb, var(--c) 50%, transparent); } }
:root[data-tema="escuro"] .chip { color:var(--c); background:color-mix(in srgb, var(--c) 18%, transparent);
  border-color:color-mix(in srgb, var(--c) 50%, transparent); }

/* barra de busca e filtro (fica presa no topo ao rolar) */
.ferramentas { position:sticky; top:0; z-index:5; margin:26px -6px 6px; padding:10px 6px; background:var(--barra);
  -webkit-backdrop-filter:blur(8px); backdrop-filter:blur(8px); display:flex; gap:10px; align-items:center; flex-wrap:wrap; }
.ferramentas h2 { margin:0 auto 0 0; font-size:15px; letter-spacing:.5px; color:var(--txt2); }
#busca { font:inherit; font-size:16px; padding:9px 14px; border-radius:10px; border:1px solid var(--borda); background:var(--card);
  color:var(--txt); width:340px; max-width:100%; min-height:42px; }
#busca:focus { outline:none; border-color:var(--acento); }
.filtros { display:flex; gap:6px; flex-wrap:wrap; }
.filtros button { font:inherit; cursor:pointer; min-height:36px; }
.filtros button.chip[aria-pressed="false"] { opacity:.4; filter:grayscale(.6); }
#contador { font-size:12px; color:var(--txt2); width:100%; text-align:right; }
#contador:empty { display:none; }

/* ações */
.grupo { margin-top:18px; scroll-margin-top:120px; }
.grupo h2 { font-size:16px; margin:0 0 10px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.grupo h2 .chip { font-size:13px; }
.grupo h2 small { font-weight:400; color:var(--txt2); font-size:13px; }
.lista { display:grid; grid-template-columns:repeat(auto-fill, minmax(min(100%, 560px), 1fr)); gap:10px; align-items:start; }
.acao { container-type:inline-size; background:var(--card); border-radius:12px; border-left:5px solid var(--c);
  box-shadow:var(--sombra); break-inside:avoid; min-width:0; }
.acao .in { display:flex; gap:16px; padding:12px 16px; }
.acao .esq { flex:1; min-width:0; }
.acao .dir { flex:0 0 200px; text-align:right; display:flex; flex-direction:column; align-items:flex-end; gap:2px; }
.trilha { font-size:11.5px; color:var(--txt2); overflow-wrap:anywhere; }
.acao h3 { font-size:15px; margin:3px 0 7px; line-height:1.35; overflow-wrap:anywhere; }
.meta { display:flex; gap:6px 14px; flex-wrap:wrap; font-size:12.5px; color:var(--txt2); }
.meta b { color:var(--txt); font-weight:600; } .meta .id { color:var(--txt3); }
.obs { margin-top:7px; font-size:12.5px; color:var(--txt2); white-space:pre-wrap; overflow-wrap:anywhere;
  border-top:1px dashed var(--borda); padding-top:6px; }
.sit { font-weight:900; font-size:17px; color:var(--c); font-variant-numeric:tabular-nums; line-height:1.2; }
.prazo { font-size:12.5px; color:var(--txt2); } .prazo b { color:var(--txt); }
.inv-v { margin-top:4px; font-weight:800; color:var(--acento); font-variant-numeric:tabular-nums; }
.whats { margin-top:auto; padding-top:8px; }
.whats a { display:inline-flex; align-items:center; gap:7px; text-decoration:none; font-weight:700; font-size:12.5px;
  color:#fff; background:#1FA855; border-radius:999px; padding:7px 14px 7px 10px; min-height:36px;
  box-shadow:0 1px 2px rgba(0,0,0,.15); transition:background .15s, transform .15s; }
.whats a:hover { background:#168A45; transform:translateY(-1px); }
.whats a:focus-visible { outline:3px solid color-mix(in srgb, #1FA855 45%, transparent); outline-offset:2px; }
.whats svg { width:18px; height:18px; flex:none; }
@container (max-width:520px) {
  .acao .in { flex-direction:column; gap:8px; }
  .acao .dir { flex-basis:auto; text-align:left; align-items:stretch; flex-direction:row; flex-wrap:wrap;
    column-gap:12px; row-gap:2px; border-top:1px solid var(--borda); padding-top:8px; }
  .acao .dir .sit { width:100%; }
  .acao .whats { width:100%; }
  .acao .whats a { width:100%; justify-content:center; min-height:42px; font-size:14px; }
}
.vazio { color:var(--txt2); padding:30px; text-align:center; }
#nada { display:none; }
footer { text-align:center; color:var(--txt3); font-size:11.5px; padding:20px 20px max(20px, env(safe-area-inset-bottom)); }
#topo { position:fixed; right:max(16px, env(safe-area-inset-right)); bottom:max(16px, env(safe-area-inset-bottom)); z-index:6;
  width:44px; height:44px; border-radius:50%; border:0; background:#0B1F3F; color:#fff; font-size:20px; cursor:pointer;
  box-shadow:var(--sombra2); opacity:0; pointer-events:none; transition:opacity .2s; }
#topo.vis { opacity:1; pointer-events:auto; }

/* tela média (notebook / tablet paisagem) */
@media (max-width:1200px) {
  .kpis { grid-template-columns:repeat(4, minmax(0, 1fr)); }
}
/* celular / tablet em retrato */
@media (max-width:760px) {
  header { padding-top:16px; padding-bottom:12px; }
  header h1 { font-size:21px; }
  header .sub { font-size:12px; }
  .cab { align-items:flex-start; }
  main { padding-top:14px; }
  .kpis { grid-template-columns:repeat(3, minmax(0, 1fr)); gap:8px; margin-bottom:18px; }
  .kpi { padding:9px 10px; border-left-width:4px; }
  .kpi b { font-size:24px; }
  .kpi small { display:none; }
  .kpi span { font-size:10px; }
  .kpi.total, .kpi.inv { grid-column:1 / -1; display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
  .kpi.total span, .kpi.inv span { flex:1 1 auto; }
  .kpi.total b, .kpi.inv b { font-size:26px; order:2; }
  .kpi.total small, .kpi.inv small { display:block; width:100%; order:3; }
  .tabela { display:none; }
  .orgs-m { display:grid; gap:8px; }
  .ferramentas { margin-top:18px; }
  .ferramentas h2 { width:100%; font-size:13px; }
  #busca { width:100%; }
  .filtros { flex-wrap:nowrap; overflow-x:auto; width:100%; padding-bottom:2px; scrollbar-width:none; }
  .filtros::-webkit-scrollbar { display:none; }
  .lista { grid-template-columns:1fr; }
}
/* celular em paisagem (tela larga e baixa) */
@media (orientation:landscape) and (max-height:540px) {
  header { padding-top:10px; padding-bottom:8px; border-bottom-width:4px; }
  header .gov { display:none; }
  header h1 { font-size:18px; margin:0; }
  header .sub { font-size:11.5px; }
  main { padding-top:12px; }
  .kpis { grid-template-columns:minmax(0, .9fr) repeat(6, minmax(0, .75fr)) minmax(0, 1.6fr); gap:6px; margin-bottom:16px; }
  .kpi { padding:7px 8px; border-left-width:4px; display:block; }
  .kpi.total, .kpi.inv { grid-column:auto; display:block; }
  .kpi b, .kpi.total b, .kpi.inv b { font-size:20px; }
  .kpi.inv b { font-size:18px; } .kpi.inv b.ni { font-size:13px; }
  .kpi span { font-size:9.5px; }
  .kpi small, .kpi.total small, .kpi.inv small { display:none; }
  .tabela { display:block; } .orgs-m { display:none; }
  .ferramentas { position:static; margin-top:16px; }
  .ferramentas h2 { width:auto; }
  #busca { width:260px; }
  .filtros { width:auto; flex-wrap:wrap; }
  .lista { grid-template-columns:repeat(auto-fill, minmax(min(100%, 360px), 1fr)); }
}
@media print {
  body { background:#fff; } #tema, #topo, .filtros, #busca, .whats { display:none !important; }
  .ferramentas { position:static; background:none; }
  .acao, .kpi, .tabela, .org-c { box-shadow:none; border:1px solid #D8DCE3; }
  .lista { display:block; } .acao { margin-bottom:8px; }
}
"""

_JS_HTML = r"""
(function () {
  var raiz = document.documentElement, botTema = document.getElementById("tema");
  function temaAtual() {
    return raiz.dataset.tema || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "escuro" : "claro");
  }
  function rotulo() { botTema.textContent = temaAtual() === "escuro" ? "☀ Modo claro" : "☾ Modo escuro"; }
  try { var salvo = localStorage.getItem("prz-rel-tema"); if (salvo) raiz.dataset.tema = salvo; } catch (e) {}
  rotulo();
  botTema.addEventListener("click", function () {
    raiz.dataset.tema = temaAtual() === "escuro" ? "claro" : "escuro"; rotulo();
    try { localStorage.setItem("prz-rel-tema", raiz.dataset.tema); } catch (e) {}
  });

  var caixa = document.getElementById("busca"), contador = document.getElementById("contador");
  var botoes = [].slice.call(document.querySelectorAll(".filtros button"));
  var total = document.querySelectorAll(".acao").length;
  function norm(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase(); }
  function aplicar() {
    var t = norm(caixa.value).split(/\s+/).filter(Boolean), vis = 0;
    var ligados = botoes.filter(function (b) { return b.getAttribute("aria-pressed") === "true"; })
                        .map(function (b) { return b.dataset.alerta; });
    document.querySelectorAll(".grupo").forEach(function (g) {
      var n = 0, liberado = ligados.indexOf(g.dataset.alerta) !== -1;
      g.querySelectorAll(".acao").forEach(function (a) {
        var ok = liberado && t.every(function (p) { return a.dataset.busca.indexOf(p) !== -1; });
        a.style.display = ok ? "" : "none"; if (ok) n++;
      });
      g.style.display = n ? "" : "none"; vis += n;
    });
    contador.textContent = vis === total ? "" : vis + " de " + total + " ações";
    document.getElementById("nada").style.display = vis || !total ? "none" : "block";
  }
  caixa.addEventListener("input", aplicar);
  botoes.forEach(function (b) {
    b.addEventListener("click", function () {
      b.setAttribute("aria-pressed", b.getAttribute("aria-pressed") === "true" ? "false" : "true"); aplicar();
    });
  });
  // clicar num indicador garante que o grupo esteja visível antes de rolar até ele
  document.querySelectorAll("a.kpi").forEach(function (k) {
    k.addEventListener("click", function () {
      var alvo = k.getAttribute("href").slice(1), g = document.getElementById(alvo);
      if (!g) return;
      botoes.forEach(function (b) { if (b.dataset.alerta === g.dataset.alerta) b.setAttribute("aria-pressed", "true"); });
      aplicar();
    });
  });
  var topo = document.getElementById("topo");
  window.addEventListener("scroll", function () { topo.classList.toggle("vis", window.scrollY > 600); }, {passive: true});
  topo.addEventListener("click", function () { window.scrollTo(0, 0); });
})();
"""


def _html(linhas, res, info):
    e = html.escape
    kpis = [f'<div class="kpi total"><span>AÇÕES NO RECORTE</span><b>{res["total"]}</b>'
            f'<small>{e(info["recorte"])}</small></div>']
    for a in ALERTAS:
        n = res["cont"][a]
        kpis.append(f'<a class="kpi{" zero" if not n else ""}" href="#g-{_slug(a)}" '
                    f'style="--c:#{ALERTA_COR[a]}"><span>{e(a.upper())}</span><b>{n}</b>'
                    f'<small>{e(ALERTA_DESC[a])}</small></a>')
    inv_txt = moeda(res["inv_total"], True) if res["n_inv"] else "não informado"
    kpis.append(f'<div class="kpi inv"><span>INVESTIMENTO</span><b{"" if res["n_inv"] else " class=ni"}>{e(inv_txt)}</b>'
                f'<small>{res["n_inv"]} de {res["total"]} ações com valor informado</small></div>')

    cab_org = "".join(f'<th style="--c:#{ALERTA_COR[a]}">{e(a)}</th>' for a in ALERTAS)
    lin_org, cards_org = [], []
    for nome, o in res["orgaos"]:
        cels = "".join(f'<td class="n{" z" if not o[a] else ""}" style="--c:#{ALERTA_COR[a]}">{o[a] or "–"}</td>'
                       for a in ALERTAS)
        inv_o = moeda(o["inv"], True) if o["n_inv"] else ""
        lin_org.append(f'<tr><td class="org">{e(nome)}</td>{cels}<td class="n tot">{o["total"]}</td>'
                       f'<td class="r">{e(inv_o) or "–"}</td></tr>')
        chips = "".join(f'<span class="chip" style="--c:#{ALERTA_COR[a]}">{o[a]} {e(a.lower())}</span>'
                        for a in ALERTAS if o[a])
        cards_org.append(f'<div class="org-c"><div class="t"><span>{e(nome)}</span><b>{o["total"]}</b></div>'
                         f'<div class="chips">{chips}</div>'
                         + (f'<div class="inv">Investimento: <b>{e(inv_o)}</b></div>' if inv_o else "") + "</div>")

    secoes = []
    for a, lst in grupos(linhas):
        cor = ALERTA_COR[a]
        cards = []
        for x in lst:
            obs = (f'<div class="obs"><b>Observações:</b> {e(x["observacoes"])}</div>' if x["observacoes"] else "")
            inf = f' <i>({e(x["prazo_informado"])})</i>' if x["prazo_informado"] else ""
            inv = (f'<div class="inv-v">{e(moeda(x["investimento"]))}</div>' if x["investimento"] is not None else "")
            busca = _norm(" ".join([x["id"], x["descricao"], x["projeto"], x["ministerio"], x["eixo"], x["status"],
                                    x["responsavel"], x["observacoes"], a]))
            trilha = " · ".join(v for v in (x["ministerio"], x["eixo"], x["projeto"]) if v)
            cards.append(f'''<article class="acao" data-busca="{e(busca)}"><div class="in">
  <div class="esq">
    <div class="trilha">{e(trilha)}</div>
    <h3>{e(x["descricao"])}</h3>
    <div class="meta"><span><b>Status:</b> {e(x["status"] or "—")}</span><span><b>Responsável:</b> {e(x["responsavel"])}</span><span class="id">#{e(x["id"])}</span></div>
    {obs}
  </div>
  <div class="dir">
    <div class="sit">{e(x["situacao"])}</div>
    <div class="prazo">Prazo: <b>{e(x["data_txt"])}</b>{inf}</div>
    {inv}
    <div class="whats"><a href="{e(_link_whats(x, info))}" target="_blank" rel="noopener" title="Compartilhar esta ação no WhatsApp">{_ICONE_WHATS}Compartilhar</a></div>
  </div>
</div></article>''')
        secoes.append(f'''<section class="grupo" id="g-{_slug(a)}" data-alerta="{e(a)}" style="--c:#{cor}">
  <h2><span class="chip">{e(a)}</span> {len(lst)} {"ação" if len(lst) == 1 else "ações"} <small>· {e(ALERTA_DESC[a])}</small></h2>
  <div class="lista">{"".join(cards)}</div>
</section>''')

    filtros = "".join(f'<button type="button" class="chip" style="--c:#{ALERTA_COR[a]}" data-alerta="{e(a)}" '
                      f'aria-pressed="true">{e(a)} · {len(lst)}</button>' for a, lst in grupos(linhas))
    vazio = '<p class="vazio">Nenhuma ação no recorte.</p>' if not linhas else ""
    atual = f' · cadastro atualizado em {e(info["atualizado"])}' if info["atualizado"] else ""
    return f'''<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#0B1F3F">
<title>Relatório · Controle de Prazos</title>
<style>{_CSS_HTML}</style></head>
<body>
<header><div class="cab">
  <div>
    <div class="gov">GOVERNO DO ESTADO DA BAHIA · CGAPE</div>
    <h1>{TITULO}</h1>
    <div class="sub">{e(SUBTITULO)} · relatório gerado em {e(info["gerado"])}{atual}</div>
  </div>
  <button type="button" id="tema">Tema</button>
</div></header>
<main>
  <div class="kpis">{"".join(kpis)}</div>
  <h2 class="tit">Resumo por órgão / ministério</h2>
  <div class="tabela"><table><thead><tr><th>Órgão / Ministério</th>{cab_org}<th>Total</th><th>Investimento</th></tr></thead>
  <tbody>{"".join(lin_org)}</tbody></table></div>
  <div class="orgs-m">{"".join(cards_org)}</div>
  <div class="ferramentas"><h2>AÇÕES POR ALERTA DE PRAZO</h2>
    <input id="busca" type="search" placeholder="Buscar ação, órgão, responsável…" aria-label="Buscar nas ações">
    <div class="filtros" role="group" aria-label="Filtrar por alerta">{filtros}</div>
    <div id="contador" aria-live="polite"></div></div>
  {vazio}<p class="vazio" id="nada">Nenhuma ação encontrada com essa busca e esses filtros.</p>{"".join(secoes)}
</main>
<footer>CGAPE · Controle de Prazos — recorte: {e(info["recorte"])}</footer>
<button type="button" id="topo" aria-label="Voltar ao topo">↑</button>
<script>{_JS_HTML}</script>
</body></html>'''.encode("utf-8")


# ------------------------------------------------------------------ PDF ----
def _fontes_pdf():
    from reportlab.pdfbase import pdfmetrics
    nomes = pdfmetrics.getRegisteredFontNames()
    if "Calibri" in nomes and "Calibri-Bold" in nomes:
        return "Calibri", "Calibri-Bold"
    return "Helvetica", "Helvetica-Bold"


def _pdf(linhas, res, info):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table,
                                    TableStyle)
    e = html.escape
    F, FB = _fontes_pdf()
    c = lambda h: colors.HexColor("#" + h)
    larg_pag, alt_pag = landscape(A4)
    margem = 12 * mm
    util = larg_pag - 2 * margem

    st = {
        "cel": ParagraphStyle("cel", fontName=F, fontSize=8.5, leading=10.5, textColor=c(NAVY)),
        "celb": ParagraphStyle("celb", fontName=FB, fontSize=8.5, leading=10.5, textColor=c(NAVY)),
        "num": ParagraphStyle("num", fontName=FB, fontSize=8.5, leading=10.5, textColor=c(NAVY), alignment=1),
        "dir": ParagraphStyle("dir", fontName=FB, fontSize=8.5, leading=10.5, textColor=c(NAVY), alignment=2),
        "cabc": ParagraphStyle("cabc", fontName=FB, fontSize=8, leading=10, textColor=colors.white, alignment=1),
        "cabd": ParagraphStyle("cabd", fontName=FB, fontSize=8, leading=10, textColor=colors.white, alignment=2),
        "pq": ParagraphStyle("pq", fontName=F, fontSize=7.5, leading=9.2, textColor=c(CINZA)),
        "cab": ParagraphStyle("cab", fontName=FB, fontSize=8, leading=10, textColor=colors.white),
        "sec": ParagraphStyle("sec", fontName=FB, fontSize=12, leading=15, textColor=c(NAVY), spaceBefore=8, spaceAfter=4),
        "kv": ParagraphStyle("kv", fontName=FB, fontSize=18, leading=21, alignment=1),
        "kr": ParagraphStyle("kr", fontName=FB, fontSize=7, leading=9, alignment=1, textColor=c(CINZA)),
        "ks": ParagraphStyle("ks", fontName=F, fontSize=6.8, leading=8.4, alignment=1, textColor=c("8A93A3")),
    }

    def topo(cv, doc):
        cv.saveState()
        cv.setFillColor(c(NAVY))
        cv.rect(0, alt_pag - 20 * mm, larg_pag, 20 * mm, stroke=0, fill=1)
        cv.setFillColor(c(VERMELHO))
        cv.rect(0, alt_pag - 21.5 * mm, larg_pag, 1.5 * mm, stroke=0, fill=1)
        cv.setFillColor(c("4FA3E0"))
        cv.setFont(FB, 7.5)
        cv.drawString(margem, alt_pag - 7 * mm, "GOVERNO DO ESTADO DA BAHIA · CGAPE")
        cv.setFillColor(colors.white)
        cv.setFont(FB, 15)
        cv.drawString(margem, alt_pag - 13.5 * mm, TITULO)
        cv.setFillColor(c("C9D6EA"))
        cv.setFont(F, 8.5)
        sub = f"{SUBTITULO} · gerado em {info['gerado']}"
        if info["atualizado"]:
            sub += f" · cadastro atualizado em {info['atualizado']}"
        cv.drawString(margem, alt_pag - 18 * mm, sub)
        cv.setFillColor(c(CINZA))
        cv.setFont(F, 7.5)
        cv.drawString(margem, 7 * mm, f"Recorte: {info['recorte']}"[:180])
        cv.drawRightString(larg_pag - margem, 7 * mm, f"página {doc.page}")
        cv.restoreState()

    el = []
    # indicadores
    cel_k, cores_k = [], []
    itens_k = [("AÇÕES", str(res["total"]), "no recorte", NAVY)]
    itens_k += [(a.upper(), str(res["cont"][a]), ALERTA_DESC[a], ALERTA_COR[a]) for a in ALERTAS]
    itens_k.append(("INVESTIMENTO", moeda(res["inv_total"], True) if res["n_inv"] else "—",
                    f"{res['n_inv']} de {res['total']} com valor", VERMELHO))
    for rot, val, sub, cor in itens_k:
        cel_k.append([Paragraph(e(rot), st["kr"]),
                      Paragraph(f'<font color="#{cor}">{e(val)}</font>', st["kv"]),
                      Paragraph(e(sub), st["ks"])])
        cores_k.append(cor)
    larg_k = util / len(itens_k)
    tk = Table([[Table([[p] for p in cel], colWidths=[larg_k - 6]) for cel in cel_k]],
               colWidths=[larg_k] * len(itens_k))
    est_k = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (-1, -1), c("F2F4F7")),
             ("BOX", (0, 0), (-1, -1), 0.5, c("D8DCE3")), ("INNERGRID", (0, 0), (-1, -1), 2, colors.white)]
    for i, cor in enumerate(cores_k):
        est_k.append(("LINEABOVE", (i, 0), (i, 0), 3, c(cor)))
    tk.setStyle(TableStyle(est_k))
    el += [tk, Spacer(1, 6)]

    # resumo por órgão
    el.append(Paragraph("Resumo por órgão / ministério", st["sec"]))
    cab = [Paragraph("Órgão / Ministério", st["cab"])] + [Paragraph(e(a), st["cabc"]) for a in ALERTAS] + \
          [Paragraph("Total", st["cabc"]), Paragraph("Investimento", st["cabd"])]
    dados = [cab]
    for nome, o in res["orgaos"]:
        dados.append([Paragraph(e(nome), st["celb"])] +
                     [Paragraph(f'<font color="#{ALERTA_COR[a] if o[a] else "8A93A3"}">{o[a] or "–"}</font>', st["num"])
                      for a in ALERTAS] +
                     [Paragraph(str(o["total"]), st["num"]),
                      Paragraph(e(moeda(o["inv"], True)) if o["n_inv"] else "–", st["dir"])])
    lc = [util * 0.34] + [util * 0.07] * 6 + [util * 0.07, util * 0.17]
    to = Table(dados, colWidths=lc, repeatRows=1)
    to.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), c(NAVY)), ("LINEBELOW", (0, 0), (-1, 0), 2, c(VERMELHO)),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, c("F7F8FA")]),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, c("D8DCE3")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-2, -1), "CENTER"), ("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    el += [to, Spacer(1, 4)]

    # ações por alerta
    lc = [util * 0.36, util * 0.18, util * 0.13, util * 0.12, util * 0.11, util * 0.10]
    for a, lst in grupos(linhas):
        cor = ALERTA_COR[a]
        tit = Paragraph(f'<font color="#{cor}">{e(a.upper())}</font> · {len(lst)} '
                        f'{"ação" if len(lst) == 1 else "ações"} <font size="9" color="#{CINZA}">({e(ALERTA_DESC[a])})</font>',
                        st["sec"])
        dados = [[Paragraph(t, st["cab"]) for t in ("Ação", "Órgão · Eixo · Projeto", "Prazo", "Situação",
                                                    "Status · Responsável", "Investimento")]]
        for x in lst:
            desc = f'<b>{e(x["descricao"])}</b>'
            if x["observacoes"]:
                obs = x["observacoes"] if len(x["observacoes"]) <= 300 else x["observacoes"][:297] + "…"
                desc += f'<br/><font size="7.5" color="#{CINZA}">Obs.: {e(obs)}</font>'
            prazo = e(x["data_txt"]) + (f'<br/><font size="7.5" color="#{CINZA}">{e(x["prazo_informado"])}</font>'
                                        if x["prazo_informado"] else "")
            dados.append([
                Paragraph(desc, st["cel"]),
                Paragraph(e(" · ".join(v for v in (x["ministerio"], x["eixo"], x["projeto"]) if v)), st["pq"]),
                Paragraph(prazo, st["cel"]),
                Paragraph(f'<font color="#{cor}"><b>{e(x["situacao"])}</b></font>', st["cel"]),
                Paragraph(e(x["status"] or "—") + f'<br/><font size="7.5" color="#{CINZA}">{e(x["responsavel"])}</font>',
                          st["cel"]),
                Paragraph(e(moeda(x["investimento"])) if x["investimento"] is not None else "–", st["cel"]),
            ])
        t = Table(dados, colWidths=lc, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), c(NAVY)), ("LINEBELOW", (0, 0), (-1, 0), 2, c(cor)),
            ("LINEBEFORE", (0, 1), (0, -1), 3, c(cor)),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, c("F7F8FA")]),
            ("LINEBELOW", (0, 1), (-1, -1), 0.4, c("D8DCE3")), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        el.append(KeepTogether([tit, t]) if len(lst) <= 3 else tit)
        if len(lst) > 3:
            el.append(t)
    if not linhas:
        el.append(Paragraph("Nenhuma ação no recorte.", st["cel"]))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=margem, rightMargin=margem,
                            topMargin=26 * mm, bottomMargin=13 * mm, title="Relatório · Controle de Prazos",
                            author="CGAPE")
    doc.build(el, onFirstPage=topo, onLaterPages=topo)
    return buf.getvalue()


# ----------------------------------------------------------------- PPTX ----
def _pptx(linhas, res, info):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Pt

    try:
        import identidade_visual as iv
        fonte = iv.FONTE_PPTX
    except Exception:
        fonte = "Calibri"
    rgb = lambda h: RGBColor.from_string(h)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    W, H = prs.slide_width, prs.slide_height
    cm = lambda v: Emu(int(v * 360000))
    vazio = prs.slide_layouts[6]

    def caixa(sl, x, y, w, h, cor=None, linha=None, forma=MSO_SHAPE.RECTANGLE):
        s = sl.shapes.add_shape(forma, x, y, w, h)
        if cor:
            s.fill.solid()
            s.fill.fore_color.rgb = rgb(cor)
        else:
            s.fill.background()
        if linha:
            s.line.color.rgb = rgb(linha)
            s.line.width = Pt(0.75)
        else:
            s.line.fill.background()
        s.shadow.inherit = False
        return s

    def texto(sl, x, y, w, h, partes, alin=PP_ALIGN.LEFT, ancora=MSO_ANCHOR.TOP):
        """partes: lista de parágrafos; cada parágrafo é lista de (txt, pt, negrito, cor)."""
        tb = sl.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = ancora
        tf.margin_left = tf.margin_right = cm(0.1)
        tf.margin_top = tf.margin_bottom = cm(0.05)
        for i, par in enumerate(partes):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = alin
            for txt, pt, neg, cor in par:
                r = p.add_run()
                r.text = txt
                r.font.size = Pt(pt)
                r.font.bold = neg
                r.font.name = fonte
                r.font.color.rgb = rgb(cor)
        return tb

    def cabecalho(sl, subtitulo):
        caixa(sl, 0, 0, W, cm(1.9), NAVY)
        caixa(sl, 0, cm(1.9), W, cm(0.14), VERMELHO)
        texto(sl, cm(0.8), cm(0.18), cm(24), cm(0.5), [[("GOVERNO DO ESTADO DA BAHIA · CGAPE", 9, True, "4FA3E0")]])
        texto(sl, cm(0.8), cm(0.62), cm(24), cm(1.2), [[(TITULO, 22, True, "FFFFFF"),
                                                        ("   " + subtitulo, 12, False, "C9D6EA")]])
        texto(sl, cm(22), cm(0.62), cm(11.1), cm(1.0), [[(f"gerado em {info['gerado']}", 10, False, "C9D6EA")]],
              alin=PP_ALIGN.RIGHT)
        texto(sl, cm(0.8), H - cm(0.75), cm(32), cm(0.6),
              [[(f"Recorte: {info['recorte']}"[:200], 9, False, CINZA)]])

    # --- slide 1: visão geral ---
    sl = prs.slides.add_slide(vazio)
    cabecalho(sl, "Visão geral")
    itens_k = [("AÇÕES NO RECORTE", str(res["total"]), "", NAVY)]
    itens_k += [(a.upper(), str(res["cont"][a]), ALERTA_DESC[a], ALERTA_COR[a]) for a in ALERTAS]
    x0, y0, gap = cm(0.8), cm(2.6), cm(0.3)
    lk = (W - 2 * x0 - gap * (len(itens_k) - 1)) // len(itens_k)
    for i, (rot, val, sub, cor) in enumerate(itens_k):
        x = x0 + i * (lk + gap)
        caixa(sl, x, y0, lk, cm(2.9), "F2F4F7")
        caixa(sl, x, y0, lk, cm(0.12), cor)
        texto(sl, x, y0 + cm(0.25), lk, cm(2.6),
              [[(rot, 9, True, CINZA)], [(val, 30, True, cor)], [(sub, 8, False, "8A93A3")]],
              alin=PP_ALIGN.CENTER)
    yi = y0 + cm(3.3)
    caixa(sl, x0, yi, W - 2 * x0, cm(1.3), "FDECEC")
    texto(sl, x0 + cm(0.4), yi, W - 2 * x0 - cm(0.8), cm(1.3),
          [[("INVESTIMENTO TOTAL  ", 11, True, CINZA),
            (moeda(res["inv_total"]) if res["n_inv"] else "não informado", 20 if res["n_inv"] else 14, True, VERMELHO),
            (f"    {res['n_inv']} de {res['total']} ações com valor informado", 10, False, CINZA)]],
          ancora=MSO_ANCHOR.MIDDLE)
    # tabela por órgão (até 9 linhas aqui; o resto segue em slides próprios)
    orgaos = res["orgaos"]

    def tabela_orgaos(sl, y, lst):
        cols = ["Órgão / Ministério"] + ALERTAS + ["Total", "Investimento"]
        larg = [cm(12.77)] + [cm(2.2)] * 6 + [cm(1.9), cm(4.4)]
        tb = sl.shapes.add_table(len(lst) + 1, len(cols), x0, y, sum(larg, Emu(0)), cm(0.75) * (len(lst) + 1)).table
        for j, wj in enumerate(larg):
            tb.columns[j].width = wj
        for j, ctxt in enumerate(cols):
            _cel(tb.cell(0, j), ctxt, 10, True, "FFFFFF", NAVY if j not in range(1, 7) else ALERTA_COR[ALERTAS[j - 1]],
                 PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
        for i, (nome, o) in enumerate(lst, 1):
            fundo = "FFFFFF" if i % 2 else "F7F8FA"
            _cel(tb.cell(i, 0), nome, 10, True, NAVY, fundo)
            for j, a in enumerate(ALERTAS, 1):
                _cel(tb.cell(i, j), str(o[a]) if o[a] else "–", 11, bool(o[a]),
                     ALERTA_COR[a] if o[a] else "8A93A3", fundo, PP_ALIGN.CENTER)
            _cel(tb.cell(i, 7), str(o["total"]), 11, True, NAVY, fundo, PP_ALIGN.CENTER)
            _cel(tb.cell(i, 8), moeda(o["inv"], True) if o["n_inv"] else "–", 10, True, NAVY, fundo, PP_ALIGN.RIGHT)

    def _cel(cel, txt, pt, neg, cor, fundo, alin=PP_ALIGN.LEFT):
        cel.fill.solid()
        cel.fill.fore_color.rgb = rgb(fundo)
        cel.margin_left = cel.margin_right = cm(0.15)
        cel.margin_top = cel.margin_bottom = cm(0.05)
        cel.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cel.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = alin
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(pt)
        r.font.bold = neg
        r.font.name = fonte
        r.font.color.rgb = rgb(cor)

    yt = yi + cm(1.7)
    texto(sl, x0, yt, cm(20), cm(0.7), [[("RESUMO POR ÓRGÃO / MINISTÉRIO", 11, True, CINZA)]])
    POR_SLIDE_ORG = 7
    tabela_orgaos(sl, yt + cm(0.7), orgaos[:POR_SLIDE_ORG])
    resto = orgaos[POR_SLIDE_ORG:]
    while resto:
        sl = prs.slides.add_slide(vazio)
        cabecalho(sl, "Resumo por órgão (continuação)")
        tabela_orgaos(sl, cm(2.6), resto[:13])
        resto = resto[13:]

    # --- slides das ações: fluxo contínuo por alerta (o selo do alerta abre
    # cada grupo e se repete, "continuação", quando o grupo vira de slide) ---
    def corta(t, n):
        t = " ".join(str(t or "").split())
        return t if len(t) <= n else t[:n - 1].rstrip() + "…"

    Y_INI, Y_FIM = cm(2.4), H - cm(1.0)
    ALT_CAB, ALT_ITEM = cm(1.05), cm(2.15)
    estado = {"sl": None, "y": Y_FIM}

    def novo_slide():
        estado["sl"] = prs.slides.add_slide(vazio)
        cabecalho(estado["sl"], "Ações por alerta de prazo")
        estado["y"] = Y_INI

    def selo_grupo(a, n, cont):
        sl, y, cor = estado["sl"], estado["y"], ALERTA_COR[a]
        selo = caixa(sl, x0, y, cm(4.2), cm(0.75), cor, forma=MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = selo.text_frame
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = a.upper()
        r.font.size, r.font.bold, r.font.name = Pt(13), True, fonte
        r.font.color.rgb = rgb("FFFFFF")
        texto(sl, x0 + cm(4.4), y, cm(26), cm(0.75),
              [[(f"{n} {'ação' if n == 1 else 'ações'}", 13, True, NAVY),
                (f"  ·  {ALERTA_DESC[a]}" + ("  ·  continuação" if cont else ""), 11, False, CINZA)]],
              ancora=MSO_ANCHOR.MIDDLE)
        estado["y"] += ALT_CAB

    for a, lst in grupos(linhas):
        cor = ALERTA_COR[a]
        for k, x in enumerate(lst):
            precisa = ALT_ITEM + (ALT_CAB if k == 0 else 0)
            if estado["sl"] is None or estado["y"] + precisa > Y_FIM:
                novo_slide()
                selo_grupo(a, len(lst), k > 0)
            elif k == 0:
                selo_grupo(a, len(lst), False)
            sl, y, alt = estado["sl"], estado["y"], ALT_ITEM
            caixa(sl, x0, y, W - 2 * x0, alt - cm(0.2), "FFFFFF", linha="D8DCE3")
            caixa(sl, x0, y, cm(0.18), alt - cm(0.2), cor)
            trilha = " · ".join(v for v in (x["ministerio"], x["eixo"], x["projeto"]) if v)
            meta = f"Status: {x['status'] or '—'}   ·   Responsável: {x['responsavel']}"
            texto(sl, x0 + cm(0.4), y + cm(0.05), cm(23.4), alt - cm(0.3),
                  [[(corta(trilha, 140), 9, False, CINZA)],
                   [(corta(x["descricao"], 170) or "(sem descrição)", 12, True, NAVY)],
                   [(meta, 9, False, CINZA)]])
            dir_ = [[(x["situacao"], 14, True, cor)],
                    [("Prazo: ", 9, False, CINZA), (x["data_txt"], 10, True, NAVY)]]
            if x["prazo_informado"]:
                dir_[1].append((f"  ({corta(x['prazo_informado'], 30)})", 8, False, CINZA))
            if x["investimento"] is not None:
                dir_.append([(moeda(x["investimento"]), 11, True, VERMELHO)])
            texto(sl, W - x0 - cm(8.2), y + cm(0.05), cm(8.0), alt - cm(0.3), dir_, alin=PP_ALIGN.RIGHT)
            estado["y"] += alt
    if not linhas:
        texto(sl, x0, cm(8), cm(20), cm(1), [[("Nenhuma ação no recorte.", 14, False, CINZA)]])

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
