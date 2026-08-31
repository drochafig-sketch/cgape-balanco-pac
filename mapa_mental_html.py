"""Monta o HTML/CSS/JS de uma página só com o mapa mental interativo do
BALANÇO PAC.

Portado do mesmo componente já usado no projeto "Controle de Prazos"
(pasta "19. APRESENTAÇÃO GABINETE - PAUTAS REUNIÕES\\Controle de prazos",
arquivo mapa_html.py): mesma mecânica de árvore horizontal com pan/zoom,
busca, filtro por órgão, painel de detalhes e legenda — só que a árvore
aqui é Secretaria/Órgão > Objeto > Ação (item), a cor de cada nó segue a
FASE da ação (mesma paleta já usada no resto do painel: Captação de
Recurso, Licitação, Execução do Objeto, Concluída), e o restante do design
(cores, tipografia, cartões) usa exatamente o mesmo design system do
painel principal — mesmas variáveis (--cor-fundo, --cor-card,
--cor-acento-teal etc.), sem tema claro/escuro à parte: é sempre o mesmo
visual do resto do BALANÇO PAC.

Usado por _api_mapa_mental (ver "CGAPE - BALANÇO PAC.py"): recebe a árvore
já filtrada pelos mesmos filtros do painel principal, e devolve uma página
HTML autônoma (sem depender de nenhum outro arquivo/CSS/JS do painel) —
mostrada pelo botão "MAPA MENTAL" num iframe em tela cheia dentro do
próprio painel (ver #mapa-mental-overlay em montar_html_painel).
"""
from __future__ import annotations

import json


def _json_seguro(obj) -> str:
    return json.dumps(obj, ensure_ascii=False).replace("</", "<\\/")


# Mesmas variáveis de design do painel principal (ver :root em
# montar_html_painel) — repetidas aqui porque esta página é um documento
# HTML autônomo à parte (carregada num <iframe>, isolada de propósito do
# CSS/JS do painel). Qualquer ajuste de paleta feito lá deveria ser
# replicado aqui também.
CSS = """
:root{
  --cor-fundo:#303030; --cor-card:#353B47; --cor-card-elevado:#41454F;
  --cor-texto-primario:#FFFFFF; --cor-texto-secundario:#B3BAC9; --cor-texto-terciario:#858B97;
  --cor-acento-mint:#B8EAE1; --cor-acento-teal:#72B4AE; --cor-acento-teal-hover:#8AC4BF;
  --cor-acento-peach:#EEB489; --cor-acento-gold:#E0AB45; --cor-acento-gold-claro:#F2CE83;
  --borda-card:rgba(114,180,174,.25);
  --fase-vermelho:#BB6060; --fase-amarelo:#BC9E2C; --fase-verde:#49925C; --fase-azul:#4E92BA;
  --atrasada:#E2574C;
  --raio-sm:4px; --raio-md:8px; --raio-lg:16px;
  --sombra-card:0px 5px 40px 0px rgba(9,14,21,.16);
  --transicao-rapida:.15s ease; --transicao-padrao:.25s ease-in-out;
  --fonte:"Segoe UI","Roboto",system-ui,-apple-system,sans-serif;
}
*{box-sizing:border-box;}
html,body{margin:0;height:100%;overflow:hidden;}
body{
  font-family:var(--fonte); color:var(--cor-texto-secundario); background:var(--cor-fundo);
  display:flex; flex-direction:column;
}
button{font-family:inherit;cursor:pointer;}

#topo{
  background:var(--cor-fundo); border-bottom:1px solid var(--cor-card-elevado);
  color:var(--cor-texto-primario); padding:14px 22px; display:flex; align-items:center; gap:18px;
  flex:0 0 auto;
}
#topo h1{font-size:18px; margin:0; font-weight:700; letter-spacing:.2px; color:var(--cor-texto-primario);}
#topo .sub{font-size:12px; color:var(--cor-texto-secundario); margin-top:2px;}
#fechar-mapa{
  flex:0 0 auto; width:34px; height:34px; border-radius:50%; border:1px solid var(--borda-card);
  background:var(--cor-card-elevado); color:var(--cor-texto-primario); display:flex; align-items:center; justify-content:center;
  transition:background var(--transicao-rapida);
}
#fechar-mapa:hover{background:var(--cor-acento-teal-hover); color:#1A1A1A;}
#relogio{margin-left:auto; text-align:right; font-size:12px; color:var(--cor-texto-secundario);}
#relogio b{display:block; font-size:15px; color:var(--cor-texto-primario); font-variant-numeric:tabular-nums;}

#dashboard{
  display:flex; gap:10px; padding:10px 22px; background:var(--cor-card); border-bottom:1px solid var(--cor-card-elevado);
  flex:0 0 auto; overflow-x:auto;
}
.chip{
  display:flex; align-items:center; gap:8px; padding:7px 12px; border-radius:999px;
  border:1px solid var(--borda-card); background:var(--cor-card-elevado); font-size:12.5px; font-weight:600; color:var(--cor-texto-primario);
  white-space:nowrap; user-select:none; transition:all var(--transicao-rapida);
}
.chip .dot{width:10px; height:10px; border-radius:50%;}
.chip .n{background:rgba(0,0,0,.18); border-radius:999px; padding:1px 7px; font-size:11px;}
.chip:hover{border-color:var(--cor-acento-teal);}
.btn-limpar{
  flex:0 0 auto; margin-left:auto; align-self:center; padding:7px 14px; border-radius:999px;
  border:1px solid var(--borda-card); background:var(--cor-card-elevado); color:var(--cor-texto-primario);
  font-size:12.5px; font-weight:600; white-space:nowrap; transition:all var(--transicao-rapida);
}
.btn-limpar:hover:not(:disabled){background:var(--cor-acento-teal-hover); color:#1A1A1A;}
.btn-limpar:disabled{opacity:.4; cursor:not-allowed;}

#radar{
  flex:0 0 auto; padding:10px 22px; background:var(--cor-card-elevado); border-bottom:1px solid var(--cor-card-elevado);
  display:flex; gap:10px; align-items:center; overflow-x:auto;
}
#radar .rotulo{font-size:11.5px; font-weight:700; color:var(--cor-texto-secundario); text-transform:uppercase; letter-spacing:.5px; flex:0 0 auto;}
.radar-card{
  flex:0 0 auto; min-width:190px; background:var(--cor-card); border:1px solid var(--borda-card); border-left:5px solid var(--cor-texto-terciario);
  border-radius:var(--raio-md); padding:7px 10px; font-size:12px; cursor:pointer; transition:border-color var(--transicao-rapida);
}
.radar-card:hover{border-color:var(--cor-acento-teal);}
.radar-card .rp{font-weight:700; font-size:12.5px; margin-bottom:2px; color:var(--cor-texto-primario);}
.radar-card .rd{color:var(--cor-texto-secundario); font-size:11px;}
.radar-card .rc{font-weight:700; margin-top:3px; font-variant-numeric:tabular-nums;}

#corpo{flex:1 1 auto; display:flex; min-height:0; position:relative;}

#lateral{
  width:250px; flex:0 0 auto; background:var(--cor-card); border-right:1px solid var(--cor-card-elevado);
  padding:14px; overflow-y:auto; overflow-x:hidden; transition:width .22s ease,padding .22s ease;
}
#corpo.lateral-recolhido #lateral{width:0; padding-left:0; padding-right:0; border-right:none;}
#lateral-toggle{
  position:absolute; top:14px; left:250px; transform:translateX(-50%); z-index:4;
  width:26px; height:26px; border-radius:50%; border:1px solid var(--borda-card);
  background:var(--cor-card-elevado); color:var(--cor-texto-primario); display:flex; align-items:center; justify-content:center;
  padding:0; font-size:14px; line-height:1; box-shadow:var(--sombra-card); transition:left .22s ease,background var(--transicao-rapida);
}
#lateral-toggle:hover{background:var(--cor-acento-teal-hover); color:#1A1A1A;}
#corpo.lateral-recolhido #lateral-toggle{left:0;}
#lateral h3{font-size:11.5px; text-transform:uppercase; letter-spacing:.5px; color:var(--cor-texto-terciario); margin:14px 0 8px;}
#lateral h3:first-child{margin-top:0;}
#busca{width:100%; padding:8px 10px; border:1px solid var(--borda-card); border-radius:var(--raio-md); font-size:13px; background:var(--cor-card-elevado); color:var(--cor-texto-primario);}
#busca:focus{outline:none; border-color:var(--cor-acento-teal);}
#busca::placeholder{color:var(--cor-texto-terciario);}
.filtro-org label{display:flex; align-items:center; gap:7px; font-size:13px; padding:4px 0; cursor:pointer; color:var(--cor-texto-secundario);}
.filtro-org input{accent-color:var(--cor-acento-teal);}
#legenda-item{font-size:11.5px; color:var(--cor-texto-secundario); line-height:1.7;}
#legenda-item span{display:inline-block; width:9px; height:9px; border-radius:50%; margin-right:6px;}
#btn-reset{margin-top:12px; width:100%; padding:8px; border-radius:var(--raio-md); border:1px solid var(--borda-card); background:var(--cor-card-elevado); color:var(--cor-texto-primario); font-size:12.5px; font-weight:600; transition:background var(--transicao-rapida);}
#btn-reset:hover{background:var(--cor-acento-teal-hover); color:#1A1A1A;}

#viewport{flex:1 1 auto; position:relative; overflow:hidden; cursor:grab; touch-action:none; background:
  radial-gradient(circle,rgba(255,255,255,.05) 1px,transparent 1px) 0 0/22px 22px, var(--cor-fundo);}
#viewport.arrastando{cursor:grabbing;}
#canvas{position:absolute; top:0; left:0; transform-origin:0 0;}
#svg-linhas{position:absolute; top:0; left:0; overflow:visible; pointer-events:none; transform-origin:0 0;}
#svg-linhas path{stroke:var(--cor-card-elevado);}

.no{
  position:absolute; background:var(--cor-card); border:1px solid var(--borda-card); border-radius:var(--raio-lg);
  padding:8px 12px; box-shadow:var(--sombra-card); transition:left .28s,top .28s,opacity .2s,border-color var(--transicao-rapida);
  user-select:none; color:var(--cor-texto-primario);
}
.no.raiz{background:var(--cor-acento-teal); color:#16211F; border:none; padding:14px 20px; width:260px; cursor:pointer;}
.no.raiz .tit{font-size:15px; font-weight:700;}
.no.raiz .sub{font-size:11px; color:#1A1A1A; opacity:.75; margin-top:2px;}
.no.orgao{border-left:5px solid var(--cor-acento-teal); font-weight:700; font-size:13px; cursor:pointer; width:220px;}
.no.executor{border-left:5px solid var(--cor-acento-mint); font-weight:600; font-size:12.5px; cursor:pointer; width:200px;}
.no.objeto{border-left:5px solid var(--cor-texto-terciario); font-weight:600; font-size:12.5px; cursor:pointer; width:230px;}
.no.item{border-left:5px solid var(--cor-texto-terciario); font-size:12px; width:250px; cursor:pointer;}
.no.item .desc{font-weight:600; margin-bottom:4px; line-height:1.35; color:var(--cor-texto-primario);}
.no.item .meta{color:var(--cor-texto-secundario); font-size:11px; margin-bottom:4px;}
.no .toggle{position:absolute; right:8px; top:50%; transform:translateY(-50%); font-size:10px; color:var(--cor-texto-terciario);}
.badge{display:inline-block; padding:2px 8px; border-radius:999px; color:#fff; font-size:10.5px; font-weight:700; letter-spacing:.2px;}
.contagem{font-variant-numeric:tabular-nums; font-weight:700; font-size:11.5px; margin-top:3px;}

.no.alerta-atrasada{border-left-color:var(--atrasada); box-shadow:0 0 0 1px var(--atrasada),var(--sombra-card);}
.no.alerta-semprevisao{border-left-color:var(--cor-texto-terciario);}

.no.destaque{box-shadow:0 0 0 3px var(--cor-acento-gold); z-index:5;}

#zoom-ctl{position:absolute; right:16px; bottom:16px; display:flex; flex-direction:column; gap:6px; z-index:3;}
#zoom-ctl button{width:34px; height:34px; border-radius:var(--raio-md); border:1px solid var(--borda-card); background:var(--cor-card-elevado); color:var(--cor-texto-primario); font-size:16px; box-shadow:var(--sombra-card); transition:background var(--transicao-rapida);}
#zoom-ctl button:hover{background:var(--cor-acento-teal-hover); color:#1A1A1A;}

#painel{
  position:absolute; top:0; right:-380px; width:360px; height:100%; background:var(--cor-card); border-left:1px solid var(--borda-card);
  box-shadow:var(--sombra-card); transition:right .25s; padding:20px; overflow-y:auto; z-index:6; color:var(--cor-texto-secundario);
}
#painel.aberto{right:0;}
#painel .fechar{position:absolute; top:14px; right:14px; border:none; background:none; font-size:18px; color:var(--cor-texto-terciario);}
#painel .trilha{font-size:11.5px; color:var(--cor-texto-terciario); margin-bottom:8px;}
#painel h2{font-size:16px; margin:0 0 12px; line-height:1.3; text-wrap:balance; color:var(--cor-texto-primario);}
#painel h3{font-size:11px; text-transform:uppercase; letter-spacing:.4px; color:var(--cor-texto-terciario); margin:16px 0 8px; border-top:1px solid var(--cor-card-elevado); padding-top:12px;}
#painel h3:first-of-type{border-top:none; padding-top:0; margin-top:12px;}
#painel .campo{margin-bottom:12px;}
#painel .campo b{display:block; font-size:11px; text-transform:uppercase; letter-spacing:.4px; color:var(--cor-texto-terciario); margin-bottom:3px;}
#painel .campo span{font-size:13.5px; color:var(--cor-texto-primario);}
#painel .contagem-grande{font-size:22px; font-weight:700; font-variant-numeric:tabular-nums;}

#painel .voltar-lista{
  display:inline-flex; align-items:center; gap:4px; background:none; border:none; padding:0; margin-bottom:12px;
  color:var(--cor-acento-teal); font-size:12px; font-weight:700; cursor:pointer;
}
#painel .voltar-lista:hover{color:var(--cor-acento-teal-hover);}
#painel .lista-rotulo{font-size:11.5px; color:var(--cor-texto-terciario); margin-bottom:10px;}
#painel .lista-acoes{display:flex; flex-direction:column; gap:8px;}
.lista-acao-item{
  background:var(--cor-card-elevado); border:1px solid var(--borda-card); border-radius:var(--raio-md);
  padding:8px 10px; cursor:pointer; transition:border-color var(--transicao-rapida);
}
.lista-acao-item:hover{border-color:var(--cor-acento-teal);}
.la-desc{font-weight:600; font-size:12.5px; color:var(--cor-texto-primario); margin-bottom:4px; line-height:1.3;}
.la-meta{color:var(--cor-texto-secundario); font-size:11px; margin-bottom:5px;}
.la-contagem{font-size:11px; font-weight:700; margin-top:4px; font-variant-numeric:tabular-nums;}

@media (prefers-reduced-motion:reduce){
  .no{transition:none;} #painel{transition:none;} .radar-card,.chip{transition:none;}
}
"""


def montar_html_mapa_mental(arvore: dict, meta: dict) -> str:
    dados_js = _json_seguro(arvore)
    gerado_em = meta.get("gerado_em", "")
    total = meta.get("total", 0)
    titulo = meta.get("titulo", "BALANÇO PAC - BAHIA")
    titulo_aba = meta.get("titulo_aba", titulo)
    subtitulo = meta.get("subtitulo", "")

    html = f"""<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{titulo_aba}</title>
<style>{CSS}</style>
</head>
<body>

<div id="topo">
  <button id="fechar-mapa" title="Voltar aos filtros" aria-label="Voltar aos filtros">
    <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <line x1="19" y1="12" x2="5" y2="12"></line>
      <polyline points="12 19 5 12 12 5"></polyline>
    </svg>
  </button>
  <div class="titulo-wrap">
    <h1>{titulo} &middot; Mapa Mental</h1>
    <div class="sub">{subtitulo} &middot; {total} ações no recorte &middot; gerado em {gerado_em}</div>
  </div>
  <div id="relogio"><span>hoje</span><b id="relogio-data"></b></div>
</div>

<div id="dashboard"></div>

<div id="radar">
  <div class="rotulo">Radar de prazos (previsão de conclusão atual)</div>
  <div id="radar-cards" style="display:flex; gap:10px;"></div>
</div>

<div id="corpo">
  <div id="lateral">
    <h3>Buscar</h3>
    <input id="busca" placeholder="órgão, objeto, item, executor...">
    <h3>Secretarias/Órgãos</h3>
    <div class="filtro-org" id="filtro-org"></div>
    <h3>Legenda</h3>
    <div id="legenda-item"></div>
    <button id="btn-reset">Expandir tudo / recolher tudo</button>
  </div>
  <button id="lateral-toggle" title="Recolher painel" aria-label="Recolher painel">&lsaquo;</button>

  <div id="viewport">
    <svg id="svg-linhas"></svg>
    <div id="canvas"></div>
    <div id="zoom-ctl">
      <button id="zoom-mais">+</button>
      <button id="zoom-menos">&minus;</button>
      <button id="zoom-fit">&#8862;</button>
    </div>
  </div>

  <div id="painel">
    <button class="fechar" id="painel-fechar">&times;</button>
    <div id="painel-conteudo"></div>
  </div>
</div>

<script>
const DADOS = {dados_js};
</script>
<script>
{JS}
</script>
</body>
</html>"""
    return html


JS = r"""
const FASE_COR = {
  'CAPTAÇÃO DE RECURSO':'var(--fase-vermelho)', 'LICITAÇÃO':'var(--fase-amarelo)',
  'EXECUÇÃO DO OBJETO':'var(--fase-verde)', 'CONCLUÍDA':'var(--fase-azul)'
};
const FASE_ORDEM = ['CAPTAÇÃO DE RECURSO','LICITAÇÃO','EXECUÇÃO DO OBJETO','CONCLUÍDA'];

function corResolvida(nomeVar){
  const v = nomeVar.match(/--[a-z-]+/)[0];
  return getComputedStyle(document.documentElement).getPropertyValue(v).trim();
}

function corFase(fase){
  return corResolvida(FASE_COR[fase] || 'var(--cor-texto-terciario)');
}

// "situação de prazo": independente da fase (que já dá a cor do nó) — diz
// se a ação está com a Previsão de Conclusão Atual vencida, sem previsão
// definida, ou dentro do prazo. Concluída não entra nessa conta.
function situacaoItem(d){
  if(d.fase === 'CONCLUÍDA') return 'Concluída';
  if(!d.prazo_atual_iso) return 'Sem previsão';
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const prazo = new Date(d.prazo_atual_iso + 'T00:00:00');
  return prazo < hoje ? 'Atrasada' : 'Em dia';
}

function diasRestantes(d){
  if(!d.prazo_atual_iso) return null;
  const agora = new Date();
  const prazo = new Date(d.prazo_atual_iso + 'T23:59:59');
  return prazo - agora;
}

function formatarContagem(d){
  const sit = situacaoItem(d);
  if(sit === 'Concluída') return 'Concluída';
  if(sit === 'Sem previsão') return 'sem previsão definida';
  const ms = diasRestantes(d);
  const dias = Math.floor(Math.abs(ms)/86400000);
  return ms < 0 ? `previsão vencida há ${dias} dia(s)` : `faltam ${dias} dia(s) (previsão)`;
}

function corContagem(d){
  const sit = situacaoItem(d);
  if(sit === 'Atrasada') return corResolvida('var(--atrasada)');
  if(sit === 'Concluída') return corResolvida('var(--fase-azul)');
  return corResolvida('var(--cor-texto-terciario)');
}

// ---------- coleta todos os itens (folhas) ----------
let todosItens = [];
(function coletar(no){
  if(no.tipo === 'item') todosItens.push(no.dados);
  no.filhos.forEach(coletar);
})(DADOS);

// ---------- estado ----------
let colapsado = new Set();     // ids de nó colapsado (guardo por caminho)
let filtroFases = new Set();   // fases ativas no filtro (vazio = todas)
let orgaosDesligados = new Set();
let termoBusca = '';
let zoom = 1, panX = 60, panY = 40;

// começa com a árvore toda recolhida: só a raiz fica aberta, revelando
// apenas as secretarias/órgãos (executor e objeto ficam escondidos até o
// usuário clicar pra abrir cada ramo — e o objeto nunca chega a expandir
// dentro do mapa: ver o clique especial para ele mais abaixo)
function definirColapsoPadrao(no, caminho){
  no._id = caminho;
  if(no.tipo !== 'raiz') colapsado.add(caminho);
  no.filhos.forEach((f,i)=>definirColapsoPadrao(f, caminho+'/'+i));
}
definirColapsoPadrao(DADOS, 'r');

// ---------- layout tipo árvore horizontal ----------
const COL_W = 320, ROW_H = 96, ROW_GAP = 14;

function visivel(no){
  return !colapsado.has(no._id) && filhosAtivos(no).length>0;
}

// pior situação de prazo entre os itens da subárvore de um nó — assim um
// ramo recolhido (secretaria/objeto) também acende vermelho quando esconde
// alguma ação atrasada, sem precisar abrir tudo. Quando nenhuma ação da
// subárvore está atrasada e todas as ativas estão sem previsão definida, o
// ramo herda a cor neutra de "sem previsão".
function calcularPiorSituacao(no){
  if(no.tipo === 'item'){
    no._pior = situacaoItem(no.dados);
    return [no._pior];
  }
  let situacoesFolhas = [];
  no.filhos.forEach(f=>{ situacoesFolhas = situacoesFolhas.concat(calcularPiorSituacao(f)); });
  if(situacoesFolhas.includes('Atrasada')) no._pior = 'Atrasada';
  else if(situacoesFolhas.length>0 && situacoesFolhas.every(s=>s==='Sem previsão'||s==='Concluída') && situacoesFolhas.includes('Sem previsão')) no._pior = 'Sem previsão';
  else no._pior = null;
  return situacoesFolhas;
}

// ---------- filtros ----------
function correspondeAoFiltro(itemDados){
  if(filtroFases.size>0 && !filtroFases.has(itemDados.fase)) return false;
  if(orgaosDesligados.has(itemDados.secretaria)) return false;
  if(termoBusca){
    const alvo = (itemDados.descricao+' '+itemDados.objeto+' '+itemDados.secretaria+' '+itemDados.executor+' '+itemDados.municipio+' '+String(itemDados.item||'')).toLowerCase();
    if(!alvo.includes(termoBusca)) return false;
  }
  return true;
}

function noAtivo(no){
  if(no.tipo === 'item') return correspondeAoFiltro(no.dados);
  return no.filhos.some(noAtivo);
}

// filhos que sobrevivem aos filtros ativos — usado para desenhar apenas o
// que corresponde à busca/filtro, suprimindo por completo o resto (em vez
// de apenas esmaecer)
function filhosAtivos(no){
  return no.filhos.filter(noAtivo);
}

// ---------- desenho ----------
const elCanvas = document.getElementById('canvas');
const elSvg = document.getElementById('svg-linhas');

function montarConteudo(no){
  if(no.tipo === 'raiz'){
    return `<div class="tit">${no.nome}</div><div class="sub">${todosItens.length} ações &middot; clique nos ramos para abrir/fechar &middot; clique no objeto para ver as ações</div>`;
  } else if(no.tipo === 'item'){
    const d = no.dados;
    const fase = d.fase || 'Sem fase';
    const cor = corFase(fase);
    return `<div class="desc">${d.descricao || d.objeto}</div>
      <div class="meta">${d.objeto}${d.municipio? ' &middot; '+d.municipio:''}${d.executor? ' &middot; '+d.executor:''}</div>
      <span class="badge" style="background:${cor}">${fase}</span>
      <div class="contagem" style="color:${corContagem(d)}" data-contagem="${d.id}">${formatarContagem(d)}</div>`;
  }
  const rotulo = no.tipo==='orgao'?' executor(es)': no.tipo==='executor'?' objeto(s)':' ação(ões)';
  const filhosTxt = filhosAtivos(no).length + rotulo;
  return `<div>${no.nome}</div><div class="meta" style="font-weight:400;color:var(--cor-texto-secundario);font-size:11px;">${filhosTxt}</div>`;
}

function montarDiv(no, top){
  const tem = filhosAtivos(no).length>0;
  let classe = 'no '+no.tipo;
  if(no._destaque) classe += ' destaque';
  if(no._pior === 'Atrasada') classe += ' alerta-atrasada';
  else if(no._pior === 'Sem previsão') classe += ' alerta-semprevisao';
  // O nó "objeto" nunca expande dentro do próprio mapa (não tem filho
  // desenhado abaixo dele) — clicar nele abre a lista de ações no painel
  // lateral, por isso a seta é uma indicação de "abrir lista" (›) fixa,
  // em vez do ▸/▾ de abrir/fechar usado nos demais ramos.
  let seta = '';
  if(no.tipo === 'objeto'){
    seta = tem ? `<span class="toggle">&rsaquo;</span>` : '';
  } else if(tem){
    seta = `<span class="toggle">${colapsado.has(no._id)?'▸':'▾'}</span>`;
  }
  return `<div class="${classe}" style="left:${no._x}px; top:${top}px;" data-id="${no._id}">${montarConteudo(no)}${seta}</div>`;
}

function larguraAltura(){
  let maxX=0, maxY=0;
  (function v(no){
    maxX = Math.max(maxX, no._x + 280);
    maxY = Math.max(maxY, no._y + 90);
    if(visivel(no)) filhosAtivos(no).forEach(v);
  })(DADOS);
  return [maxX, maxY];
}

function desenhar(){
  calcularPiorSituacao(DADOS);

  // profundidade (x) de cada nó
  (function definirX(no, prof){
    no._x = prof*COL_W;
    if(visivel(no)) filhosAtivos(no).forEach(f=>definirX(f, prof+1));
  })(DADOS, 0);

  // nós "folha" da árvore visível (itens, ou ramos colapsados) — são
  // eles que definem as linhas; medimos a altura real de cada card (o
  // texto pode quebrar em várias linhas) para as linhas não se
  // sobreporem quando um card fica mais alto que o padrão. Ramos sem
  // nenhuma ação ativa (filtrada) são suprimidos por completo aqui, em
  // vez de apenas esmaecidos — não entram na coleta nem reservam espaço
  // no layout.
  const folhas = [];
  (function coletarFolhas(no){
    if(visivel(no)) filhosAtivos(no).forEach(coletarFolhas);
    else folhas.push(no);
  })(DADOS);

  elCanvas.style.visibility = 'hidden';
  elCanvas.innerHTML = folhas.map(no=>montarDiv(no, 0)).join('');
  let cursor = 0;
  folhas.forEach(no=>{
    const el = elCanvas.querySelector(`[data-id="${no._id}"]`);
    const altura = el ? el.offsetHeight : ROW_H;
    no._y = cursor;
    cursor += Math.max(altura, ROW_H) + ROW_GAP;
  });

  // nós internos ficam centralizados entre o primeiro e o último filho
  (function centralizar(no){
    if(visivel(no)){
      const ativos = filhosAtivos(no);
      ativos.forEach(centralizar);
      const ys = ativos.map(f=>f._y);
      no._y = (Math.min(...ys)+Math.max(...ys))/2;
    }
  })(DADOS);

  const htmlAcc = [], linhasAcc = [];
  (function montar(no){
    htmlAcc.push(montarDiv(no, no._y));
    if(visivel(no)){
      filhosAtivos(no).forEach(f=>{
        const x1=no._x+ (no.tipo==='item'?250:no.tipo==='raiz'?260:no.tipo==='orgao'?220:no.tipo==='executor'?200:230), y1=no._y+26;
        const x2=f._x, y2=f._y+26;
        const midx=(x1+x2)/2;
        linhasAcc.push(`<path d="M${x1},${y1} C${midx},${y1} ${midx},${y2} ${x2},${y2}" fill="none" stroke="#5a6072" stroke-width="2"/>`);
        montar(f);
      });
    }
  })(DADOS);

  elCanvas.innerHTML = htmlAcc.join('');
  elCanvas.style.visibility = 'visible';
  const [w,h] = larguraAltura();
  elSvg.setAttribute('width', w); elSvg.setAttribute('height', h);
  elSvg.innerHTML = linhasAcc.join('');
  aplicarTransform();

  elCanvas.querySelectorAll('.no').forEach(el=>{
    el.addEventListener('click', (ev)=>{
      ev.stopPropagation();
      const id = el.dataset.id;
      const no = encontrarNo(DADOS, id);
      if(no.tipo === 'item'){ abrirPainel(no.dados); }
      else if(no.tipo === 'objeto'){ abrirListaAcoes(no); }
      else if(no.filhos.length>0){
        if(colapsado.has(id)) colapsado.delete(id); else colapsado.add(id);
        desenhar();
      }
    });
  });
}

function encontrarNo(no, id){
  if(no._id === id) return no;
  for(const f of no.filhos){ const r = encontrarNo(f, id); if(r) return r; }
  return null;
}

// ---------- pan / zoom ----------
const viewport = document.getElementById('viewport');
function aplicarTransform(){
  elCanvas.style.transform = `translate(${panX}px,${panY}px) scale(${zoom})`;
  elSvg.style.transform = `translate(${panX}px,${panY}px) scale(${zoom})`;
}

let arrastando=false, ax=0, ay=0, arrastouBastante=false;
viewport.addEventListener('mousedown', e=>{
  if(e.target.closest('.no')) return;
  arrastando=true; ax=e.clientX; ay=e.clientY; arrastouBastante=false; viewport.classList.add('arrastando');
});
window.addEventListener('mousemove', e=>{
  if(!arrastando) return;
  const dx=e.clientX-ax, dy=e.clientY-ay;
  if(Math.abs(dx)+Math.abs(dy) > 3) arrastouBastante = true;
  panX += dx; panY += dy; ax=e.clientX; ay=e.clientY;
  aplicarTransform();
});
window.addEventListener('mouseup', ()=>{arrastando=false; viewport.classList.remove('arrastando');});
viewport.addEventListener('wheel', e=>{
  e.preventDefault();
  const fator = e.deltaY<0?1.08:0.93;
  const novoZoom = Math.min(2.2, Math.max(0.25, zoom*fator));
  if(novoZoom === zoom) return;
  const rect = viewport.getBoundingClientRect();
  const mx = e.clientX - rect.left, my = e.clientY - rect.top;
  const razao = novoZoom/zoom;
  panX = mx - (mx-panX)*razao;
  panY = my - (my-panY)*razao;
  zoom = novoZoom;
  aplicarTransform();
}, {passive:false});
// telas touch: 1 dedo arrasta (pan, mesma lógica do mouse acima) e 2 dedos
// em pinça dão zoom, ancorado no ponto médio entre os dois toques (mesma
// ideia do zoom da roda do mouse, só que a "âncora" se move junto com os
// dedos em vez de ficar fixa)
let touchModo = null; // 'pan' | 'pinch'
let touchDistInicial = 0, touchZoomInicial = 1;
let touchMeioInicial = {x:0,y:0}, touchPanInicial = {x:0,y:0};

function distanciaToques(t0, t1){
  return Math.hypot(t1.clientX-t0.clientX, t1.clientY-t0.clientY);
}
function meioToques(t0, t1, rect){
  return {x:(t0.clientX+t1.clientX)/2-rect.left, y:(t0.clientY+t1.clientY)/2-rect.top};
}

viewport.addEventListener('touchstart', e=>{
  if(e.target.closest('.no')) return;
  if(e.touches.length === 1){
    touchModo = 'pan';
    arrastando = true; arrastouBastante = false;
    ax = e.touches[0].clientX; ay = e.touches[0].clientY;
    viewport.classList.add('arrastando');
  } else if(e.touches.length === 2){
    e.preventDefault();
    touchModo = 'pinch';
    arrastando = false;
    touchDistInicial = distanciaToques(e.touches[0], e.touches[1]);
    touchZoomInicial = zoom;
    const rect = viewport.getBoundingClientRect();
    touchMeioInicial = meioToques(e.touches[0], e.touches[1], rect);
    touchPanInicial = {x:panX, y:panY};
  }
}, {passive:false});

viewport.addEventListener('touchmove', e=>{
  if(touchModo === 'pan' && e.touches.length === 1){
    e.preventDefault();
    const t = e.touches[0];
    const dx = t.clientX-ax, dy = t.clientY-ay;
    if(Math.abs(dx)+Math.abs(dy) > 3) arrastouBastante = true;
    panX += dx; panY += dy; ax = t.clientX; ay = t.clientY;
    aplicarTransform();
  } else if(touchModo === 'pinch' && e.touches.length === 2){
    e.preventDefault();
    const distAtual = distanciaToques(e.touches[0], e.touches[1]);
    const novoZoom = Math.min(2.2, Math.max(0.25, touchZoomInicial*(distAtual/touchDistInicial)));
    const razao = novoZoom/touchZoomInicial;
    panX = touchMeioInicial.x - (touchMeioInicial.x-touchPanInicial.x)*razao;
    panY = touchMeioInicial.y - (touchMeioInicial.y-touchPanInicial.y)*razao;
    zoom = novoZoom;
    aplicarTransform();
  }
}, {passive:false});

function finalizarToque(e){
  if(e.touches.length === 0){
    touchModo = null; arrastando = false; viewport.classList.remove('arrastando');
  } else if(e.touches.length === 1){
    // saiu da pinça e ainda sobrou um dedo -> vira pan a partir daqui
    touchModo = 'pan'; arrastando = true; arrastouBastante = false;
    ax = e.touches[0].clientX; ay = e.touches[0].clientY;
  }
}
viewport.addEventListener('touchend', finalizarToque);
viewport.addEventListener('touchcancel', finalizarToque);

document.getElementById('zoom-mais').onclick = ()=>{zoom=Math.min(2.2,zoom*1.15); aplicarTransform();};
document.getElementById('zoom-menos').onclick = ()=>{zoom=Math.max(0.25,zoom*0.87); aplicarTransform();};
document.getElementById('zoom-fit').onclick = ()=>{zoom=1; panX=60; panY=40; aplicarTransform();};

// ---------- painel de detalhes ----------
const painel = document.getElementById('painel');
function campo(rotulo, valor){
  if(!valor) return '';
  return `<div class="campo"><b>${rotulo}</b><span>${valor}</span></div>`;
}
function abrirPainel(d, origemLista){
  const fase = d.fase || 'Sem fase';
  const cor = corFase(fase);
  const corSit = corContagem(d);
  const voltar = origemLista
    ? `<button class="voltar-lista" id="painel-voltar">&larr; Voltar à lista de ações</button>`
    : '';
  document.getElementById('painel-conteudo').innerHTML = `
    ${voltar}
    <div class="trilha">${d.secretaria} &rsaquo; ${d.executor} &rsaquo; ${d.objeto}${d.item? ' &middot; Item '+d.item:''}</div>
    <h2>${d.descricao || d.objeto}</h2>
    <div class="campo"><b>Fase</b><span class="badge" style="background:${cor}">${fase}</span></div>
    <div class="campo"><b>Situação do prazo</b><span class="contagem-grande" style="color:${corSit}">${formatarContagem(d)}</span></div>
    ${campo('Status', d.status)}
    <h3>Prazos</h3>
    ${campo('Vigência', d.vigencia)}
    ${campo('Previsão de Conclusão Atual', d.prazo_atual)}
    ${campo('Prazo de Conclusão da Fase', d.prazo_fase)}
    ${campo('Avanço da Obra', d.avanco)}
    <h3>Execução</h3>
    ${campo('Secretaria/Órgão', d.secretaria)}
    ${campo('Órgão Executor', d.executor)}
    ${campo('Gestão', d.gestao)}
    ${campo('Eixo', d.eixo)}
    ${campo('Município', d.municipio)}
    ${campo('Fonte de Recurso', d.fonte)}
    <h3>Financeiro</h3>
    ${campo('Valor Contratado', d.valor_contratado)}
    ${campo('Financiamento', d.financiamento)}
    ${campo('Apoiado (OGU)', d.apoiado)}
    ${campo('Contrapartida', d.contrapartida)}
    ${campo('Complementar', d.complementar)}
    ${campo('Investimento Total', d.investimento_total)}
    <h3>Acompanhamento</h3>
    ${campo('Pendência/Tarefa', d.pendencia)}
    ${campo('Providências', d.providencias)}
    ${campo('Próximos Passos', d.proximos_passos)}
  `;
  if(origemLista){
    document.getElementById('painel-voltar').onclick = ()=>abrirListaAcoes(origemLista);
  }
  painel.classList.add('aberto');
}

// Nó "objeto": em vez de expandir dentro do mapa, mostra no mesmo painel
// lateral a lista das ações (itens) desse objeto — clicar numa delas abre
// a Ficha Cadastral completa (abrirPainel), com um link para voltar aqui.
function linhaAcaoHtml(d){
  const cor = corContagem(d);
  const fase = d.fase || 'Sem fase';
  return `<div class="lista-acao-item" data-id="${d.id}">
    <div class="la-desc">${d.descricao || d.objeto}</div>
    ${d.municipio? `<div class="la-meta">${d.municipio}</div>`:''}
    <span class="badge" style="background:${corFase(fase)}">${fase}</span>
    <div class="la-contagem" style="color:${cor}">${formatarContagem(d)}</div>
  </div>`;
}
function abrirListaAcoes(noObjeto){
  const itens = filhosAtivos(noObjeto).map(f=>f.dados);
  const primeiro = itens[0];
  const trilha = primeiro ? `${primeiro.secretaria} &rsaquo; ${primeiro.executor}` : '';
  document.getElementById('painel-conteudo').innerHTML = `
    <div class="trilha">${trilha}</div>
    <h2>${noObjeto.nome}</h2>
    <div class="lista-rotulo">${itens.length} ${itens.length===1?'ação':'ações'} neste objeto &middot; clique para abrir a ficha</div>
    <div class="lista-acoes">${itens.map(linhaAcaoHtml).join('')}</div>
  `;
  document.querySelectorAll('.lista-acao-item').forEach(el=>{
    el.onclick = ()=>{
      const d = itens.find(x=>String(x.id)===el.dataset.id);
      if(d) abrirPainel(d, noObjeto);
    };
  });
  painel.classList.add('aberto');
}

document.getElementById('painel-fechar').onclick = ()=>painel.classList.remove('aberto');

// fecha o painel clicando em qualquer lugar fora dele -- escuta na fase de
// captura (antes da propagação normal) porque os cards do mapa chamam
// stopPropagation() no próprio clique, então um listener comum no
// document nunca seria alcançado por eles.
document.addEventListener('click', e=>{
  if(!painel.classList.contains('aberto')) return;
  if(arrastouBastante) return; // era arrasto do mapa, não um clique
  if(e.target.closest('#painel')) return; // clique dentro do próprio painel
  painel.classList.remove('aberto');
}, true);

// ---------- dashboard (chips de fase) ----------
function filtrosAtivos(){
  return filtroFases.size>0 || orgaosDesligados.size>0 || !!termoBusca;
}

function limparFiltros(){
  filtroFases.clear();
  orgaosDesligados.clear();
  termoBusca = '';
  const busca = document.getElementById('busca');
  if(busca) busca.value = '';
  document.querySelectorAll('[data-org]').forEach(cb=>{ cb.checked = true; });

  // volta a árvore (aberto/fechado) e a posição/zoom para o mesmo estado
  // de quando o programa abriu, não só os filtros
  colapsado.clear();
  definirColapsoPadrao(DADOS, 'r');
  zoom = 1; panX = 60; panY = 40;

  montarDashboard();
  desenhar();
}

function montarDashboard(){
  const cont = {};
  FASE_ORDEM.forEach(f=>cont[f]=0);
  let outras = 0;
  todosItens.forEach(d=>{
    if(cont.hasOwnProperty(d.fase)) cont[d.fase]++; else outras++;
  });
  const el = document.getElementById('dashboard');
  let chips = FASE_ORDEM.map(f=>{
    const cor = corFase(f);
    const ativo = filtroFases.has(f);
    return `<div class="chip" data-fase="${f}" style="${ativo?`background:${cor};border-color:${cor}`:''}">
      <span class="dot" style="background:${cor}"></span>${f} <span class="n">${cont[f]}</span></div>`;
  }).join('');
  if(outras>0){
    const ativo = filtroFases.has('__outras__');
    chips += `<div class="chip" data-fase="__outras__" style="${ativo?'background:var(--cor-texto-terciario);border-color:var(--cor-texto-terciario)':''}">
      <span class="dot" style="background:var(--cor-texto-terciario)"></span>Outras <span class="n">${outras}</span></div>`;
  }
  const temFiltro = filtrosAtivos();
  el.innerHTML = chips + `<button id="btn-limpar-filtros" class="btn-limpar"${temFiltro?'':' disabled'}>Limpar filtros</button>`;
  el.querySelectorAll('.chip').forEach(c=>{
    c.onclick = ()=>{
      const f = c.dataset.fase;
      if(filtroFases.has(f)) filtroFases.delete(f); else filtroFases.add(f);
      montarDashboard(); desenhar();
    };
  });
  document.getElementById('btn-limpar-filtros').onclick = limparFiltros;
}

// ---------- radar de prazos (próximos vencimentos da previsão atual) ----------
function montarRadar(){
  const relevantes = todosItens.filter(d=>{
    const sit = situacaoItem(d);
    if(sit === 'Atrasada') return true;
    if(sit === 'Em dia'){ const ms = diasRestantes(d); return ms !== null && ms <= 30*86400000; }
    return false;
  });
  relevantes.sort((a,b)=> (diasRestantes(a)||0) - (diasRestantes(b)||0) );
  const top = relevantes.slice(0,6);
  document.getElementById('radar-cards').innerHTML = top.map(d=>{
    const cor = corContagem(d);
    return `<div class="radar-card" style="border-left-color:${cor}" data-id="${d.id}">
      <div class="rp">${d.objeto}</div>
      <div class="rd">${d.secretaria}</div>
      <div class="rc" style="color:${cor}" data-contagem-radar="${d.id}">${formatarContagem(d)}</div>
    </div>`;
  }).join('') || '<div style="font-size:12px;color:var(--cor-texto-terciario)">Nenhum prazo urgente no momento.</div>';
  document.querySelectorAll('.radar-card').forEach(c=>{
    c.onclick = ()=>{
      const d = todosItens.find(x=>String(x.id)===c.dataset.id);
      if(d) abrirPainel(d);
    };
  });
}

// ---------- filtro por secretaria/órgão ----------
function montarFiltroOrgao(){
  const orgaos = [...new Set(todosItens.map(d=>d.secretaria))];
  document.getElementById('filtro-org').innerHTML = orgaos.map(o=>`
    <label><input type="checkbox" checked data-org="${o}"> ${o}</label>`).join('');
  document.querySelectorAll('[data-org]').forEach(cb=>{
    cb.onchange = ()=>{
      const o = cb.dataset.org;
      if(cb.checked) orgaosDesligados.delete(o); else orgaosDesligados.add(o);
      montarDashboard(); desenhar();
    };
  });
}

function montarLegenda(){
  const linhasSituacao = [
    ['Atrasada','previsão de conclusão atual já passou'],
    ['Em dia','dentro da previsão atual'],
    ['Sem previsão','Previsão de Conclusão Atual não definida'],
    ['Concluída','fase já concluída'],
  ].map(([s,desc])=>`<div><span style="background:${s==='Atrasada'?corResolvida('var(--atrasada)'):s==='Concluída'?corResolvida('var(--fase-azul)'):corResolvida('var(--cor-texto-terciario)')}"></span><b>${s}</b>: ${desc}</div>`).join('');
  document.getElementById('legenda-item').innerHTML =
    '<div style="font-weight:700;color:var(--cor-texto-primario);margin-bottom:2px;">Situação do prazo</div>' + linhasSituacao;
}

// ---------- recolher/expandir o painel lateral ----------
document.getElementById('lateral-toggle').onclick = ()=>{
  const corpoEl = document.getElementById('corpo');
  const recolhido = corpoEl.classList.toggle('lateral-recolhido');
  const btn = document.getElementById('lateral-toggle');
  btn.innerHTML = recolhido ? '&rsaquo;' : '&lsaquo;';
  btn.title = recolhido ? 'Expandir painel' : 'Recolher painel';
  btn.setAttribute('aria-label', btn.title);
};

document.getElementById('busca').addEventListener('input', e=>{
  termoBusca = e.target.value.trim().toLowerCase();
  montarDashboard();
  desenhar();
});

document.getElementById('btn-reset').onclick = ()=>{
  // "Expandir tudo" para no nível do OBJETO, não da ação: o nó "objeto"
  // nunca abre dentro da árvore (as ações vivem no painel lateral), então
  // a expansão total revela secretaria/órgão > executor > objeto e mantém
  // todo objeto recolhido. Alterna com "recolher tudo".
  const expansiveis = [];
  (function rec(no){
    if(no.tipo!=='raiz' && no.tipo!=='objeto' && no.tipo!=='item' && no.filhos.length>0) expansiveis.push(no._id);
    no.filhos.forEach(rec);
  })(DADOS);
  const algumFechado = expansiveis.some(id=>colapsado.has(id));
  colapsado.clear();
  // objeto sempre recolhido — a expansão total para nele
  (function rec(no){ if(no.tipo==='objeto' && no.filhos.length>0) colapsado.add(no._id); no.filhos.forEach(rec); })(DADOS);
  if(!algumFechado){
    // já estava tudo aberto até o objeto -> agora recolhe tudo (menos a raiz)
    (function rec(no){ if(no.tipo!=='raiz' && no.filhos.length>0) colapsado.add(no._id); no.filhos.forEach(rec); })(DADOS);
  }
  desenhar();
  ajustarParaCaber();
};
function ajustarParaCaber(){
  const [w,h] = larguraAltura();
  const vw = viewport.clientWidth, vh = viewport.clientHeight;
  const margem = 40;
  const escala = Math.min((vw-margem*2)/w, (vh-margem*2)/h);
  zoom = Math.min(2.2, Math.max(0.25, escala));
  panX = Math.max(margem, (vw - w*zoom)/2);
  panY = Math.max(margem, (vh - h*zoom)/2);
  aplicarTransform();
}

function atualizarRelogio(){
  const agora = new Date();
  document.getElementById('relogio-data').textContent = agora.toLocaleDateString('pt-BR');
}

function atualizarContagensAoVivo(){
  document.querySelectorAll('[data-contagem]').forEach(el=>{
    const id = el.dataset.contagem;
    const d = todosItens.find(x=>String(x.id)===id);
    if(d){ el.textContent = formatarContagem(d); el.style.color = corContagem(d); }
  });
  document.querySelectorAll('[data-contagem-radar]').forEach(el=>{
    const id = el.dataset.contagemRadar;
    const d = todosItens.find(x=>String(x.id)===id);
    if(d) el.textContent = formatarContagem(d);
  });
}

montarDashboard();
montarRadar();
montarFiltroOrgao();
montarLegenda();
desenhar();
atualizarRelogio();
// só a data aparece (sem hora/min/seg), então não precisa atualizar a cada
// segundo — 1x por minuto já garante trocar no instante em que vira o dia
setInterval(atualizarRelogio, 60000);
setInterval(atualizarContagensAoVivo, 60000);
setInterval(montarRadar, 60000);
"""
