# Prompt — Identidade visual única + tela HOME (CGAPE - BALANÇO PAC)

Cole o texto abaixo no Claude Code (VS Code), com a pasta `Ambiente Virtual` aberta como workspace.

---

## Contexto

Este projeto (`CGAPE - BALANÇO PAC.py` + `gerar_pptx_pac.py`, `resumo_acoes.py`, `relatorio_monitoramento.py`, `servidor_web.py`) apresenta resultados de várias formas, cada uma com um visual próprio:

| Saída | Onde é montada | Visual atual |
|---|---|---|
| Painel desktop/web "BALANÇO PAC - BAHIA" (filtros, dash, modais) | `montar_html_painel()` em `CGAPE - BALANÇO PAC.py` | tema escuro AGF (`--cor-fundo #303030`, teal/mint/peach/gold), com tema claro alternativo |
| **Apresentação HTML do PAC (Análise Financeira)** | `montar_html_apresentacao()` + `CSS_*` em `gerar_pptx_pac.py` | **REFERÊNCIA** — navy `#0B1F3F`, azul `#0B5CAD`, teal `#2F7F92`, laranja `#E68A00`, cinzas `#5B6576 / #D8DCE3 / #F2F4F7`, marca `#4FA3E0 / #C9D6EA` |
| "Ações em Monitoramento" (janela do dash + HTML exportado) | `gerar_html_resumo_acoes()` em `resumo_acoes.py` | paleta própria (`--navy`, `--azul #1F5C85`, …) |
| Relatório PDF (ReportLab) e seção MONITORAMENTO | `_gerar_pdf()`, `desenhar_cabecalho()`, `relatorio_monitoramento.py` | cores ReportLab próprias |
| PPTX (Análise Financeira e Resumo de Ações) | `gerar_pptx_pac.py`, `gerar_pptx_resumo_acoes()` | constantes `NAVY/AZUL/TEAL/LARANJA…` |

Logos disponíveis na pasta: `PAC.png` e `GOVERNO.png` (carregadas via `caminho_recurso()`; no HTML use data URI, como `_data_uri_imagem()` já faz em `gerar_pptx_pac.py`).

## Objetivo

1. **Unificar a identidade visual** de todas as saídas tomando como referência o visual e o design do **HTML da apresentação do PAC** (`montar_html_apresentacao` em `gerar_pptx_pac.py`).
2. **Criar uma tela HOME** que é a primeira coisa exibida logo após a execução do programa, com dois grandes botões:
   - **PAC** (com a logo `PAC.png`) → abre o **Balanço PAC - Bahia** (o painel atual de `montar_html_painel`).
   - **MONITORAMENTO** → abre **Ações em Monitoramento** (a janela/visão de `resumo_acoes.py`, hoje acessada pelo botão `#dash-resumo-acoes` / `data-abrir-resumo-acoes`).

## Etapa 1 — Levantamento (não altere nada ainda)

1. Leia o CSS de `montar_html_apresentacao` e as constantes de cor no topo de `gerar_pptx_pac.py` e extraia o "design system" de referência: paleta, tipografia (família, pesos, tamanhos, caixa alta/letter-spacing), raios, bordas, sombras, cabeçalho/faixa de título, cards, botões, tabelas, chips/selos, cores semânticas (OGU/Financiamento/Estadual, fases, status, atraso).
2. Liste, para cada saída da tabela acima, o que diverge da referência.
3. **Mostre-me esse resumo (tokens + divergências) e um plano curto antes de editar.** Se algo da referência for ambíguo, pergunte.

## Etapa 2 — Fonte única dos tokens

- Crie um módulo `identidade_visual.py` com os tokens da referência em um só lugar:
  - dicionário de cores em hex (sem `#`, para python-pptx/ReportLab) + helpers `hex_css()`, `cor_reportlab()`, `cor_pptx()`;
  - string `CSS_TOKENS` com as variáveis CSS (`:root { --ivp-navy: …; }`) e estilos-base compartilhados (cabeçalho, card, botão, tabela, chip);
  - helper para data URI das logos (`PAC.png`, `GOVERNO.png`).
- Todas as saídas passam a importar desse módulo, em vez de repetir hex soltos. Mantenha nomes antigos como aliases onde for preciso para não quebrar nada (ex.: `NAVY = iv.NAVY`).
- Inclua `identidade_visual.py` no `.spec` do PyInstaller (`CGAPE - BALANCO PAC.spec`/`gerar_executavel.bat`) e confira que o `servidor_web.py` também o encontra.

## Etapa 3 — Aplicar a identidade

Aplique os tokens, nesta ordem, validando cada uma antes da próxima:

1. **Painel `montar_html_painel`**: substitua a paleta AGF pelos tokens da referência (cabeçalho, cards, botões, filtros, modais, dash). Decida comigo se o tema escuro continua existindo; se continuar, ele deve ser derivado da mesma paleta (navy como fundo), não da paleta AGF antiga. Preserve TODA a lógica JS (filtros, `chamarAPI`, modo web vs desktop, `localStorage`, busca de ficha, controle de qualidade).
2. **`resumo_acoes.py`** (`gerar_html_resumo_acoes` e a janela no dash): mesmos tokens; `--azul #1F5C85` passa a ser o azul da referência.
3. **PDF ReportLab** (`desenhar_background_capa`, `desenhar_cabecalho`, cards, gráficos) e `relatorio_monitoramento.py`: mesmas cores e mesma hierarquia de títulos.
4. **PPTX** (`gerar_pptx_pac.py`, `gerar_pptx_resumo_acoes`): apenas apontar as constantes para `identidade_visual.py`.

Regras:
- Só visual. Não mude cálculos, filtros, textos de dados, ordem de seções ou nomes de arquivos gerados.
- Cores semânticas (fontes de recurso, fases, status, atraso) devem ser idênticas em todas as saídas.
- Fonte mínima legível de 11px no HTML (padrão já usado na referência).
- Mantenha os comentários no estilo do projeto (português, explicando o porquê).

## Etapa 4 — Tela HOME

- Crie a HOME como primeira tela, nos tokens da referência:
  - cabeçalho no padrão da referência (faixa navy, título "CGAPE", data de atualização `__ULTIMA_ATUALIZACAO__`, `GOVERNO.png` à direita);
  - dois cartões-botão grandes, lado a lado (empilhados em telas estreitas), com hover/foco visíveis e acessíveis por teclado (Tab/Enter):
    - **PAC** — logo `PAC.png` em destaque + rótulo "BALANÇO PAC - BAHIA";
    - **MONITORAMENTO** — ícone SVG simples (ex.: alvo/lista de verificação) + rótulo "AÇÕES EM MONITORAMENTO".
- Navegação:
  - **Desktop (pywebview)**: em `abrir_interface_filtros`, a janela abre na HOME. Clique em PAC carrega o painel atual na mesma janela; clique em MONITORAMENTO abre a visão de Ações em Monitoramento. Implemente como uma página única com troca de visão (ou `window.load_url` / `evaluate_js`) — escolha a opção mais simples que não duplique a `js_api` (`APIFiltros`) e não carregue o painel pesado antes de ser pedido.
  - **Web (`servidor_web.py`)**: rota `/` passa a servir a HOME; crie `/pac` (painel atual) e `/monitoramento` (Ações em Monitoramento). Confira que `chamarAPI` continua funcionando nas novas rotas.
  - Em PAC e em MONITORAMENTO, adicione um botão discreto "← Início" no cabeçalho que volta para a HOME.
- O botão "AÇÕES EM MONITORAMENTO" existente dentro do dash continua funcionando.
- Pré-carregamento: a base do MONITORAMENTO já é lida em segundo plano no startup — mantenha isso, para os dois botões abrirem rápido.

## Etapa 5 — Verificação

1. Rode `python verificar_ambiente.py` e depois `python "CGAPE - BALANÇO PAC.py"`: a HOME aparece primeiro, os dois botões abrem a tela certa e "← Início" volta.
2. Gere: PDF do relatório, HTML e PPTX da Análise Financeira, HTML/PPTX/XLSX do Resumo de Ações — e confira que todos estão na mesma identidade.
3. Suba o `servidor_web.py` localmente e teste `/`, `/pac`, `/monitoramento`.
4. Me entregue um resumo do que mudou por arquivo e prints (ou descrição) das telas antes/depois. Não faça commit nem publique (`4_ATUALIZAR_BASE_E_PUBLICAR.bat`/git) sem eu pedir.
