# Publicar o painel PAC na web (Render)

Isto é sobre uma coisa diferente do `MIGRACAO.md`: aquele é sobre levar o
app **desktop** (.exe) para outro computador. Este aqui é sobre publicar o
mesmo painel como um **link que abre no navegador**, sem instalar nada —
via `servidor_web.py`, hospedado gratuitamente no [Render](https://render.com).

## O que roda onde

| Arquivo | Papel |
|---|---|
| `CGAPE - BALANÇO PAC.py` | Toda a lógica de negócio (planilha, filtros, PDF). Roda igual nos dois modos. |
| `servidor_web.py` | Único arquivo novo: importa o `.py` acima e expõe o painel por HTTP. |
| `requirements-web.txt` | Dependências só do servidor (Flask/gunicorn + pandas/reportlab/etc., **sem** pywebview/pyinstaller). |
| `render.yaml` | Configuração do serviço, lida automaticamente pelo Render ("Blueprint"). |

## Publicar pela primeira vez

1. Entre em [render.com](https://render.com) e crie uma conta (dá para
   usar login do GitHub direto).
2. **New +** → **Blueprint**.
3. Autorize o Render a acessar o repositório
   `drochafig-sketch/cgape-balanco-pac` (ele já está público no GitHub).
4. O Render lê o `render.yaml` sozinho e mostra o serviço
   `pac-balanco-painel` pronto — confirme e clique em **Apply**.
5. O primeiro build demora alguns minutos (instala pandas/reportlab do
   zero). Acompanhe o log; ao final ele mostra a URL pública
   (algo como `https://pac-balanco-painel.onrender.com`).
6. Abra a URL, aplique um filtro e clique em "Gerar Relatório" — se o PDF
   abrir numa nova aba, está tudo certo.

## Atualizações depois

Qualquer `git push` para o branch `master` faz o Render rebuildar e
publicar sozinho (deploy automático já vem ligado por padrão no Blueprint).
Não precisa repetir os passos acima.

## Variáveis de ambiente

`PAC_WEB_MODE` e `PYTHON_VERSION` já vêm definidas no `render.yaml`, não
precisa mexer:

- `PAC_WEB_MODE=1` — liga os pontos do código que só fazem sentido no
  servidor (não chamar `os.startfile`, não abrir diálogo do tkinter em erro
  fatal de inicialização — ver `MODO_WEB` no topo do `.py` original).
- `PYTHON_VERSION=3.12.10` — trava a mesma versão usada na máquina que gera
  o relatório hoje (ver `_EXPORTADO/AMBIENTE_PYTHON.txt`), pelo mesmo motivo
  que o `MIGRACAO.md` já explica: versão diferente de Python/biblioteca
  pode mudar o relatório visualmente.

Já `PAC_WEB_SENHA` **não** vem definida no `render.yaml` (só a chave, com
`sync: false`) — de propósito, porque este repositório é público: um valor
fixo aqui ficaria visível pra qualquer um no GitHub. O servidor nem sobe
sem ela (`RuntimeError` na inicialização). Precisa ser configurada à mão,
uma vez, direto no painel do Render:

1. No serviço `pac-balanco-painel`, vá em **Environment**.
2. **Add Environment Variable** → chave `PAC_WEB_SENHA`, valor = a senha
   escolhida.
3. **Save Changes** — o Render redeploya sozinho já com o valor novo.

É login HTTP Basic: ao abrir o painel, o navegador mostra a janelinha
nativa de usuário/senha — deixe o usuário em branco, só a senha importa.
Trocar a senha depois é só repetir o passo 2 com o novo valor (nunca via
`render.yaml`/commit, senão ela para de ser secreta).

## Avisos conhecidos

- **Fonte DIN ausente.** `DIN.ttf`/`DIN-Bold.ttf` (usada nos números de
  destaque dos cards) não está versionada neste repositório — só Calibri e
  Bahnschrift estão. Sem ela, o PDF web (e o `.exe` também, se a fonte
  faltar na pasta) cai para Calibri em negrito no lugar da DIN. Isso já
  acontecia no desktop antes desta mudança; não é uma regressão. Se quiser
  o visual idêntico ao gerado localmente, é preciso adicionar os `.ttf` ao
  repositório (avalie antes se a fonte pode ser redistribuída).
- **Plano gratuito "dorme".** Sem acesso por 15 minutos, o Render desliga a
  instância; o próximo acesso demora uns 30–50s para "acordar" o serviço,
  antes do painel abrir. Normal do plano free, não é um bug.
- **Cada requisição de PDF gera um arquivo temporário no servidor** (em
  `/tmp`, apagado logo em seguida) — normal, é assim que o servidor entrega
  os bytes do PDF pro navegador baixar/abrir.

## Testar localmente antes de publicar

```
pip install -r requirements-web.txt
```
PowerShell (o `PYTHONUTF8=1` evita o mesmo `UnicodeEncodeError` de acento
que o `.vscode/launch.json` já contorna no modo desktop — só é preciso no
Windows; no Render/Linux o padrão já é UTF-8):
```
$env:PAC_WEB_MODE = "1"
$env:PYTHONUTF8 = "1"
$env:PAC_WEB_SENHA = "algumasenha"
python servidor_web.py
```
Bash/Git Bash:
```
PAC_WEB_MODE=1 PYTHONUTF8=1 PAC_WEB_SENHA=algumasenha python servidor_web.py
```
Depois abra `http://localhost:5000` num navegador comum (login: usuário em
branco, senha = o que você pôs em `PAC_WEB_SENHA`). `http://localhost:5000/saude`
devolve um JSON simples confirmando que a planilha carregou (quantas linhas,
quando foi atualizada) — útil para checar rápido se o servidor subiu certo
sem precisar abrir o painel inteiro.
