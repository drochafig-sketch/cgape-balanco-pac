# Migração do CGAPE / BALANÇO PAC para uma máquina nova

A regra que resolve 90% do problema: a máquina antiga é a fonte da verdade. Nada de reinstalar as bibliotecas "na versão mais nova" e torcer. Você congela o que existe hoje, leva o arquivo congelado, e a máquina nova reproduz exatamente aquilo dentro de um ambiente virtual isolado.

Este kit tem cinco arquivos de apoio:

| Arquivo | Onde roda | Para quê |
|---|---|---|
| `1_EXPORTAR_maquina_antiga.bat` | máquina antiga | congela versões e faz inventário dos arquivos |
| `2_INSTALAR_maquina_nova.bat` | máquina nova | cria o `.venv` e instala tudo igual |
| `verificar_ambiente.py` | máquina nova | diagnóstico completo antes de rodar o relatório |
| `3_EMPACOTAR_exe.bat` | máquina nova | gera o executável com PyInstaller |
| `.vscode/` | máquina nova | interpretador, debug e encoding já configurados |

---

## Antes de começar

Coloque a pasta `MIGRACAO_CGAPE` **dentro da pasta do projeto**, junto com o `CGAPE - BALANÇO PAC.py`. Depois mova os arquivos para o nível do projeto, de forma que a estrutura final fique assim na máquina nova:

```
C:\CGAPE\
    CGAPE - BALANÇO PAC.py
    PANORAMA - PAC ORIGINAL - PAC SELEÇÕES - 2026.xlsx
    PAC.png
    GOVERNO.PNG
    municipios_bahia.geojson
    calibri.ttf
    calibrib.ttf
    DIN.ttf
    DIN-Bold.ttf
    bahnschrift.ttf
    1_EXPORTAR_maquina_antiga.bat
    2_INSTALAR_maquina_nova.bat
    3_EMPACOTAR_exe.bat
    verificar_ambiente.py
    requirements-minimo.txt
    .gitignore
    .vscode\
        settings.json
        launch.json
        extensions.json
    _EXPORTADO\          (vem da máquina antiga)
    .venv\               (criado pelo passo 4)
```

Essa lista de arquivos não é opinião minha, é o que o próprio script procura via `caminho_recurso()`. A planilha, o `PAC.png` e o `GOVERNO.PNG` são obrigatórios. As fontes `.ttf` e o `municipios_bahia.geojson` são opcionais no sentido de que o código tem reserva para eles, mas se faltarem o PDF sai com Helvetica no lugar de Calibri/DIN/Bahnschrift e a página do mapa coroplético simplesmente some. Visualmente é outro relatório. Leve todos.

Escolha um caminho curto e sem espaços exóticos para a pasta, tipo `C:\CGAPE`. Evite colocar o projeto dentro de pasta sincronizada do OneDrive: a sincronização trava arquivos no meio da geração do PDF e dá erro de permissão que parece bug do código e não é.

---

## Passo 1: congelar o ambiente na máquina antiga

Se você usa ambiente virtual hoje, ative ele primeiro no terminal. Se sempre rodou no Python global, pode ignorar isso.

Dê duplo clique em `1_EXPORTAR_maquina_antiga.bat`. Ele cria a pasta `_EXPORTADO` com:

* `AMBIENTE_PYTHON.txt` com a versão exata do Python, se é 32 ou 64 bits e onde ele está instalado
* `requirements-travado.txt`, que é o `pip freeze` completo, com todas as versões travadas em `==`
* `BIBLIOTECAS_CRITICAS.txt` com as versões de pandas, numpy, reportlab, openpyxl e pywebview
* `INVENTARIO_ARQUIVOS.txt` dizendo quais arquivos de apoio existem e o tamanho de cada um
* `HASH_SCRIPTS.txt` com a assinatura SHA256 dos `.py`, para você conferir depois que a cópia não corrompeu

O arquivo que importa de verdade é o `requirements-travado.txt`. Guarde ele em mais de um lugar. Ele é o que garante que o relatório vai sair idêntico.

## Passo 2: copiar tudo

Leve a pasta do projeto inteira, incluindo a `_EXPORTADO`, mas **sem** a pasta `.venv` antiga. Ambiente virtual não é portátil, ele guarda caminhos absolutos da máquina onde foi criado, e copiar ele para outro computador gera erros de "python não encontrado" difíceis de rastrear. O `2_INSTALAR` recria do zero.

Se a cópia for por rede ou pen drive, confira o hash do script depois:

```cmd
certutil -hashfile "CGAPE - BALANÇO PAC.py" SHA256
```

O resultado tem que bater com o que está no `HASH_SCRIPTS.txt`.

## Passo 3: instalar o Python certo na máquina nova

Abra o `_EXPORTADO\AMBIENTE_PYTHON.txt` e veja a versão. Instale a **mesma versão maior.menor** em [python.org](https://www.python.org/downloads/windows/). Se a antiga era 3.11.9, pode instalar 3.11.x qualquer. Não pule para 3.12 ou 3.13 nesse momento.

O motivo é prático: pandas, numpy e reportlab distribuem wheels compilados por versão de Python. Uma versão travada no `requirements` que existia para 3.11 pode simplesmente não ter wheel para 3.13, e aí o pip tenta compilar do zero, não acha compilador, e a instalação quebra. Migração e atualização de Python são duas mudanças diferentes. Faça uma de cada vez.

Na instalação, marque **Add python.exe to PATH** e mantenha marcado **tcl/tk and IDLE** (o script usa tkinter). Instale a versão 64 bits.

## Passo 4: criar o ambiente

Duplo clique em `2_INSTALAR_maquina_nova.bat`. Ele confere a versão do Python, mostra lado a lado com a da máquina antiga e pede confirmação, cria o `.venv`, atualiza o pip e instala o `requirements-travado.txt` inteiro. No fim ele já roda o diagnóstico sozinho.

Se alguma linha do requirements falhar, geralmente é pacote que estava instalado por caminho local ou pacote que não é do projeto. Abra o `requirements-travado.txt`, comente a linha com `#` e rode de novo. As cinco que não podem falhar são pandas, numpy, reportlab, openpyxl e pywebview.

## Passo 5: apontar o VS Code

Instale o VS Code e a extensão Python da Microsoft. Abra a pasta do projeto com **File > Open Folder** (abrir o arquivo solto em vez da pasta é o erro mais comum, e faz o `.vscode` ser ignorado).

Aperte `Ctrl+Shift+P`, digite `Python: Select Interpreter` e escolha o que aparece como `.venv\Scripts\python.exe`. Confirme no canto inferior direito que está escrito `.venv` e não uma versão global.

O `launch.json` já vem com três configurações: rodar o relatório, rodar o arquivo aberto e rodar o diagnóstico. Ele também força `PYTHONUTF8=1`, o que evita aquele erro de `UnicodeEncodeError` quando o script imprime acento no terminal do Windows.

## Passo 6: validar antes de confiar

Rode `verificar_ambiente.py` (pelo F5, escolhendo a configuração de diagnóstico, ou pelo terminal). Ele testa em sequência a versão e arquitetura do Python, a presença e versão de cada biblioteca, a divergência contra o `requirements-travado.txt`, a presença de cada arquivo de apoio, tkinter, sockets locais, o WebView2 Runtime e, por último, gera um PDF de teste com o ReportLab registrando a Calibri.

Se o resumo final disser "nenhum problema bloqueante", rode o relatório de verdade e compare o PDF gerado com um PDF antigo, página por página. É a única validação que realmente conta.

---

## Problemas conhecidos e o que fazer

**`OSError: [WinError 10022]` no pywebview.** É o mesmo erro que você já pegou. Não tem relação com a planilha nem com o código. É o catálogo Winsock do Windows corrompido. Abra o CMD como administrador, rode `netsh winsock reset` e reinicie a máquina. Se por algum motivo não puder reiniciar, os dois contornos que já funcionaram são passar `http_port=8765` no `webview.start()`, evitando a escolha de porta aleatória, ou envolver a chamada em `try/except OSError` e cair para `webbrowser.open()` do HTML. O diagnóstico já testa isso e avisa antes de você descobrir na hora errada.

**O painel abre em branco ou não abre.** Falta o WebView2 Runtime. Baixe o "Evergreen Standalone Installer" no site da Microsoft. O Windows 11 já traz de fábrica, o Windows 10 nem sempre.

**O PDF saiu com fonte errada.** As `.ttf` não estão na pasta. O código não avisa, ele cai silenciosamente para Helvetica por design.

**A página do mapa da Bahia sumiu.** Falta o `municipios_bahia.geojson`. Mesmo comportamento tolerante do item anterior.

**`PermissionError` ao gerar o PDF.** A planilha ou o PDF anterior está aberto no Excel ou no leitor de PDF. Feche e rode de novo.

**Gráfico ou tabela saiu diferente do esperado.** Aí sim é divergência de versão de biblioteca. Compare o `BIBLIOTECAS_CRITICAS.txt` com o que o diagnóstico reporta. Mudança de versão maior do pandas costuma alterar comportamento de `groupby`, ordenação e tratamento de tipos, que é exatamente onde o relatório é sensível.

**Antivírus apagou o `.exe` recém gerado.** Executável de PyInstaller é falso positivo comum. Adicione a pasta `dist` na exclusão do antivírus.

---

## Checklist final

- [ ] `_EXPORTADO\requirements-travado.txt` gerado na máquina antiga e copiado
- [ ] Python com a mesma versão maior.menor, 64 bits, com PATH e tkinter
- [ ] Pasta do projeto fora do OneDrive, caminho curto
- [ ] Todos os arquivos de apoio presentes, incluindo as cinco `.ttf` e o geojson
- [ ] `.venv` criado pelo `2_INSTALAR`, sem cópia do `.venv` antigo
- [ ] Interpretador do VS Code apontando para o `.venv`
- [ ] `verificar_ambiente.py` sem problema bloqueante
- [ ] PDF de teste gerado e comparado com um relatório antigo
- [ ] Painel HTML abrindo normalmente
- [ ] `requirements-travado.txt` guardado em backup, fora da máquina
