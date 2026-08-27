@echo off
chcp 65001 >nul
setlocal
title CGAPE - Empacotar .exe (PyInstaller)
color 0B

REM ---------------------------------------------------------------
REM MODELO. Compare com o comando que voce ja usava na maquina
REM antiga antes de rodar. Se o seu comando atual for diferente,
REM ele manda: substitua a linha do pyinstaller aqui embaixo.
REM
REM Observacao importante: o script usa
REM     PASTA_BASE = dirname(sys.executable)
REM ou seja, planilha, PAC.png, GOVERNO.PNG, fontes .ttf e
REM municipios_bahia.geojson ficam AO LADO do .exe, nao embutidos.
REM Por isso nao ha --add-data aqui.
REM ---------------------------------------------------------------

set "PASTA=%~dp0"
set "SCRIPT=%PASTA%CGAPE - BALANÇO PAC.py"

REM O ambiente virtual pode estar em dois formatos, dependendo de como esta
REM pasta foi criada:
REM   - maquina migrada via MIGRACAO.md (2_INSTALAR_maquina_nova.bat): o
REM     .venv fica numa SUBPASTA ".venv" dentro do projeto.
REM   - maquina onde o proprio "python -m venv" foi rodado apontando pra
REM     esta pasta: ELA MESMA e o ambiente virtual (tem um pyvenv.cfg na
REM     raiz, Scripts\ direto aqui, sem subpasta ".venv").
REM Confere os dois, nessa ordem.
set "VENV=%PASTA%.venv\Scripts"
if not exist "%VENV%\python.exe" (
    if exist "%PASTA%Scripts\python.exe" set "VENV=%PASTA%Scripts"
)

if not exist "%VENV%\python.exe" (
    color 0C
    echo   ERRO: nao encontrei o Python do ambiente virtual. Procurei em:
    echo     %PASTA%.venv\Scripts\python.exe
    echo     %PASTA%Scripts\python.exe
    echo   Rode 2_INSTALAR_maquina_nova.bat antes, ou confira se esta pasta
    echo   e mesmo um ambiente virtual (deve ter um arquivo pyvenv.cfg).
    pause
    exit /b 1
)
if not exist "%SCRIPT%" (
    color 0C
    echo   ERRO: script principal nao encontrado em:
    echo   %SCRIPT%
    echo   Ajuste a variavel SCRIPT dentro deste .bat.
    pause
    exit /b 1
)

echo Instalando/conferindo o PyInstaller no .venv...
call "%VENV%\python.exe" -m pip install --upgrade pyinstaller

echo.
echo Empacotando...
REM --collect-submodules reportlab.graphics.barcode: o reportlab importa o
REM modulo de cada tipo de codigo de barras (code128 etc.) de forma
REM DINAMICA em tempo de execucao, entao o PyInstaller nao enxerga essa
REM dependencia sozinho na analise estatica — sem essa linha, o .exe abre
REM e fecha na hora com "ModuleNotFoundError: No module named
REM 'reportlab.graphics.barcode.code128'" assim que tenta desenhar o QR
REM Code da Ficha Cadastral.
REM --collect-all clr_loader / pythonnet: dependencias do backend
REM WebView2 do pywebview (--hidden-import clr, logo abaixo, precisa
REM dessas duas pra funcionar de verdade, nao so pra importar sem erro).
call "%VENV%\python.exe" -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --clean ^
    --noconfirm ^
    --name "CGAPE - BALANCO PAC" ^
    --collect-all webview ^
    --collect-all clr_loader ^
    --collect-all pythonnet ^
    --collect-submodules reportlab.graphics.barcode ^
    --hidden-import "clr" ^
    --distpath "%PASTA%dist" ^
    --workpath "%PASTA%build" ^
    --specpath "%PASTA%" ^
    "%SCRIPT%"

if errorlevel 1 (
    color 0C
    echo.
    echo   Falha no empacotamento. Leia a ultima mensagem de erro acima.
    pause
    exit /b 1
)

REM ---------------------------------------------------------------
REM Copia o .exe recem-gerado de dist\ para a RAIZ do projeto,
REM sobrescrevendo o antigo. E dessa raiz que:
REM   - o .exe roda direto (planilha, PNGs, fontes e geojson ja
REM     estao aqui do lado), e
REM   - a Tarefa Agendada "Sync Panorama Backups" (a cada 15 min,
REM     via SINCRONIZAR_BACKUPS.bat) espelha pro OneDrive e pra
REM     rede COAM. O dist\ NAO e sincronizado.
REM ---------------------------------------------------------------
echo.
echo Copiando o .exe novo para a raiz do projeto...
copy /Y "%PASTA%dist\CGAPE - BALANCO PAC.exe" "%PASTA%CGAPE - BALANCO PAC.exe" >nul
if errorlevel 1 (
    color 0C
    echo.
    echo   AVISO: o .exe foi gerado em dist\ mas nao consegui copiar
    echo   para a raiz do projeto ^(%PASTA%^).
    echo   Copie na mao: dist\CGAPE - BALANCO PAC.exe  ->  raiz do projeto.
    pause
    exit /b 1
)

echo.
color 0E
echo ==========================================================
echo   .exe gerado em: %PASTA%dist
echo   e copiado para: %PASTA%CGAPE - BALANCO PAC.exe
echo   ^(a sincronizacao automatica leva essa copia da raiz pro
echo    OneDrive e pra rede COAM em ate 15 min^)
echo.
echo   COPIE PARA A MESMA PASTA DO .exe:
echo     PANORAMA - PAC ORIGINAL - PAC SELECOES - 2026.xlsx
echo     PAC.png
echo     GOVERNO.PNG
echo     municipios_bahia.geojson
echo     calibri.ttf  calibrib.ttf
echo     DIN.ttf  DIN-Bold.ttf  bahnschrift.ttf
echo ==========================================================
pause
