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
call "%VENV%\python.exe" -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --clean ^
    --noconfirm ^
    --name "CGAPE - BALANCO PAC" ^
    --collect-all webview ^
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

echo.
color 0E
echo ==========================================================
echo   .exe gerado em: %PASTA%dist
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
