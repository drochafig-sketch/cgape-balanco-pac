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
set "VENV=%PASTA%.venv\Scripts"
set "SCRIPT=%PASTA%CGAPE - BALANÇO PAC.py"

if not exist "%VENV%\python.exe" (
    color 0C
    echo   ERRO: .venv nao encontrado. Rode 2_INSTALAR_maquina_nova.bat antes.
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
