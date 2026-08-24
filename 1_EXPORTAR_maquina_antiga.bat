@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title CGAPE - Exportar ambiente (MAQUINA ANTIGA)
color 0B

echo.
echo ==========================================================
echo   CGAPE / BALANCO PAC - EXPORTAR AMBIENTE
echo   Rode este arquivo na MAQUINA ANTIGA
echo ==========================================================
echo.
echo   Se voce usa um ambiente virtual (.venv), ATIVE ele antes
echo   de rodar este .bat, senao o pip freeze vai capturar as
echo   bibliotecas erradas.
echo.
pause

set "PASTA=%~dp0"
set "SAIDA=%PASTA%_EXPORTADO"
if not exist "%SAIDA%" mkdir "%SAIDA%"

echo.
echo [1/6] Localizando o Python...
where python >"%SAIDA%\python_localizacao.txt" 2>&1
if errorlevel 1 (
    color 0C
    echo   ERRO: comando "python" nao encontrado no PATH.
    echo   Abra o VS Code, rode: python --version
    echo   Se nao funcionar, reinstale o Python marcando "Add to PATH".
    pause
    exit /b 1
)
python -c "import sys,platform,struct;print('Python:',platform.python_version());print('Build:',struct.calcsize('P')*8,'bits');print('Executavel:',sys.executable);print('Sistema:',platform.platform())" > "%SAIDA%\AMBIENTE_PYTHON.txt"
type "%SAIDA%\AMBIENTE_PYTHON.txt"

echo.
echo [2/6] Congelando as bibliotecas instaladas...
python -m pip freeze > "%SAIDA%\requirements-travado.txt"
python -m pip list > "%SAIDA%\pip_list_legivel.txt"
python -m pip --version > "%SAIDA%\versao_pip.txt"

echo.
echo [3/6] Versoes das bibliotecas criticas do relatorio...
python -c "import importlib.metadata as m;[print(p.ljust(14), m.version(p)) for p in ['pandas','numpy','reportlab','openpyxl','pywebview'] ]" > "%SAIDA%\BIBLIOTECAS_CRITICAS.txt" 2>&1
type "%SAIDA%\BIBLIOTECAS_CRITICAS.txt"

echo.
echo [4/6] Conferindo os arquivos de apoio na pasta do projeto...
> "%SAIDA%\INVENTARIO_ARQUIVOS.txt" (
    echo INVENTARIO DE ARQUIVOS DE APOIO
    echo Pasta verificada: %PASTA%
    echo.
)
call :CONFERIR "PANORAMA - PAC ORIGINAL - PAC SELECOES - 2026.xlsx"
call :CONFERIR "PAC.png"
call :CONFERIR "GOVERNO.PNG"
call :CONFERIR "calibri.ttf"
call :CONFERIR "calibrib.ttf"
call :CONFERIR "DIN.ttf"
call :CONFERIR "DIN-Bold.ttf"
call :CONFERIR "bahnschrift.ttf"
call :CONFERIR "municipios_bahia.geojson"
type "%SAIDA%\INVENTARIO_ARQUIVOS.txt"

echo.
echo [5/6] Gerando assinatura (hash) do script principal...
for %%A in ("%PASTA%*.py") do (
    certutil -hashfile "%%~fA" SHA256 >> "%SAIDA%\HASH_SCRIPTS.txt" 2>nul
)

echo.
echo [6/6] Listando variaveis de ambiente relevantes...
> "%SAIDA%\VARIAVEIS.txt" (
    echo PATH=%PATH%
    echo.
    echo PYTHONPATH=%PYTHONPATH%
    echo PYTHONHOME=%PYTHONHOME%
    echo VIRTUAL_ENV=%VIRTUAL_ENV%
)

echo.
color 0A
echo ==========================================================
echo   CONCLUIDO
echo ==========================================================
echo   Tudo foi salvo em:
echo   %SAIDA%
echo.
echo   Copie para a maquina nova:
echo     1. A pasta _EXPORTADO inteira
echo     2. A pasta do projeto (script .py, planilha, fontes,
echo        PAC.png, GOVERNO.PNG, municipios_bahia.geojson)
echo.
pause
exit /b 0

:CONFERIR
if exist "%PASTA%%~1" (
    for %%F in ("%PASTA%%~1") do echo [OK]     %~1  ^(%%~zF bytes^) >> "%SAIDA%\INVENTARIO_ARQUIVOS.txt"
) else (
    echo [FALTA] %~1 >> "%SAIDA%\INVENTARIO_ARQUIVOS.txt"
)
exit /b 0
