@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title CGAPE - Instalar ambiente (MAQUINA NOVA)
color 0B

echo.
echo ==========================================================
echo   CGAPE / BALANCO PAC - INSTALAR AMBIENTE
echo   Rode este arquivo na MAQUINA NOVA
echo ==========================================================
echo.

set "PASTA=%~dp0"
set "REQ=%PASTA%_EXPORTADO\requirements-travado.txt"
set "VENV=%PASTA%.venv"

if not exist "%REQ%" (
    echo   Nao achei _EXPORTADO\requirements-travado.txt
    echo   Vou usar requirements-minimo.txt como reserva.
    set "REQ=%PASTA%requirements-minimo.txt"
    if not exist "!REQ!" (
        color 0C
        echo   ERRO: nenhum arquivo de requisitos encontrado.
        pause
        exit /b 1
    )
)
echo   Arquivo de requisitos: %REQ%
echo.

echo [1/5] Verificando o Python...
where python >nul 2>&1
if errorlevel 1 (
    color 0C
    echo   ERRO: "python" nao esta no PATH.
    echo   Instale a MESMA versao anotada em _EXPORTADO\AMBIENTE_PYTHON.txt
    echo   e marque "Add python.exe to PATH" na instalacao.
    pause
    exit /b 1
)
python -c "import platform,struct;print('Encontrado: Python',platform.python_version(),struct.calcsize('P')*8,'bits')"
if exist "%PASTA%_EXPORTADO\AMBIENTE_PYTHON.txt" (
    echo.
    echo   Na maquina antiga era:
    type "%PASTA%_EXPORTADO\AMBIENTE_PYTHON.txt"
    echo.
    echo   Se a versao MAIOR.MENOR for diferente ^(ex: 3.11 vs 3.12^),
    echo   pare agora e instale a versao correta. Feche esta janela.
    echo.
    pause
)

echo.
echo [2/5] Criando ambiente virtual em .venv ...
if exist "%VENV%" (
    echo   Ja existe uma pasta .venv. Apagando para comecar limpo...
    rmdir /s /q "%VENV%"
)
python -m venv "%VENV%"
if errorlevel 1 (
    color 0C
    echo   ERRO ao criar o ambiente virtual.
    pause
    exit /b 1
)

echo.
echo [3/5] Atualizando pip dentro do .venv ...
call "%VENV%\Scripts\python.exe" -m pip install --upgrade pip setuptools wheel

echo.
echo [4/5] Instalando as bibliotecas travadas ...
call "%VENV%\Scripts\python.exe" -m pip install -r "%REQ%"
if errorlevel 1 (
    color 0E
    echo.
    echo   Alguma biblioteca falhou. Causas mais comuns:
    echo     - pacote que so existia local ^(caminho de arquivo no requirements^)
    echo     - versao antiga sem wheel para a versao nova do Python
    echo   Abra o requirements-travado.txt, comente a linha problematica
    echo   com # e rode este .bat de novo.
    echo.
    pause
)

echo.
echo [5/5] Rodando o diagnostico ...
call "%VENV%\Scripts\python.exe" "%PASTA%verificar_ambiente.py"

echo.
color 0A
echo ==========================================================
echo   Ambiente pronto.
echo   No VS Code: Ctrl+Shift+P, "Python: Select Interpreter",
echo   escolha .venv\Scripts\python.exe
echo ==========================================================
pause
