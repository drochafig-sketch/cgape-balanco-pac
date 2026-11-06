@echo off
chcp 65001 >nul
title Painel PAC - Compartilhado na rede local
color 0B
cd /d "%~dp0"

REM ---------------------------------------------------------------
REM Liga o painel para quem esta na mesma rede abrir pelo navegador
REM (servidor_web.py). Precisa continuar aberto enquanto houver gente
REM usando. A planilha e lida ao ligar: depois de salvar uma versao
REM nova, feche esta janela e abra de novo.
REM ---------------------------------------------------------------
set PAC_WEB_MODE=1
set PORT=5000

echo.
echo   Carregando a planilha, aguarde...
echo.
echo   Quando aparecer "Running on", o link para compartilhar e:
echo.
echo       http://%COMPUTERNAME%:%PORT%
echo.
echo   Mantenha esta janela aberta. Fechar a janela desliga o painel.
echo.

"%~dp0Scripts\python.exe" "%~dp0servidor_web.py"

echo.
echo   O painel foi desligado.
pause
