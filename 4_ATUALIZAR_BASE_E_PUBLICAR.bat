@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title CGAPE - Atualizar planilha/codigo e publicar
color 0B

REM ---------------------------------------------------------------
REM Use este .bat sempre que:
REM   - salvar uma versao nova da planilha
REM     (PANORAMA - PAC ORIGINAL - PAC SELECOES - 2026.xlsx), e/ou
REM   - mudar algo no "CGAPE - BALANCO PAC.py" (ou em qualquer outro
REM     arquivo do projeto)
REM e quiser levar essa mudanca pro GitHub - o que faz o link do
REM painel web (Render) se atualizar sozinho em alguns minutos.
REM
REM NAO mexe no .exe: quem gera/atualiza o executavel continua sendo
REM o 3_EMPACOTAR_exe.bat, rodado a parte quando voce quiser.
REM
REM Este script MOSTRA o que vai ser publicado e PEDE confirmacao
REM antes de mandar qualquer coisa pro repositorio publico.
REM ---------------------------------------------------------------

set "PASTA=%~dp0"
cd /d "%PASTA%"

where git >nul 2>nul
if errorlevel 1 (
    color 0C
    echo   ERRO: git nao encontrado no PATH.
    echo   Instale o Git for Windows: https://git-scm.com/download/win
    pause
    exit /b 1
)

if not exist "%PASTA%.git" (
    color 0C
    echo   ERRO: esta pasta nao e um repositorio git.
    echo   Rode este .bat sempre a partir da pasta do projeto.
    pause
    exit /b 1
)

echo Verificando o que mudou desde a ultima publicacao...
echo.
git status --short
echo.

for /f %%C in ('git status --porcelain ^| find /c /v ""') do set QTD_MUDANCAS=%%C
if "%QTD_MUDANCAS%"=="0" (
    color 0E
    echo   Nada mudou desde a ultima publicacao - nao ha nada pra enviar.
    pause
    exit /b 0
)

echo ==========================================================
echo   As linhas acima sao o que vai ser publicado no repositorio
echo   PUBLICO do GitHub (drochafig-sketch/cgape-balanco-pac).
echo   O link do painel web vai se atualizar sozinho, em alguns
echo   minutos, depois disso.
echo ==========================================================
echo.
choice /c SN /n /m "Confirma o envio? (S=sim, N=nao): "
if errorlevel 2 (
    echo Cancelado - nada foi enviado.
    pause
    exit /b 0
)

REM Usa expansao atrasada (!VAR! em vez de %VAR%) daqui pra frente porque o
REM texto digitado pode ter parenteses, & ou outros caracteres especiais -
REM com %MENSAGEM% expandido "cedo demais" isso confunde o interpretador do
REM cmd.exe (foi exatamente o erro "sintaxe do comando incorreta" visto ao
REM testar com uma mensagem contendo parenteses).
set "MENSAGEM="
set /p MENSAGEM="Descreva rapidamente o que mudou (ou deixe em branco): "
if "!MENSAGEM!"=="" set "MENSAGEM=Atualiza planilha/codigo"

echo.
echo Enviando para o GitHub...
git add -A
git commit -m "!MENSAGEM!"
if errorlevel 1 (
    color 0C
    echo.
    echo   Nada foi commitado - confira a mensagem acima.
    pause
    exit /b 1
)

git push origin master
if errorlevel 1 (
    color 0C
    echo.
    echo   Falha ao enviar pro GitHub. Motivos comuns:
    echo     - sem internet
    echo     - alguem mais publicou antes de voce ^(rode "git pull" e tente de novo^)
    echo   O commit ja foi feito localmente - nada foi perdido, so nao
    echo   chegou no GitHub ainda.
    pause
    exit /b 1
)

echo.
color 0E
echo ==========================================================
echo   Publicado com sucesso!
echo.
echo   O Render vai reconstruir e publicar sozinho em alguns
echo   minutos. Acompanhe em:
echo     https://dashboard.render.com
echo   Link do painel (mesmo de sempre, so atualiza o conteudo):
echo     https://pac-balanco-painel.onrender.com
echo ==========================================================
pause
