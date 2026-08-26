@echo off
setlocal

set "PASTA_VENV=%~dp0"
if "%PASTA_VENV:~-1%"=="\" set "PASTA_VENV=%PASTA_VENV:~0,-1%"
for %%I in ("%PASTA_VENV%\..\..") do set "PASTA_2026=%%~fI"

REM Curinga em vez do nome exato (com acento) do arquivo: evita colocar
REM caracteres acentuados neste .bat (cmd.exe, mesmo com chcp 65001,
REM corrompe a leitura do proprio script quando ha acento no arquivo --
REM mesmo problema que os outros .bat deste projeto ja evitam). So existe
REM um arquivo "PANORAMA - PAC ORIGINAL...xlsx" na pasta "2026".
set "PADRAO_ARQ=PANORAMA - PAC ORIGINAL*.xlsx"
set "LOG=%PASTA_VENV%\_vigia_sincronizacao.log"

REM ---------------------------------------------------------------
REM Roda para sempre (Ctrl+C para parar, se executado manualmente com
REM janela aberta), verificando a cada poucos segundos:
REM
REM   1) Se a planilha PANORAMA (pasta "PANORAMA\2026", editada a mao)
REM      ficou mais nova que a copia de dentro deste venv -- se sim,
REM      copia por cima da copia daqui, substituindo a anterior.
REM   2) Chama SINCRONIZAR_BACKUPS.bat, que espelha esta pasta (venv)
REM      inteira para o OneDrive e para o COAM (robocopy /MIR -- so
REM      copia o que realmente mudou, entao rodar isso a cada poucos
REM      segundos sem nada novo e praticamente instantaneo).
REM
REM Ligado automaticamente a cada login do Windows por
REM LIGAR_SINCRONIZACAO_AO_MUDAR.bat (roda este arquivo escondido, via
REM sincronizar_ao_mudar_oculto.vbs). Substitui a Tarefa Agendada antiga
REM "Sync Panorama Backups" (rodava so a cada 15 min) -- o gatilho agora
REM e "mudou, sincroniza" em vez de horario fixo.
REM ---------------------------------------------------------------

echo [%date% %time%] Vigia de sincronizacao iniciado (planilha PANORAMA + backups OneDrive/COAM).>> "%LOG%"

:loop
robocopy "%PASTA_2026%" "%PASTA_VENV%" "%PADRAO_ARQ%" /NFL /NDL /NJH /NJS /NP /R:2 /W:5 >nul
set "RC=%errorlevel%"
if %RC% GEQ 8 (
    echo [%date% %time%] ERRO ao copiar a planilha PANORAMA para o venv ^(codigo %RC%^).>> "%LOG%"
) else if %RC% GEQ 1 (
    echo [%date% %time%] Planilha PANORAMA atualizada: copiada da pasta "2026" para o venv, substituindo a anterior.>> "%LOG%"
)

call "%PASTA_VENV%\SINCRONIZAR_BACKUPS.bat"

ping -n 6 127.0.0.1 >nul
goto :loop
