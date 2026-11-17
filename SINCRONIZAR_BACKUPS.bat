@echo off
REM Sincronizacao MANUAL (uma passada so), com o GitHub como referencia
REM (16/11/2026): envia a planilha da pasta "2026" ao GitHub se mudou,
REM atualiza esta pasta com o GitHub e espelha o GitHub (+ .exe) no
REM OneDrive e no COAHS. E o mesmo que o vigia automatico faz a cada ciclo
REM -- ver VIGIAR_GITHUB.ps1. Clique 2x para rodar.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0VIGIAR_GITHUB.ps1" -UmaVez
pause
