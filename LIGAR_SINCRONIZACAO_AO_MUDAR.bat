@echo off
REM Liga a sincronizacao automatica com o GITHUB COMO REFERENCIA (16/11/2026):
REM   - planilha PAC da pasta "PANORAMA\2026" -> GitHub (commit so dela)
REM   - GitHub -> esta pasta (a cada 5 min; gera o .exe se veio codigo novo)
REM   - GitHub (+ .exe) -> OneDrive e COAHS (espelho)
REM Detalhes no cabecalho de VIGIAR_GITHUB.ps1. Roda escondido e comeca
REM sozinho a cada login do Windows (chave HKCU...\Run do Registro: nesta
REM maquina a pasta Inicializar esta corrompida e a conta nao pode criar
REM Tarefas Agendadas; .vbs tambem nao roda -- por isso o PowerShell).

echo Parando vigias antigos, se estiverem rodando...
wmic process where "name='cmd.exe' and CommandLine like '%%VIGIAR_SINCRONIZACAO.bat%%'" call terminate >nul 2>&1
wmic process where "name='powershell.exe' and CommandLine like '%%VIGIAR_GITHUB.ps1%%'" call terminate >nul 2>&1
schtasks /Delete /TN "Sync Panorama Backups" /F >nul 2>&1

echo Ligando o inicio automatico (a cada login do Windows)...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Sincronizar PANORAMA - venv" /t REG_SZ /d "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"%~dp0lancar_vigia_oculto.ps1\"" /f >nul

echo Iniciando agora mesmo...
powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0lancar_vigia_oculto.ps1"

echo.
echo Pronto! Sincronizacao ligada (GitHub como referencia).
echo Acompanhe pelo arquivo _vigia_sincronizacao.log nesta pasta.
echo Para desligar, rode DESLIGAR_SINCRONIZACAO_AO_MUDAR.bat
pause
