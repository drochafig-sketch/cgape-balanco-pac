@echo off
REM Desliga a sincronizacao automatica ligada por
REM LIGAR_SINCRONIZACAO_AO_MUDAR.bat: para o vigia que estiver rodando e
REM remove o inicio automatico do Windows (chave HKCU...\Run).

echo Parando a sincronizacao que estiver rodando agora...
wmic process where "name='powershell.exe' and CommandLine like '%%VIGIAR_GITHUB.ps1%%'" call terminate >nul 2>&1
wmic process where "name='cmd.exe' and CommandLine like '%%VIGIAR_SINCRONIZACAO.bat%%'" call terminate >nul 2>&1

echo Removendo o inicio automatico...
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Sincronizar PANORAMA - venv" /f >nul 2>&1
del /Q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Sincronizar PANORAMA - venv.lnk" 2>nul

echo.
echo Pronto! A sincronizacao automatica foi desligada.
echo (Nada e apagado do GitHub, do OneDrive ou do COAHS -- so para de
echo   sincronizar ate rodar LIGAR_SINCRONIZACAO_AO_MUDAR.bat de novo.)
pause
