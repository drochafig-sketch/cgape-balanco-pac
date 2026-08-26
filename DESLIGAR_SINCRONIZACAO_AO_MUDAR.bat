@echo off
REM Desliga a sincronizacao automatica por mudanca, ligada por
REM LIGAR_SINCRONIZACAO_AO_MUDAR.bat: para o processo que estiver
REM rodando agora e remove o inicio automatico do Windows (chave
REM HKCU...\Run do Registro).

echo Parando a sincronizacao que estiver rodando agora...
wmic process where "name='cmd.exe' and CommandLine like '%%VIGIAR_SINCRONIZACAO.bat%%'" call terminate >nul 2>&1

echo Removendo o inicio automatico...
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Sincronizar PANORAMA - venv" /f >nul 2>&1
del /Q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Sincronizar PANORAMA - venv.lnk" 2>nul

echo.
echo Pronto! A sincronizacao automatica por mudanca foi desligada.
echo (Nada e apagado do OneDrive/COAM por causa disso -- so para de
echo   sincronizar novas mudancas ate rodar LIGAR_SINCRONIZACAO_AO_MUDAR.bat
echo   de novo. Se quiser voltar a tarefa antiga de 15 em 15 minutos, sera
echo   preciso recria-la manualmente no Agendador de Tarefas do Windows.)
pause
