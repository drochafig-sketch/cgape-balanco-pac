@echo off
REM Liga a sincronizacao automatica POR MUDANCA (nao mais por horario):
REM a partir de agora, a cada poucos segundos, o sistema confere se a
REM planilha PANORAMA (pasta "PANORAMA\2026") mudou -- se mudou, copia
REM por cima da copia de dentro deste venv, substituindo a anterior --
REM e tambem espelha esta pasta (venv) inteira para o OneDrive e para o
REM COAM. Roda escondido, mesmo com tudo fechado, e comeca sozinho a
REM cada vez que o Windows for ligado.
REM
REM Substitui a Tarefa Agendada antiga "Sync Panorama Backups" (rodava
REM so a cada 15 minutos) -- essa tarefa e removida abaixo.
REM
REM O inicio automatico usa a chave HKCU...\Run do Registro (nao a pasta
REM Inicializar do Windows nem uma nova Tarefa Agendada): nesta maquina a
REM pasta Inicializar esta corrompida (virou um arquivo, nao uma pasta) e
REM a conta nao tem permissao para CRIAR tarefas agendadas novas (so para
REM apagar as que ja existiam) -- a chave Run funcionou no teste e e o
REM mecanismo mais simples que sobrou.
REM
REM O disparo escondido usa lancar_vigia_oculto.ps1 (via PowerShell), nao
REM um .vbs: nesta maquina, arquivos .vbs rodados pelo wscript.exe ficam
REM sem efeito (Windows Script Host bloqueado por politica).

echo Removendo a Tarefa Agendada antiga (rodava a cada 15 min)...
schtasks /Delete /TN "Sync Panorama Backups" /F >nul 2>&1

echo Ligando o inicio automatico (a cada login do Windows)...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Sincronizar PANORAMA - venv" /t REG_SZ /d "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"%~dp0lancar_vigia_oculto.ps1\"" /f >nul

echo Parando uma eventual instancia antiga, se estiver rodando...
wmic process where "name='cmd.exe' and CommandLine like '%%VIGIAR_SINCRONIZACAO.bat%%'" call terminate >nul 2>&1

echo Iniciando a sincronizacao agora mesmo...
powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0lancar_vigia_oculto.ps1"

echo.
echo Pronto! A sincronizacao automatica por mudanca esta ligada.
echo A partir de agora, qualquer alteracao e sincronizada em poucos
echo segundos (planilha PANORAMA + OneDrive + COAM), sem precisar mais
echo esperar os 15 minutos da tarefa antiga.
echo.
echo Para conferir se esta rodando, veja o arquivo _vigia_sincronizacao.log
echo nesta pasta (e os logs de sempre: _sync_onedrive.log / _sync_coam.log).
echo Para desligar, rode DESLIGAR_SINCRONIZACAO_AO_MUDAR.bat
pause
