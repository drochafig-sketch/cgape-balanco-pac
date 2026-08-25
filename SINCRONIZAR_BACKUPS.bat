@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "ORIGEM=%~dp0"
if "%ORIGEM:~-1%"=="\" set "ORIGEM=%ORIGEM:~0,-1%"

REM ---------------------------------------------------------------
REM Espelha os arquivos do PROJETO (nao o ambiente virtual) para
REM cada um dos destinos de backup/monitoramento abaixo.
REM
REM Fica de fora do espelhamento (nao e copiado nem apagado nos
REM destinos por causa dele):
REM   - Lib, Scripts, Include, share, pyvenv.cfg  -> ambiente virtual
REM     Python (reinstalavel com 2_INSTALAR_maquina_nova.bat, nao
REM     precisa de backup)
REM   - build, dist                        -> gerados pelo empacotamento
REM     do .exe (3_EMPACOTAR_exe.bat)
REM   - __pycache__                        -> cache do Python
REM   - .git                                -> controle de versao (ja
REM     tem backup no GitHub)
REM   - .claude                             -> configuracao local do
REM     Claude Code
REM
REM ATENCAO: cada destino e espelhado (/MIR) - qualquer arquivo que
REM exista SO no destino (e nao entre nas exclusoes acima) e apagado
REM para o destino ficar identico a origem.
REM
REM E chamado automaticamente a cada 15 min por uma Tarefa Agendada
REM do Windows ("Sync Panorama Backups"). Tambem pode ser rodado
REM manualmente a qualquer momento, so clicando 2x neste arquivo.
REM ---------------------------------------------------------------

call :sync "C:\Users\diego.figueiredo\OneDrive\2025 2026 Onedrive\MONITORAMENTO\PANORAMA\Ambiente Virtual" "_sync_onedrive.log"
call :sync "S:\00 NOVA REDE - COAM\venv" "_sync_coam.log"

exit /b 0

:sync
set "DESTINO=%~1"
set "LOGNOME=%~2"
robocopy "%ORIGEM%" "%DESTINO%" /MIR ^
  /XD "Lib" "Scripts" "Include" "share" "build" "dist" "__pycache__" ".git" ".claude" ^
  /XF "Thumbs.db" "pyvenv.cfg" "desktop.ini" "_sync_onedrive.log" "_sync_coam.log" ^
  /XA:SH ^
  /FFT /R:2 /W:5 /NFL /NDL /NP /NJH /NJS ^
  /LOG:"%ORIGEM%\%LOGNOME%"
exit /b 0
