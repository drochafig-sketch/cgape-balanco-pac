@echo off
REM =====================================================
REM  Gera o executavel (.exe) do Relatorio Gerencial PAC
REM  Rode este arquivo na MESMA pasta onde esta o script
REM  relatorio_gerencial_pac.py
REM =====================================================

echo.
echo Instalando/atualizando o PyInstaller...
python -m pip install --upgrade pyinstaller openpyxl

echo.
echo Gerando o executavel (isso pode demorar alguns minutos)...
python -m PyInstaller --onefile --noconsole ^
    --name "RelatorioGerencialPAC" ^
    --hidden-import openpyxl ^
    --hidden-import et_xmlfile ^
    relatorio_gerencial_pac.py

echo.
echo =====================================================
echo  Pronto! O executavel foi criado em:
echo  dist\RelatorioGerencialPAC.exe
echo.
echo  PROXIMO PASSO: copie o arquivo
echo  dist\RelatorioGerencialPAC.exe
echo  para a pasta de rede compartilhada, junto com:
echo    - a planilha .xlsx
echo    - PAC.png
echo    - GOVERNO.PNG
echo    - DIN.ttf e DIN-Bold.ttf (se estiver usando)
echo =====================================================
pause
