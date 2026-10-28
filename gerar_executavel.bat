@echo off
setlocal
cd /d "%~dp0"
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY (
    where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo Python 3 nao encontrado. Instale o Python 3.10 ou mais novo e rode este arquivo de novo.
    pause
    exit /b 1
)
echo Instalando PyInstaller e dependencias...
%PY% -m pip install --user pyinstaller -r requirements.txt
if errorlevel 1 goto erro
echo Gerando o executavel...
%PY% -m PyInstaller --noconfirm --clean --windowed --name GeradorListaMaterial --icon recursos\icone.ico ^
    --add-data "recursos;recursos" --add-data "modelos;modelos" app.py
if errorlevel 1 goto erro
xcopy /E /I /Y modelos dist\GeradorListaMaterial\modelos >nul
xcopy /E /I /Y recursos dist\GeradorListaMaterial\recursos >nul
echo.
echo Pronto: dist\GeradorListaMaterial\GeradorListaMaterial.exe
echo Para distribuir, copie a pasta dist\GeradorListaMaterial inteira.
pause
exit /b 0
:erro
echo A geracao do executavel falhou. Veja as mensagens acima.
pause
exit /b 1
