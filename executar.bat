@echo off
setlocal
cd /d "%~dp0"
set "PY="
set "PYW="
where py >nul 2>&1 && (set "PY=py -3" & set "PYW=pyw -3")
if not defined PY (
    where python >nul 2>&1 && (set "PY=python" & set "PYW=pythonw")
)
if not defined PY (
    echo Python 3 nao encontrado. Instale o Python 3.10 ou mais novo e rode este arquivo de novo.
    pause
    exit /b 1
)
%PY% -c "import openpyxl, lxml, tkinter" >nul 2>&1
if errorlevel 1 (
    echo Instalando as dependencias do programa...
    %PY% -m pip install --user -r requirements.txt
    if errorlevel 1 (
        echo Nao foi possivel instalar as dependencias. Confira a conexao com a internet e tente de novo.
        pause
        exit /b 1
    )
)
start "" %PYW% app.py
