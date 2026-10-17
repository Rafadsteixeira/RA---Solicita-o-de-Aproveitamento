@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    py -m venv --without-pip ".venv"
    if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" -c "import openpyxl, pypdf, cryptography" >nul 2>&1
if errorlevel 1 (
    echo Instalando as dependencias do R.A - Equivalencias no ambiente local...
    py -m pip install --target "%~dp0.venv\Lib\site-packages" -r "%~dp0requirements.txt"
    if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" "%~dp0main.py"
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo Nao foi possivel iniciar o sistema. Verifique o Python e a conexao com a internet.
pause
exit /b 1
