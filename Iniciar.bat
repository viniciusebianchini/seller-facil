@echo off
cd /d "%~dp0"
set "SHOPEE_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%SHOPEE_PYTHON%" goto bundled
if exist ".venv\Scripts\python.exe" goto venv
echo Preparando o aplicativo pela primeira vez...
py -3 -m venv .venv
if errorlevel 1 goto missing
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto missing
:venv
.venv\Scripts\python.exe server.py
goto end
:bundled
"%SHOPEE_PYTHON%" server.py
goto end
:missing
echo Instale Python 3.11 ou superior e execute novamente.
:end
pause
