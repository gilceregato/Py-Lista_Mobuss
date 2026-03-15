@echo off
setlocal

pushd "%~dp0"

rem Uso:
rem   start_app.bat -> Flask (app.py)

set "VENV_DIR=.venv"
set "PYTHON=python"

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Criando ambiente virtual em %VENV_DIR%...
    %PYTHON% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Erro ao criar o ambiente virtual.
        pause
        exit /b 1
    )
)

echo Atualizando pip e instalando dependencias...
"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
    echo Erro ao atualizar o pip.
    pause
    exit /b 1
)

"%VENV_DIR%\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo Erro ao instalar dependencias.
    pause
    exit /b 1
)

set FLASK_APP=app.py
set FLASK_DEBUG=1
set FLASK_HOST=127.0.0.1
set FLASK_PORT=8000

set BROWSER_HOST=%FLASK_HOST%
if /i "%BROWSER_HOST%"=="0.0.0.0" set BROWSER_HOST=127.0.0.1

echo Iniciando Flask...
start "" "%VENV_DIR%\Scripts\python.exe" -m flask run --host %FLASK_HOST% --port %FLASK_PORT%

rem Aguarda o servidor responder e abre o navegador
powershell -NoProfile -Command "for ($i=0; $i -lt 30; $i++) { try { Invoke-WebRequest -UseBasicParsing -TimeoutSec 1 http://%BROWSER_HOST%:%FLASK_PORT%/ | Out-Null; exit 0 } catch { Start-Sleep -Milliseconds 500 } } exit 1"
if errorlevel 1 (
    echo Nao foi possivel detectar o servidor. Abrindo o navegador mesmo assim...
)
start "" "http://%BROWSER_HOST%:%FLASK_PORT%/"
if errorlevel 1 (
    echo Erro ao iniciar o Flask.
    pause
    exit /b 1
)

exit /b 0
