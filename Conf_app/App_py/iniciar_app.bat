@echo off
REM ===============================
REM Executa updater.py sem console
REM ===============================

REM Garante que o script rode no diretório atual
cd /d "%~dp0"

REM Caminhos absolutos
set "BASE_DIR=%~dp0"
set "PY_DIR=%BASE_DIR%Python313"
set "PYTHONW=%PY_DIR%\pythonw.exe"
set "PY_ZIP=%BASE_DIR%Python313.zip"
set "SCRIPT=%BASE_DIR%updater.py"

REM Extrai o Python embarcado caso ainda esteja compactado
if not exist "%PYTHONW%" (
    if exist "%PY_ZIP%" (
        echo [INFO] Preparando ambiente Python...
        powershell -NoProfile -Command "Expand-Archive -LiteralPath '%PY_ZIP%' -DestinationPath '%BASE_DIR%' -Force" >nul 2>&1
    )
)

REM Verifica se pythonw.exe existe
if not exist "%PYTHONW%" (
    msg * "ERRO: pythonw.exe não encontrado em %PYTHONW%"
    exit /b
)

REM Verifica se updater.py existe
if not exist "%SCRIPT%" (
    msg * "ERRO: updater.py não encontrado em %SCRIPT%"
    exit /b
)

REM Executa silenciosamente usando pythonw (sem console)
start "" /b "%PYTHONW%" "%SCRIPT%"
exit
