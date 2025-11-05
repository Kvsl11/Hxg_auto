@echo off
REM 🚀 Iniciador universal do Hxg_auto

REM Define diretório base (onde o BAT está salvo)
cd /d "%~dp0"

REM Caminhos absolutos (sempre válidos, independente de onde for executado)
set "BASE_DIR=%~dp0"
set "PYTHONW=%BASE_DIR%Python313\pythonw.exe"
set "UPDATER=%BASE_DIR%updater.py"

REM Mostra os caminhos detectados (para debug)
echo Verificando caminhos...
echo Pythonw: "%PYTHONW%"
echo Updater: "%UPDATER%"
echo.

REM Verifica se o Python existe
if not exist "%PYTHONW%" (
    echo [ERRO] Python nao encontrado em: "%PYTHONW%"
    pause
    exit /b 1
)

REM Verifica se o updater existe
if not exist "%UPDATER%" (
    echo [ERRO] updater.py nao encontrado em: "%UPDATER%"
    pause
    exit /b 1
)

REM Executa o updater sem console
echo Iniciando o Updater...
start "" "%PYTHONW%" "%UPDATER%"
exit /b 0
