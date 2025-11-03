@echo off
REM ===============================
REM Executa updater.py sem console
REM ===============================

REM Garante que o script rode no diretório atual
cd /d "%~dp0"

REM Caminhos absolutos
set "PYTHONW=%~dp0Python313\Python313\pythonw.exe"
set "SCRIPT=%~dp0updater.py"

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
