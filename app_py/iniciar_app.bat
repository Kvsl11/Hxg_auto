@echo off
REM Garante que o script rode na pasta onde ele está
cd /d "%~dp0"

echo Iniciando Hxg_auto...

REM Define caminho absoluto do Python e do script
set "PYTHONW=%~dp0Python313\Python313\pythonw.exe"
set "SCRIPT=%~dp0updater.py"

REM Verifica se o executável pythonw existe
if not exist "%PYTHONW%" (
    echo [ERRO] pythonw.exe não encontrado em: "%PYTHONW%"
    pause
    exit /b 1
)

REM Verifica se o script existe
if not exist "%SCRIPT%" (
    echo [ERRO] updater.py não encontrado em: "%SCRIPT%"
    pause
    exit /b 1
)

REM Executa o Python sem console
start "" "%PYTHONW%" "%SCRIPT%"
exit
