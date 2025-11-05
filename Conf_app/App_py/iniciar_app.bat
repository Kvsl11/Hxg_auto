@echo off
REM 🚀 Iniciador universal do Hxg_auto (com espera segura)

REM Define diretório base (onde o BAT está salvo)
cd /d "%~dp0"

REM Caminhos absolutos (independem do local de execução)
set "BASE_DIR=%~dp0"
set "PYTHONW=%BASE_DIR%Python313\pythonw.exe"
set "UPDATER=%BASE_DIR%updater.py"

echo ============================================
echo 🚀 Iniciando Hxg_auto
echo ============================================
echo.
echo Verificando caminhos...
echo Pythonw: "%PYTHONW%"
echo Updater: "%UPDATER%"
echo.

REM Verifica se o Python existe
if not exist "%PYTHONW%" (
    echo ❌ ERRO: Python nao encontrado em "%PYTHONW%"
    pause
    exit /b 1
)

REM Verifica se o updater existe
if not exist "%UPDATER%" (
    echo ❌ ERRO: updater.py nao encontrado em "%UPDATER%"
    pause
    exit /b 1
)

REM Executa o updater e espera ele concluir (sem usar START)
echo 🔍 Executando updater (aguarde)...
"%PYTHONW%" "%UPDATER%"

echo.
echo ✅ Processo concluído.
pause
exit /b 0
