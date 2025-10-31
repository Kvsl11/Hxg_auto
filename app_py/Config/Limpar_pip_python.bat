@echo off
title Auto-verificador e reinstalador de ambiente Python
setlocal enabledelayedexpansion

REM ======== CONFIGURAÇÕES DINÂMICAS ========
REM Define o diretório base (pasta onde este script está)
set "BASE_DIR=%~dp0"

REM Caminho completo do Python e do requirements.txt
set "PYTHON_PATH=%BASE_DIR%..\app_py\Python313\Python313\python.exe"
set "REQ_FILE=%BASE_DIR%..\app_py\requirements.txt"

REM Caminhos auxiliares
set "LOG_FILE=%BASE_DIR%pip_repair_log.txt"
set "SSL_URL=https://github.com/python/cpython/raw/main/PCbuild/amd64"

echo ===============================================
echo   VERIFICADOR E REPARADOR DE AMBIENTE PYTHON
echo ===============================================
echo.

REM ======== VERIFICA SE PYTHON EXISTE ========
if not exist "%PYTHON_PATH%" (
    echo [ERRO] Python nao encontrado em:
    echo "%PYTHON_PATH%"
    echo.
    echo Certifique-se de que a pasta app_py\Python313 existe.
    pause
    exit /b 1
)

REM ======== CHECA SE O PIP FUNCIONA ========
echo [1/7] Verificando pip e SSL...
"%PYTHON_PATH%" -m pip -V >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Pip ausente. Tentando reparar...
    "%PYTHON_PATH%" -m ensurepip --upgrade >nul 2>&1
    "%PYTHON_PATH%" -m pip install --upgrade pip setuptools wheel >nul 2>&1
)

REM ======== TESTA DOWNLOAD PARA DETECTAR ERRO DE SSL ========
"%PYTHON_PATH%" -m pip install requests --dry-run >nul 2>&1
if %errorlevel% neq 0 (
    echo [ALERTA] Problema SSL detectado. Tentando reparar...
    for %%D in (libssl-3.dll libcrypto-3.dll) do (
        if not exist "%BASE_DIR%%%D" (
            echo Baixando %%D ...
            powershell -Command "Invoke-WebRequest -Uri '%SSL_URL%/%%D' -OutFile '%BASE_DIR%%%D'" >nul 2>&1
        )
        copy /Y "%BASE_DIR%%%D" "%BASE_DIR%" >nul 2>&1
        copy /Y "%BASE_DIR%%%D" "%BASE_DIR%..\app_py\Python313" >nul 2>&1
    )
    echo [INFO] DLLs SSL adicionadas (libssl-3.dll / libcrypto-3.dll)
)

REM ======== VERIFICA E AJUSTA requirements.txt ========
echo [2/7] Validando e corrigindo requirements.txt...
if not exist "%REQ_FILE%" (
    echo [ERRO] Nao foi encontrado o arquivo requirements.txt:
    echo "%REQ_FILE%"
    echo.
    echo Gere-o com o script gerar_requirements.bat na pasta Config.
    pause
    exit /b 1
)

set "TMP_FILE=%BASE_DIR%req_fixed.txt"
copy "%REQ_FILE%" "%TMP_FILE%" >nul 2>&1

REM Corrige nomes comuns de pacotes
powershell -Command "(Get-Content '%TMP_FILE%') -replace '\bPIL\b','Pillow' -replace '\bfitz\b','PyMuPDF' | Set-Content '%TMP_FILE%' -Encoding UTF8"

echo [INFO] Corrigido: PIL → Pillow / fitz → PyMuPDF

REM ======== MOSTRA PACOTES ========
echo.
echo Pacotes a serem instalados:
type "%TMP_FILE%"
echo.
pause

REM ======== REMOVE PACOTES EXISTENTES ========
echo [3/7] Removendo pacotes existentes...
"%PYTHON_PATH%" -m pip freeze > "%BASE_DIR%all.txt" 2>nul
for /f "usebackq delims=" %%P in ("%BASE_DIR%all.txt") do (
    echo Removendo %%P ...
    "%PYTHON_PATH%" -m pip uninstall -y %%P >> "%LOG_FILE%" 2>&1
)
del "%BASE_DIR%all.txt" >nul 2>&1

REM ======== ATUALIZA PIP E FERRAMENTAS ========
echo.
echo [4/7] Atualizando pip, setuptools e wheel...
"%PYTHON_PATH%" -m pip install --upgrade pip setuptools wheel >> "%LOG_FILE%" 2>&1

REM ======== PERGUNTA SE DESEJA REINSTALAR ========
echo.
echo Deseja instalar os pacotes do requirements.txt agora? (Y/N)
set /p USER_CHOICE="> "

if /i "%USER_CHOICE%"=="N" (
    echo.
    echo [INFO] Instalacao cancelada pelo usuario.
    echo Apenas limpeza e atualizacao de pip/setuptools foram realizadas.
    echo.
    echo Log salvo em: %LOG_FILE%
    echo ===============================================
    pause
    exit /b 0
)

if /i not "%USER_CHOICE%"=="Y" (
    echo.
    echo [ERRO] Opcao invalida. Digite apenas Y ou N.
    pause
    exit /b 1
)

REM ======== INSTALA PACOTES ========
echo.
echo [5/7] Instalando pacotes corrigidos...
"%PYTHON_PATH%" -m pip install -r "%TMP_FILE%" >> "%LOG_FILE%" 2>&1

if %errorlevel% neq 0 (
    echo.
    echo [AVISO] Ocorreram erros durante a instalacao. Veja o log:
    echo %LOG_FILE%
) else (
    echo.
    echo [SUCESSO] Todos os pacotes foram instalados corretamente!
)

REM ======== MOSTRA RESULTADO ========
echo.
echo [6/7] Pacotes instalados:
"%PYTHON_PATH%" -m pip list

REM ======== FINALIZA ========
echo.
echo [7/7] Processo concluido!
echo Log salvo em: %LOG_FILE%
echo ===============================================
pause
exit /b 0
