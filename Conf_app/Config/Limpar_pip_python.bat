@echo off
title Auto-verificador e reinstalador de ambiente Python
setlocal enabledelayedexpansion

REM ======== CONFIGURAÇÕES ========
set "BASE_DIR=%~dp0"
set "PYTHON_PATH=%BASE_DIR%..\app_py\Python313\Python313\python.exe"
set "REQ_FILE=%BASE_DIR%..\app_py\requirements.txt"
set "MAIN_FILE=%BASE_DIR%..\app_py\main.py"
set "LOG_FILE=%BASE_DIR%pip_repair_log.txt"
set "TEST_LOG=%BASE_DIR%test_output.txt"
set "SSL_URL=https://github.com/python/cpython/raw/main/PCbuild/amd64"

echo ===============================================
echo   AUTO-VERIFICADOR E REINSTALADOR PYTHON
echo ===============================================
echo.

REM ======== VERIFICA PYTHON ========
if not exist "%PYTHON_PATH%" (
    echo [ERRO] Python nao encontrado em:
    echo "%PYTHON_PATH%"
    echo.
    echo Certifique-se de que a pasta app_py\Python313 existe.
    pause
    exit /b 1
)

REM ======== TESTA PIP ========
echo [1/9] Verificando pip e SSL...
"%PYTHON_PATH%" -m pip -V >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Pip ausente. Tentando reparar...
    "%PYTHON_PATH%" -m ensurepip --upgrade >nul 2>&1
    "%PYTHON_PATH%" -m pip install --upgrade pip setuptools wheel >nul 2>&1
)

REM ======== TESTA SSL ========
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

REM ======== AJUSTA requirements.txt ========
echo [2/9] Corrigindo requirements.txt...
if not exist "%REQ_FILE%" (
    echo [ERRO] Arquivo requirements.txt nao encontrado:
    echo "%REQ_FILE%"
    pause
    exit /b 1
)

set "TMP_FILE=%BASE_DIR%req_fixed.txt"
copy "%REQ_FILE%" "%TMP_FILE%" >nul 2>&1

powershell -Command "(Get-Content '%TMP_FILE%') -replace '\bPIL\b','Pillow' -replace '\bfitz\b','PyMuPDF' -replace 'ttkbootstrap.*','ttkbootstrap==1.10.1' | Set-Content '%TMP_FILE%' -Encoding UTF8"
echo [INFO] Corrigido: PIL → Pillow / fitz → PyMuPDF / ttkbootstrap → versão 1.10.1

REM ======== REMOVE PACOTES ========
echo [3/9] Limpando ambiente anterior...
"%PYTHON_PATH%" -m pip freeze > "%BASE_DIR%all.txt" 2>nul
for /f "usebackq delims=" %%P in ("%BASE_DIR%all.txt") do (
    "%PYTHON_PATH%" -m pip uninstall -y %%P >> "%LOG_FILE%" 2>&1
)
del "%BASE_DIR%all.txt" >nul 2>&1

REM ======== ATUALIZA FERRAMENTAS ========
echo [4/9] Atualizando pip, setuptools e wheel...
"%PYTHON_PATH%" -m pip install --upgrade pip setuptools wheel >> "%LOG_FILE%" 2>&1

REM ======== INSTALA PACOTES ========
echo [5/9] Instalando pacotes corrigidos...
"%PYTHON_PATH%" -m pip install -r "%TMP_FILE%" >> "%LOG_FILE%" 2>&1

REM ======== TESTA E CORRIGE main.py ========
if exist "%MAIN_FILE%" (
    echo [6/9] Verificando execucao de main.py...
    set /a ATTEMPT=1
    :TEST_LOOP
    echo.
    echo Tentativa !ATTEMPT! de teste...
    "%PYTHON_PATH%" "%MAIN_FILE%" --test-run > "%TEST_LOG%" 2>&1

    findstr /C:"Traceback" "%TEST_LOG%" >nul
    if %errorlevel%==0 (
        echo [ERRO] Erro detectado na execucao.
        for /f "tokens=1,* delims=:" %%a in ('findstr /C:"ModuleNotFoundError:" "%TEST_LOG%"') do (
            set "MISSING=%%b"
        )
        if defined MISSING (
            for /f "tokens=2 delims=''" %%X in ("!MISSING!") do (
                set "PKG=%%~X"
            )
            set "PKG=!PKG:~0,-1!"
            echo [INFO] Tentando instalar automaticamente o pacote ausente: !PKG!
            "%PYTHON_PATH%" -m pip install !PKG! >> "%LOG_FILE%" 2>&1
            set /a ATTEMPT+=1
            if !ATTEMPT! leq 3 goto TEST_LOOP
        ) else (
            echo [ALERTA] Traceback encontrado mas nao foi um erro de importacao.
            echo Veja o arquivo: %TEST_LOG%
        )
    ) else (
        echo [OK] main.py executou sem erros aparentes.
    )
) else (
    echo [AVISO] main.py nao encontrado em: "%MAIN_FILE%"
)

REM ======== LISTA PACOTES ========
echo [7/9] Pacotes instalados:
"%PYTHON_PATH%" -m pip list

REM ======== LOGS ========
echo [8/9] Salvando logs...
echo Log de instalacao: %LOG_FILE%
echo Log de teste: %TEST_LOG%

REM ======== FINALIZA ========
echo.
echo [9/9] Processo concluido!
echo ===============================================
pause
exit /b 0
