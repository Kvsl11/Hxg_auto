@echo off
title Gerar requirements.txt automaticamente
setlocal enabledelayedexpansion

REM ============================================
REM   GERADOR AUTOMÁTICO DE requirements.txt
REM   Compatível com estrutura:
REM   Config\gerar_requirements.bat
REM   app_py\main.py
REM ============================================

REM === Caminho base: pasta onde este .bat está ===
set "BASE_DIR=%~dp0"

REM === Caminho do main.py e destino do requirements.txt ===
set "MAIN_FILE=%BASE_DIR%..\app_py\main.py"
set "REQ_FILE=%BASE_DIR%..\app_py\requirements.txt"

REM === Se não encontrar, tenta variações de nome ===
if not exist "%MAIN_FILE%" (
    if exist "%BASE_DIR%..\App_py\main.py" (
        set "MAIN_FILE=%BASE_DIR%..\App_py\main.py"
        set "REQ_FILE=%BASE_DIR%..\App_py\requirements.txt"
    ) else (
        echo [ERRO] Nao foi encontrado o arquivo main.py em:
        echo "%MAIN_FILE%"
        echo.
        echo Estrutura esperada:
        echo   %BASE_DIR%..\app_py\main.py
        echo   ou %BASE_DIR%..\App_py\main.py
        echo.
        pause
        exit /b 1
    )
)

echo ============================================
echo   GERANDO ARQUIVO requirements.txt
echo ============================================
echo.
echo [INFO] Arquivo localizado em: "%MAIN_FILE%"
echo.

REM === Remove o antigo e cria novo ===
del "%REQ_FILE%" >nul 2>&1
echo.>"%REQ_FILE%"

REM === Extrai imports do main.py ===
for /f "tokens=2 delims= " %%i in ('findstr /r "^import ^from" "%MAIN_FILE%"') do (
    set "mod=%%i"
    for /f "delims=. tokens=1" %%a in ("!mod!") do set "mod=%%a"
    echo !mod!>>"%REQ_FILE%"
)

REM === Remove duplicados e módulos nativos ===
set "TMP_FILE=%REQ_FILE%.tmp"
type nul > "%TMP_FILE%"

set "CORE_MODULES=asyncio base64 calendar collections concurrent contextlib copy csv ctypes datetime difflib email enum fnmatch functools glob hashlib heapq http io itertools json linecache locale logging math mimetypes numbers operator os pathlib pickle pkgutil platform queue random re shutil signal socket sqlite3 ssl statistics string struct subprocess sys tempfile textwrap threading time tkinter types typing unicodedata urllib uuid warnings weakref xml zipfile webbrowser argparse inspect dataclasses"

for /f "usebackq delims=" %%M in ("%REQ_FILE%") do (
    set "mod=%%M"
    if not "!mod!"=="" (
        echo !CORE_MODULES! | findstr /i "\<!mod!\>" >nul
        if errorlevel 1 (
            findstr /ix "!mod!" "%TMP_FILE%" >nul || echo !mod!>>"%TMP_FILE%"
        )
    )
)

move /y "%TMP_FILE%" "%REQ_FILE%" >nul

REM === Adiciona setuptools e wheel ===
echo setuptools>>"%REQ_FILE%"
echo wheel>>"%REQ_FILE%"

echo.
echo ============================================
echo  Arquivo requirements.txt gerado com sucesso!
echo ============================================
echo Local: %REQ_FILE%
echo.
type "%REQ_FILE%"
echo.
pause
exit /b 0
