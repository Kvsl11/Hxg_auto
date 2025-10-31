@echo off
title Gerar requirements.txt automaticamente
setlocal enabledelayedexpansion

REM ============================================
REM   GERADOR AUTOMÁTICO DE requirements.txt
REM   Compatível com qualquer máquina/pasta
REM ============================================

REM === CONFIGURAÇÕES AUTOMÁTICAS ===
REM Define a pasta base (onde o script está localizado)
set "BASE_DIR=%~dp0"

REM Caminho do arquivo principal e destino do requirements.txt
set "MAIN_FILE=%BASE_DIR%App_py\main.py"
set "REQ_FILE=%BASE_DIR%App_py\requirements.txt"

echo ============================================
echo   GERANDO ARQUIVO requirements.txt
echo ============================================
echo.

if not exist "%MAIN_FILE%" (
    echo [ERRO] Nao foi encontrado o arquivo main.py em:
    echo "%MAIN_FILE%"
    echo.
    echo Certifique-se de que o main.py esteja dentro da pasta App_py.
    pause
    exit /b 1
)

REM === REMOVE O ARQUIVO ANTIGO E CRIA UM NOVO ===
del "%REQ_FILE%" >nul 2>&1
echo.>"%REQ_FILE%"

echo Lendo imports do arquivo: "%MAIN_FILE%"
echo.

REM === VARRE TODAS AS LINHAS DE IMPORT ===
for /f "tokens=2 delims= " %%i in ('findstr /r "^import ^from" "%MAIN_FILE%"') do (
    set "mod=%%i"
    for /f "delims=. tokens=1" %%a in ("!mod!") do set "mod=%%a"
    echo !mod!>>"%REQ_FILE%"
)

REM === LIMPA DUPLICADOS E MÓDULOS NATIVOS ===
set "TMP_FILE=%REQ_FILE%.tmp"
type nul > "%TMP_FILE%"

REM Lista de módulos internos (não serão incluídos)
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

REM === ADICIONA SETUPTOOLS E WHEEL ===
echo setuptools>>"%REQ_FILE%"
echo wheel>>"%REQ_FILE%"

echo.
echo ============================================
echo  Arquivo requirements.txt gerado com sucesso!
echo  Local: %REQ_FILE%
echo ============================================
echo.
type "%REQ_FILE%"
echo.
pause
exit /b 0
