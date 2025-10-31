@echo off
cd /d "%~dp0"
echo Iniciando Hxg_auto...
start "" "%~dp0Python313\pythonw.exe" "%~dp0updater.py"
exit
