@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo [run] %PY% easy-fulfill.py
echo ----------------------------------------
"%PY%" "%~dp0easy-fulfill.py"

echo ----------------------------------------
echo [run] exit code: %ERRORLEVEL%
pause
