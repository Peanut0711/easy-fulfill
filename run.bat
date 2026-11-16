@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"
set "SETUP=%~dp0tools\prepare_run.py"
if exist "%PY%" (
  "%PY%" "%SETUP%"
) else (
  py -3 -c "pass" >nul 2>&1
  if errorlevel 1 (
    python "%SETUP%"
  ) else (
    py -3 "%SETUP%"
  )
)

if errorlevel 1 (
  echo ----------------------------------------
  echo [setup] 실행 환경 준비 실패. 위 메시지를 확인하세요.
  pause
  exit /b 1
)

echo [run] %PY% easy-fulfill.py
echo ----------------------------------------
"%PY%" "%~dp0easy-fulfill.py"
set "EC=%ERRORLEVEL%"

rem Close console automatically on normal exit (0); pause only on error.
if not "%EC%"=="0" (
  echo ----------------------------------------
  echo [run] exit code: %EC%
  pause
)
