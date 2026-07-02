@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem --- venv 확보(없으면 생성) ---
set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" (
  echo [setup] .venv 가 없어 새로 만듭니다...
  py -3 -m venv "%~dp0.venv" 2>nul || python -m venv "%~dp0.venv"
)
if not exist "%PY%" set "PY=python"

rem --- 의존성 동기화(새 requirements 자동 설치) ---
echo [setup] 의존성 확인/설치 중...
"%PY%" -m pip install -q --upgrade pip
"%PY%" -m pip install -q -r "%~dp0requirements.txt"
if not "%ERRORLEVEL%"=="0" (
  echo ----------------------------------------
  echo [setup] 의존성 설치 실패. 위 메시지를 확인하세요.
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
