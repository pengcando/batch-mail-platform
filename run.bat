@echo off
setlocal
cd /d "%~dp0"

set "VENV=%~dp0.venv"
set "PYTHON=%VENV%\Scripts\python.exe"

if not exist "%PYTHON%" (
    echo [1/3] Creating virtual environment...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo ERROR: failed to create venv. Please install Python 3.10+ and add it to PATH.
        pause
        exit /b 1
    )
    echo [2/3] Installing dependencies...
    "%PYTHON%" -m pip install --upgrade pip --quiet
    "%PYTHON%" -m pip install -r requirements.txt --quiet
)

echo [3/3] Starting server: http://127.0.0.1:8000
start "" http://127.0.0.1:8000
"%PYTHON%" -m uvicorn app.main:app --host 0.0.0.0 --port 8000
pause
