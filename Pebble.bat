@echo off
chcp 65001 >nul 2>&1
title Pebble
cd /d "%~dp0"

:: Kill previous instance if running
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8000.*LISTENING"') do taskkill /F /PID %%a >nul 2>&1

:: Create venv if missing
if not exist ".venv" (
    echo  First run — setting up...
    python -m venv .venv
)

call .venv\Scripts\activate.bat
pip install -q -r requirements.txt 2>nul

:: Load .env file (sets ANTHROPIC_API_KEY etc.)
if exist ".env" (
    for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
        if not "%%A"=="" if not "%%A:~0,1%"=="#" set "%%A=%%B"
    )
)

echo.
echo  ========================================
echo   Pebble: http://localhost:8000
echo  ========================================
echo.

start http://localhost:8000
python -m uvicorn backend.main:app --host 0.0.0.0 --port 8000
pause
