@echo off
chcp 65001 >nul 2>&1
title Pebble

:: Check if Python exists
where python >nul 2>&1
if %errorlevel%==0 (
    python --version 2>&1 | findstr /R "3\.1[0-9]" >nul 2>&1
    if %errorlevel%==0 goto :HAS_PYTHON
)

:: Try py launcher
where py >nul 2>&1
if %errorlevel%==0 (
    py -3 --version >nul 2>&1
    if %errorlevel%==0 (
        set PYTHON=py -3
        goto :SETUP
    )
)

:: No Python — download and install
echo.
echo  Python not found. Installing...
echo.

:: Download Python installer
set INSTALLER=%TEMP%\python-installer.exe
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%INSTALLER%'"

:: Install silently with PATH
%INSTALLER% /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1
del %INSTALLER%

:: Refresh PATH
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\;%LOCALAPPDATA%\Programs\Python\Python312\Scripts\;%PATH%"

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Python installation failed.
    echo  Please install Python 3.10+ manually from https://python.org
    echo.
    pause
    exit /b 1
)

:HAS_PYTHON
set PYTHON=python

:SETUP
cd /d "%~dp0"

:: Create venv if missing
if not exist ".venv" (
    echo  Creating virtual environment...
    %PYTHON% -m venv .venv
)

:: Activate venv
call .venv\Scripts\activate.bat

:: Install deps
echo  Checking dependencies...
pip install -q -r requirements.txt 2>nul

:: Start
echo.
echo  ========================================
echo   Pebble: http://localhost:8000
echo  ========================================
echo.
start http://localhost:8000
python -m uvicorn backend.main:app --host 0.0.0.0 --port 8000
pause
