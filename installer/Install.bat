@echo off
chcp 65001 >nul 2>&1
title Pebble — Installation
echo.
echo   ========================================
echo    Pebble — Installing...
echo   ========================================
echo.

:: Run the PowerShell installer from the same directory
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1"

if %errorlevel% neq 0 (
    echo.
    echo   Installation failed. See errors above.
    echo.
    pause
)
