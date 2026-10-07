@echo off
title Daily Data - Sistem Otomasi Pipeline
cd /d "%~dp0"

echo ========================================================
echo                 MEMBUKA DAILY DATA
echo ========================================================
echo.

set PYTHON_EXE=%~dp0App\.venv\Scripts\python.exe

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment Python tidak ditemukan!
    pause
    exit /b 1
)

start "" "%PYTHON_EXE%" "%~dp0App\gui.py"
exit
