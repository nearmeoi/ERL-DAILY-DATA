@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title Putar Rekaman Otomasi (Macro Player)
cd /d "%~dp0"

set PYTHON_EXE=%~dp0.venv\Scripts\python.exe

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment tidak ditemukan di .venv!
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%~dp0putar_rekaman.py"

echo.
pause
