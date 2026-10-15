@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title Rekam Alur Otomasi (Macro Recorder)
cd /d "%~dp0"

set PYTHON_EXE=%~dp0.venv\Scripts\python.exe

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment tidak ditemukan di .venv!
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%~dp0rekam_alur.py"

echo.
pause
