@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title Otomasi Export Axapta
cd /d "%~dp0"

echo ========================================================
echo          MENJALANKAN OTOMASI EXPORT DYNAMICS AX
echo ========================================================
echo.

set PYTHON_EXE=%~dp0.venv\Scripts\python.exe

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment tidak ditemukan di .venv!
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%~dp0auto_axapta.py"

echo.
echo Selesai. Tekan tombol apa saja untuk menutup...
pause >nul
