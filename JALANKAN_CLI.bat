@echo off
title Daily Data (Console Mode)
cd /d "%~dp0"
set "FOLDER=%~1"
if "%FOLDER%"=="" (
    echo ========================================================
    echo                       DAILY DATA
    echo ========================================================
    echo Folder harian yang tersedia:
    for /d %%D in ("%~dp0*") do if exist "%%D\FORMAT POSISISTOK*.xlsx" echo    - %%~nxD
    echo.
    set /p "FOLDER=Ketik nama folder (contoh: 07 OKT 2026), Enter = terbaru: "
)
"%~dp0App\.venv\Scripts\python.exe" "%~dp0App\main.py" "%FOLDER%"
echo.
pause
