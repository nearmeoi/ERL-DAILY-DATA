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
echo.
echo Pilihan Mode Eksekusi:
echo   [1] Semua Tahap (1 - 9) [Full Pipeline] (Default)
echo   [2] Posisi Stok Cabang (Tahap 1 - 6) [Pagi - Tanpa Overwrite Siang]
echo   [3] Buku Masuk & FIFO Saja (Tahap 7) [Siang - Tanpa Overwrite Format Cabang]
echo   [4] DO Summary External Saja (Tahap 8) [Sore]
echo   [5] Cek Item & Antrean SO FIFO [Inspect Mode]
set /p "MODE_CHOICE=Pilih mode (1/2/3/4/5, Enter=1): "

if "%MODE_CHOICE%"=="5" (
    echo.
    set /p "ITEM_QUERY=Ketik Kode Item atau Judul Buku (contoh: 0024200580): "
    echo.
    "%~dp0App\.venv\Scripts\python.exe" "%~dp0App\main.py" "%FOLDER%" --inspect "%ITEM_QUERY%"
    echo.
    pause
    exit /b
)

set "MODE_ARG=all"
if "%MODE_CHOICE%"=="2" set "MODE_ARG=stok"
if "%MODE_CHOICE%"=="3" set "MODE_ARG=bukmas"
if "%MODE_CHOICE%"=="4" set "MODE_ARG=dosum"

"%~dp0App\.venv\Scripts\python.exe" "%~dp0App\main.py" "%FOLDER%" --mode %MODE_ARG%
echo.
pause
