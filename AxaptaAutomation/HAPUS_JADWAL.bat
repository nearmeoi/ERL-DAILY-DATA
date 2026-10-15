@echo off
title Hapus Jadwal Otomasi
cd /d "%~dp0"

echo ========================================================
echo       MENGHAPUS JADWAL DARI WINDOWS TASK SCHEDULER
echo ========================================================
echo.

schtasks /delete /tn "OtomasiExportAxapta" /f

if %ERRORLEVEL% EQU 0 (
    echo [BERHASIL] Tugas "OtomasiExportAxapta" telah dihapus dari Task Scheduler.
) else (
    echo [INFO] Tugas tidak ditemukan atau sudah dihapus sebelumnya.
)

echo.
pause
