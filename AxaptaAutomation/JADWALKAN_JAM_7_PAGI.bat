@echo off
title Jadwalkan Otomasi Jam 07:00 Pagi
cd /d "%~dp0"

echo ========================================================
echo       MENDAFTARKAN JADWAL KE WINDOWS TASK SCHEDULER
echo ========================================================
echo.
echo Menjadwalkan tugas harian pada pukul 07:00:00 AM...
echo.

set SCRIPT_PATH=%~dp0JALANKAN_SEKARANG.bat

schtasks /create /tn "OtomasiExportAxapta" /tr "\"%SCRIPT_PATH%\"" /sc daily /st 07:00 /f /rl HIGHEST

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================================
    echo  [BERHASIL] Tugas "OtomasiExportAxapta" telah didaftarkan!
    echo  Jadwal: Setiap hari pukul 07:00 pagi.
    echo ========================================================
) else (
    echo.
    echo [GAGAL] Terjadi kesalahan saat mendaftarkan tugas.
    echo Coba jalankan file ini dengan klik kanan -> "Run as Administrator".
)

echo.
pause
