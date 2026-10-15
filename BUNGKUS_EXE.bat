@echo off
title Bungkus Executable & Send to Desktop
echo ========================================================
echo   MEMBUNGKUS DAILY DATA STANDALONE & KIRIM KE DESKTOP
echo ========================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_exe.ps1"
echo.
pause
