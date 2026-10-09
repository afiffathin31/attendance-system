@echo off
title Pasang Ekstensi Absensi Chrome
echo ========================================================
echo   PANDUAN PEMASANGAN EKSTENSI ABSENSI CHROME / EDGE
echo ========================================================
echo.
echo Langkah mudah:
echo 1. Browser akan terbuka ke halaman ekstensi.
echo 2. Aktifkan 'Developer mode' di pojok kanan atas.
echo 3. Klik tombol 'Load unpacked' di pojok kiri atas.
echo 4. Pilih folder:
echo    %~dp0extension
echo.
echo ========================================================
pause
start chrome.exe "chrome://extensions" || start msedge.exe "edge://extensions"
explorer "%~dp0extension"
