@echo off
title Ekstraksi Foto Mobil ke Excel Otomatis
color 0A
cls
echo ========================================================
echo    PROSES OTOMATIS FOTO MOBIL KE EXCEL TERKELOMPOK
echo ========================================================
echo.
echo Sedang memeriksa foto di folder 'input_foto_mobil'...
echo.

node proses_mobil.js

echo.
echo ========================================================
echo Selesai! Buka folder 'hasil_excel' untuk melihat file.
echo ========================================================
pause
