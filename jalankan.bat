@echo off
title ExcelBot AI - Bot Otonom Foto ke Excel
color 0A
echo ==========================================================
echo       MEMULAI EXCELBOT AI (FOTO KE EXCEL OTOMATIS)
echo ==========================================================
echo.

:: Cek keberadaan Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js belum terpasang di komputer ini!
    echo Silakan download dan pasang Node.js dari: https://nodejs.org
    pause
    exit /b 1
)

:: Cek dependensi node_modules
if not exist "node_modules\" (
    echo [INFO] Memasang paket dependensi (ExcelJS)...
    call npm install
    echo [OK] Paket berhasil dipasang.
    echo.
)

echo [INFO] Menyalakan server aplikasi...
start http://localhost:3000
node server.js

pause
