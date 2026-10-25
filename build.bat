@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Building SalesReport.exe
cd /d "%~dp0"

if not exist "icon.ico" (
    echo WARNING: icon.ico not found in this folder.
    echo          The exe will use the default Python icon.
    echo.
)

rem ---------- 1) install / upgrade PyInstaller ----------
echo Installing PyInstaller ^(one-time^)...
py -3 -m pip install --upgrade pyinstaller
if errorlevel 1 (
    echo Failed to install PyInstaller. Check your internet / Python install.
    pause
    exit /b 1
)

rem ---------- 2) clean previous build ----------
if exist build            rmdir /s /q build
if exist dist             rmdir /s /q dist
if exist SalesReport.spec del /q SalesReport.spec

rem ---------- 3) build ----------
echo.
echo Building SalesReport.exe ...
py -3 -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "SalesReport" ^
    --icon "icon.ico" ^
    --add-data "icon.ico;." ^
    --collect-submodules openpyxl ^
    app.py
if errorlevel 1 (
    echo.
    echo BUILD FAILED — see messages above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   BUILD COMPLETE
echo   Output: dist\SalesReport.exe
echo ============================================
echo.
echo Remember to place the "json" folder next to SalesReport.exe before running.
pause