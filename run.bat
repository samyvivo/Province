@echo off
setlocal EnableExtensions
chcp 65001 >nul
title گزارش فروش و موجودی

rem --- go to the folder containing this .bat (where app.py lives) ---
cd /d "%~dp0"

rem ---------- 1) locate a Python 3 interpreter ----------
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY (
    python -c "import sys; assert sys.version_info[0]==3" >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo.
    echo   Python 3 not found on this PC.
    echo.
    echo   Install it from:
    echo       https://www.python.org/downloads/windows/
    echo   During setup tick the box: "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

rem ---------- 2) install dependencies once ----------
if not exist ".deps_ok" (
    echo Installing dependencies ^(first run only^)...
    %PY% -m pip install --upgrade pip >nul
    %PY% -m pip install pandas openpyxl
    if errorlevel 1 (
        echo.
        echo   Dependency installation failed. Check your internet connection.
        pause
        exit /b 1
    )
    echo ok> ".deps_ok"
)

rem ---------- 3) run the app ----------
echo Starting...
%PY% "app.py"

rem ---------- 4) keep window open if the app crashed ----------
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
    echo.
    echo   App exited with code %RC%.
    pause
)
endlocal