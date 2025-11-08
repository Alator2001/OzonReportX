@echo off
chcp 65001 >nul
cd /d "%~dp0"
color 09

echo.
echo ============================================================
echo          OZON REPORT X - Master Setup
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found! Opening download page...
    start "" "https://www.python.org/downloads/"
    echo.
    echo Please install Python and restart this script.
    pause
    exit /b 1
)

echo Starting setup...
echo.

python -X utf8 "scripts\first_run_setup.py"
if errorlevel 1 (
    echo.
    echo Script finished with errors. Check messages above.
    pause
    exit /b 1
)

echo.
pause