@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
color 09

python -X utf8 "config\show_banner.py" 2>nul

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    start "" "https://www.python.org/downloads/"
    pause
    exit /b 1
)

echo [INFO] Starting setup...
echo.

python -X utf8 "scripts\first_run_setup.py"
if errorlevel 1 (
    echo [ERROR] Script finished with errors.
    pause
    exit /b 1
)

pause
