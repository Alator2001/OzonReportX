@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
color 0A

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Virtual environment not found. Running initial setup first...
    python -X utf8 "scripts\first_run_setup.py" --setup-only
    if errorlevel 1 exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] .venv was not created.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -X utf8 "scripts\first_run_setup.py" --setup-only
if errorlevel 1 (
    echo [ERROR] Dependency setup failed.
    pause
    exit /b 1
)

echo [INFO] Starting Flet UI...
echo.

".venv\Scripts\python.exe" "app_flet.py"
if errorlevel 1 (
    echo.
    echo [ERROR] UI finished with errors.
    pause
    exit /b 1
)
