@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

set "ARCHIVE=%~1"
if "%ARCHIVE%"=="" (
    for /f "delims=" %%F in ('dir /b /o-d "OzonReportX_transfer_*.zip" 2^>nul') do (
        if not defined ARCHIVE set "ARCHIVE=%%F"
    )
)
if "%ARCHIVE%"=="" (
    echo [ERROR] Положите архив OzonReportX_transfer_*.zip в папку проекта или перетащите его на этот файл.
    pause
    exit /b 1
)

"%PY%" -X utf8 "scripts\transfer_data.py" import "%ARCHIVE%"
pause
