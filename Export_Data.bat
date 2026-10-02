@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

"%PY%" -X utf8 "scripts\transfer_data.py" export
pause
