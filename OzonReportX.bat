@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python не найден в PATH. Установите Python 3.10+ и попробуйте снова.
  pause
  exit /b 1
)

python "scripts\first_run_setup.py"
pause

