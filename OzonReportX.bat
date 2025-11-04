@echo off
setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [❌] Python не найден.
  echo Установите Python 3.10+ с сайта https://www.python.org/downloads/
  echo и при установке обязательно отметьте "Add Python to PATH".
  pause
  exit /b 1
)

echo [✅] Python найден.
echo Запуск программы...

python "scripts\first_run_setup.py"
pause

