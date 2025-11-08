@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
cd /d "%~dp0"
color 09

echo   #######  #######  #######  ###  ##           #######  #######  #######  #######  ####### ########           ##   ## 
echo   ##   ##      ###  ##   ##  #### ##           ##   ##  ##       ##   ##  ##   ##  ##   ##    ##               ##### 
echo   ##   ##    ###    ##   ##  ## ####           #######  ####     #######  ##   ##  #######    ##                ### 
echo   ##   ##  ###      ##   ##  ##  ###           ##  ##   ##       ##       ##   ##  ##  ##     ##               ##### 
echo   #######  #######  #######  ##   ##           ##   ##  #######  ##       #######  ##   ##    ##              ##   ## 
echo.........................................................................................................................
echo                                           === OzonReportX — Master Setup ===

python --version >nul 2>nul
if %errorlevel% neq 0 (
    start "" "https://www.python.org/ftp/python/3.14.0"
    echo Для работы программы, необходимо установить Python.
    pause
    exit /b
)

set PY=python
%PY% -X utf8 "scripts\first_run_setup.py"
if errorlevel 1 (
  echo [X] Скрипт завершился с ошибкой. Проверьте сообщение выше.
  pause
  exit /b 1
)

echo.
pause

