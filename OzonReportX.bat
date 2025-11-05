@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
color 09

echo   #######  #######  #######  ###  ##           #######  #######  #######  #######  ####### ########           ##   ## 
echo   ##   ##      ###  ##   ##  #### ##           ##   ##  ##       ##   ##  ##   ##  ##   ##    ##               ##### 
echo   ##   ##    ###    ##   ##  ## ####           #######  ####     #######  ##   ##  #######    ##                ### 
echo   ##   ##  ###      ##   ##  ##  ###           ##  ##   ##       ##       ##   ##  ##  ##     ##               ##### 
echo   #######  #######  #######  ##   ##           ##   ##  #######  ##       #######  ##   ##    ##              ##   ## 
echo.........................................................................................................................
echo                                           === OzonReportX — Master Setup ===

REM === функция: проверка наличия Python ===
:CheckPython
where python >nul 2>nul && ( set "PY=python" & goto :Run )
where py >nul 2>nul && ( set "PY=py -3" & goto :Run )
goto :InstallPython

REM === установка Python ===
:InstallPython
echo [!] Python не найден. Пытаемся установить автоматически...

REM 1) Сначала пробуем через winget (если доступен)
winget --version >nul 2>nul
if not errorlevel 1 (
  echo [>] Используем winget для установки Python 3.x...
  REM ставим официальный Python (может показать выбор, добавляем --silent)
  winget install --id Python.Python.3 --silent --accept-package-agreements --accept-source-agreements
  if errorlevel 1 (
    echo [!] Установка через winget не удалась или отменена. Пробуем прямую установку...
  ) else (
    goto :PostInstall
  )
)

REM 2) Прямая установка: скачиваем инсталлятор python.org и ставим тихо
set "TMP_PY=%TEMP%\python-installer.exe"

REM Определяем разрядность
set "ARCH=x86"
if /i "%PROCESSOR_ARCHITECTURE%"=="AMD64" set "ARCH=amd64"
if /i "%PROCESSOR_ARCHITEW6432%"=="AMD64" set "ARCH=amd64"

REM Выбери версию (можешь обновлять при желании)
set "PY_VER=3.12.8"
if "%ARCH%"=="amd64" (
  set "PY_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-amd64.exe"
) else (
  set "PY_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%.exe"
)

echo [>] Скачиваем Python %PY_VER% (%ARCH%)...
powershell -NoLogo -NoProfile -Command ^
  "try{ Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%TMP_PY%' -UseBasicParsing }catch{ exit 1 }"
if errorlevel 1 (
  echo [X] Не удалось скачать инсталлятор Python.
  echo Установите Python вручную с https://www.python.org/downloads/ и запустите скрипт снова.
  pause
  exit /b 1
)

echo [>] Тихая установка Python...
REM Пытаемся админ-установку (для всех пользователей). Если прав нет — ставим в профиль пользователя.
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if errorlevel 1 (
  echo [i] Нет прав администратора. Ставим для текущего пользователя...
  start /wait "" "%TMP_PY%" /quiet TargetDir="%LocalAppData%\Programs\Python\Python%PY_VER:.=%" ^
    Include_launcher=1 Include_pip=1 Include_test=0 SimpleInstall=1 PrependPath=1
) else (
  echo [i] Обнаружены права администратора. Ставим для всех пользователей...
  start /wait "" "%TMP_PY%" /quiet InstallAllUsers=1 Include_launcher=1 Include_pip=1 Include_test=0 SimpleInstall=1 PrependPath=1
)

del /q "%TMP_PY%" >nul 2>nul

:PostInstall
REM После установки чаще всего доступен py-лаунчер
where py >nul 2>nul && ( set "PY=py -3" & goto :Run )

REM Если py не появился сразу (PATH ещё не обновился) — пробуем типовые пути
for %%D in (
  "%LocalAppData%\Programs\Python\Python312"
  "%LocalAppData%\Programs\Python\Python311"
  "C:\Program Files\Python312"
  "C:\Program Files\Python311"
  "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312"
  "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311"
) do (
  if exist "%%~D\python.exe" (
    set "PY=%%~D\python.exe"
    goto :Run
  )
)

echo [X] Python установлен, но не найден в PATH прямо сейчас.
echo Закройте это окно и запустите мастер заново, либо запустите вручную:
echo   "%LocalAppData%\Programs\Python\Python312\python.exe" "scripts\first_run_setup.py"
pause
exit /b 1


:Run
echo [OK] Python найден. Запуск мастера...

%PY% -X utf8 "scripts\first_run_setup.py"
if errorlevel 1 (
  echo [X] Скрипт завершился с ошибкой. Проверьте сообщение выше.
  pause
  exit /b 1
)

echo.
pause
exit /b 0
