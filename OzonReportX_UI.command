#!/bin/bash
# Запуск приложения OzonReportX на macOS (аналог OzonReportX_UI.bat).
cd "$(dirname "$0")" || exit 1
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

VENV_PY=".venv/bin/python"

find_python() {
    for candidate in python3.14 python3; do
        if command -v "$candidate" >/dev/null 2>&1; then
            echo "$candidate"
            return 0
        fi
    done
    return 1
}

if [ ! -x "$VENV_PY" ]; then
    echo "[INFO] Virtual environment not found. Running initial setup first..."
    SYS_PY="$(find_python)"
    if [ -z "$SYS_PY" ]; then
        echo "[ERROR] Python not found! Установите Python 3.14: brew install python@3.14"
        open "https://www.python.org/downloads/"
        read -r -p "Нажмите Enter, чтобы закрыть окно..."
        exit 1
    fi
    "$SYS_PY" -X utf8 "scripts/first_run_setup.py" --setup-only || { read -r -p "Нажмите Enter..."; exit 1; }
fi

if [ ! -x "$VENV_PY" ]; then
    echo "[ERROR] .venv was not created."
    read -r -p "Нажмите Enter, чтобы закрыть окно..."
    exit 1
fi

if ! "$VENV_PY" -X utf8 "scripts/first_run_setup.py" --setup-only; then
    echo "[ERROR] Dependency setup failed."
    read -r -p "Нажмите Enter, чтобы закрыть окно..."
    exit 1
fi

echo "[INFO] Starting Flet UI..."
echo

if ! "$VENV_PY" "app_flet.py"; then
    echo
    echo "[ERROR] UI finished with errors."
    read -r -p "Нажмите Enter, чтобы закрыть окно..."
    exit 1
fi
