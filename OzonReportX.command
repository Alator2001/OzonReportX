#!/bin/bash
# Консольное меню OzonReportX на macOS (аналог OzonReportX.bat).
cd "$(dirname "$0")" || exit 1
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

PY=""
for candidate in python3.14 python3; do
    if command -v "$candidate" >/dev/null 2>&1; then
        PY="$candidate"
        break
    fi
done

"${PY:-python3}" -X utf8 "config/show_banner.py" 2>/dev/null

if [ -z "$PY" ]; then
    echo "[ERROR] Python not found! Установите Python 3.14: brew install python@3.14"
    open "https://www.python.org/downloads/"
    read -r -p "Нажмите Enter, чтобы закрыть окно..."
    exit 1
fi

echo "[INFO] Starting setup..."
echo

if ! "$PY" -X utf8 "scripts/first_run_setup.py" "$@"; then
    echo "[ERROR] Script finished with errors."
    read -r -p "Нажмите Enter, чтобы закрыть окно..."
    exit 1
fi

read -r -p "Нажмите Enter, чтобы закрыть окно..."
