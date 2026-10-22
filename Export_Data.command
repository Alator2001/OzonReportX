#!/bin/bash
# Экспорт личных данных в zip на macOS (аналог Export_Data.bat).
cd "$(dirname "$0")" || exit 1
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

PY="python3"
[ -x ".venv/bin/python" ] && PY=".venv/bin/python"

if "$PY" -X utf8 "scripts/transfer_data.py" export; then
    # Показываем готовый архив в Finder.
    ARCHIVE="$(ls -t OzonReportX_transfer_*.zip 2>/dev/null | head -n 1)"
    [ -n "$ARCHIVE" ] && open -R "$ARCHIVE"
fi
read -r -p "Нажмите Enter, чтобы закрыть окно..."
