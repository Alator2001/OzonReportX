#!/bin/bash
# Импорт личных данных из zip на macOS (аналог Import_Data.bat).
# Finder не передаёт перетащенный файл в .command, поэтому архив ищется сам:
# аргумент -> папка проекта -> ~/Downloads (zip или распакованная Safari папка) -> окно выбора файла.
cd "$(dirname "$0")" || exit 1
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

PY="python3"
[ -x ".venv/bin/python" ] && PY=".venv/bin/python"

newest() {
    # Самый свежий OzonReportX_transfer_* (zip или папка) в указанной папке.
    ls -td "$1"/OzonReportX_transfer_* 2>/dev/null | head -n 1
}

ARCHIVE="$1"
[ -z "$ARCHIVE" ] && ARCHIVE="$(newest "$PWD")"
[ -z "$ARCHIVE" ] && ARCHIVE="$(newest "$HOME/Downloads")"
if [ -z "$ARCHIVE" ]; then
    ARCHIVE="$(osascript -e 'POSIX path of (choose file with prompt "Выберите архив OzonReportX_transfer_*.zip" of type {"zip", "public.zip-archive"})' 2>/dev/null)"
fi
if [ -z "$ARCHIVE" ]; then
    echo "[ERROR] Архив не выбран. Положите OzonReportX_transfer_*.zip в папку проекта или в «Загрузки»."
    read -r -p "Нажмите Enter, чтобы закрыть окно..."
    exit 1
fi

echo "Будет импортирован: $ARCHIVE"
echo "Существующие файлы будут сохранены в папку backup_<дата>."
read -r -p "Продолжить? [Y/n]: " ANSWER
case "$ANSWER" in
    ""|[YyДд]*) ;;
    *) echo "Импорт отменён."; read -r -p "Нажмите Enter, чтобы закрыть окно..."; exit 0 ;;
esac

"$PY" -X utf8 "scripts/transfer_data.py" import "$ARCHIVE"
read -r -p "Нажмите Enter, чтобы закрыть окно..."
