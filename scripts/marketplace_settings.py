"""Shared first-run credentials setup for the desktop and console launchers."""
import argparse
from getpass import getpass
import os
from pathlib import Path

from dotenv import dotenv_values, set_key

try:
    from scripts.file_io import atomic_output_path
except ModuleNotFoundError:
    from file_io import atomic_output_path


FIELDS = (
    ("OZON_CLIENT_ID", "Ozon Client ID", False),
    ("OZON_API_KEY", "Ozon API-ключ", True),
    ("OZON_PERF_CLIENT_ID", "Ozon Performance Client ID (необязательно)", False),
    ("OZON_PERF_API_KEY", "Ozon Performance Client Secret (необязательно)", True),
    ("WB_API_TOKEN", "Wildberries API-токен", True),
)
SETUP_KEY = "MARKETPLACE_SETUP_VERSION"
SETUP_VERSION = "1"
HELP = (
    "Ozon: Настройки → API-ключи — Client ID и API-ключ. "
    "Для рекламных отчётов: Продвижение → API — отдельные Client ID и Client Secret.\n"
    "WB: Интеграции по API → Создать токен → Ручная интеграция. "
    "Для локального приложения используйте персональный токен. "
    "Выбирайте категории только для нужных операций; для чтения данных достаточно доступа «Только чтение». "
    "Отдельный Client ID для WB не нужен.\n"
    "Для бизнес-сводки WB включите категорию «Финансы». "
    "Сохранение ключей не проверяет подключение к API."
)


def read_settings(root):
    return dict(dotenv_values(Path(root) / ".env", interpolate=False))


def needs_setup(root):
    return read_settings(root).get(SETUP_KEY) != SETUP_VERSION


def save_settings(root, updates):
    """Preserve unrelated settings and leave the original intact on write failure."""
    path = Path(root) / ".env"
    values = read_settings(root)
    cleaned = {}
    for key, _label, _secret in FIELDS:
        value = updates.get(key, values.get(key, "")) or ""
        value = value.strip()
        if any(c in value for c in ("\r", "\n", "\x00")):
            raise ValueError("Каждое значение должно занимать одну строку.")
        cleaned[key] = value
    for left, right, label in (
        ("OZON_CLIENT_ID", "OZON_API_KEY", "Ozon"),
        ("OZON_PERF_CLIENT_ID", "OZON_PERF_API_KEY", "Ozon Performance"),
    ):
        if bool(cleaned[left]) != bool(cleaned[right]):
            raise ValueError(f"{label}: заполните оба поля или оставьте оба пустыми.")
    cleaned[SETUP_KEY] = SETUP_VERSION
    with atomic_output_path(path) as temporary:
        temporary.write_text(path.read_text(encoding="utf-8") if path.exists() else "", encoding="utf-8")
        for key, value in cleaned.items():
            set_key(str(temporary), key, value, quote_mode="always")
    for key, value in cleaned.items():
        os.environ[key] = value


def configure_console(root, force=False):
    if not force and not needs_setup(root):
        return
    print("\nНастройка Ozon и Wildberries\n" + HELP)
    print("Enter — сохранить текущее значение или пропустить. '-' — очистить поле.")
    while True:
        current = read_settings(root)
        updates = {}
        for key, label, secret in FIELDS:
            state = " [уже задано]" if current.get(key) else ""
            value = (getpass if secret else input)(f"{label}{state}: ").strip()
            updates[key] = "" if value == "-" else value or current.get(key, "")
        try:
            save_settings(root, updates)
        except ValueError as exc:
            print(str(exc))
            continue
        print("Настройки сохранены в .env. Изменить их можно в разделе «Настройки».")
        return


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    configure_console(args.root, args.force)
