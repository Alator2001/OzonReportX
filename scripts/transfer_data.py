"""
Перенос локальных данных OzonReportX на другой компьютер.

В git не хранятся ключи API, себестоимость, настройки и готовые отчёты.
Этот скрипт упаковывает их в один zip-архив и распаковывает на новом компьютере.

    python scripts/transfer_data.py export            -> OzonReportX_transfer_YYYYMMDD_HHMM.zip
    python scripts/transfer_data.py import <archive>  -> восстановление в папку проекта

Импорт принимает и zip, и уже распакованную папку (Safari на macOS распаковывает zip сам).
"""
import argparse
import shutil
import sys
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

# Файлы и папки, которые нужны программе, но исключены из git.
TRANSFER_FILES = [".env", "costs.xlsx", "wb_costs.xlsx", "margin_settings.json", "search_settings.json"]
TRANSFER_DIRS = ["reports", "wb reports", "ABC&XYZ reports", "balance reports", "stocks reports", "search data"]

SKIP_PREFIXES = ("~$", "~tmp_", ".")
# Служебные файлы Finder и архиватора macOS: в данные программы они не входят.
MAC_JUNK_DIRS = {"__MACOSX"}
MAC_JUNK_FILES = {".DS_Store"}


def _nfc(name: str) -> str:
    """macOS может хранить кириллицу в форме NFD (й = и + ˘); программа ищет имена в NFC."""
    return unicodedata.normalize("NFC", name)


def _clean_member_name(name: str) -> str | None:
    """Путь записи архива в виде a/b/c (NFC) или None, если это мусор macOS."""
    parts = [part for part in _nfc(name).replace("\\", "/").split("/") if part]
    if not parts or any(part in MAC_JUNK_DIRS for part in parts):
        return None
    if parts[-1] in MAC_JUNK_FILES or parts[-1].startswith("._"):
        return None
    return "/".join(parts)


def _strip_wrapper_folder(names: list[str]) -> list[str]:
    """Если архив пересжали через Finder, всё лежит внутри одной лишней папки — снимаем её."""
    known = set(TRANSFER_FILES) | set(TRANSFER_DIRS)
    tops = {name.split("/", 1)[0] for name in names}
    if len(tops) != 1 or tops <= known or not all("/" in name for name in names):
        return names
    return [name.split("/", 1)[1] for name in names]


def _iter_items(root: Path):
    for name in TRANSFER_FILES:
        path = root / name
        if path.is_file():
            yield path
    for name in TRANSFER_DIRS:
        folder = root / name
        if not folder.is_dir():
            continue
        for path in sorted(folder.rglob("*")):
            if path.is_file() and not path.name.startswith(SKIP_PREFIXES):
                yield path


def export_data(root: Path, output: Path | None) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    archive = output or root / f"OzonReportX_transfer_{stamp}.zip"
    count = 0
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in _iter_items(root):
            zf.write(path, _nfc(path.relative_to(root).as_posix()))
            count += 1
    print(f"Упаковано файлов: {count}")
    print(f"Архив: {archive}")
    print("ВНИМАНИЕ: архив содержит ключи API (.env). Не публикуйте его и удалите после переноса.")
    return archive


def _restore(root: Path, entries: list[tuple[str, object]]) -> tuple[int, int, Path]:
    """entries: (путь внутри архива, функция, открывающая содержимое на чтение)."""
    backup_dir = root / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    names = _strip_wrapper_folder([name for name, _ in entries])
    restored = backed_up = 0
    for name, (_, open_source) in zip(names, entries):
        target = (root / name).resolve()
        if root.resolve() not in target.parents:
            print(f"Пропущен небезопасный путь: {name}")
            continue
        if target.exists():
            backup_path = backup_dir / name
            backup_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(target, backup_path)
            backed_up += 1
        target.parent.mkdir(parents=True, exist_ok=True)
        with open_source() as src, open(target, "wb") as dst:
            shutil.copyfileobj(src, dst)
        restored += 1
    return restored, backed_up, backup_dir


def import_data(root: Path, archive: Path) -> None:
    if archive.is_dir():
        entries = []
        for path in sorted(archive.rglob("*")):
            name = _clean_member_name(path.relative_to(archive).as_posix()) if path.is_file() else None
            if name:
                entries.append((name, lambda path=path: open(path, "rb")))
        restored, backed_up, backup_dir = _restore(root, entries)
    elif archive.is_file():
        with zipfile.ZipFile(archive) as zf:
            entries = []
            for member in zf.infolist():
                name = None if member.is_dir() else _clean_member_name(member.filename)
                if name:
                    entries.append((name, lambda member=member: zf.open(member)))
            restored, backed_up, backup_dir = _restore(root, entries)
    else:
        raise SystemExit(f"Архив не найден: {archive}")
    print(f"Восстановлено файлов: {restored}")
    if backed_up:
        print(f"Существующие файлы ({backed_up}) сохранены в {backup_dir.name}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Перенос данных OzonReportX между компьютерами")
    sub = parser.add_subparsers(dest="command", required=True)
    exp = sub.add_parser("export", help="Упаковать ключи, себестоимость, настройки и отчёты")
    exp.add_argument("-o", "--output", type=Path, help="Путь к создаваемому архиву")
    imp = sub.add_parser("import", help="Восстановить данные из архива")
    imp.add_argument("archive", type=Path, help="Путь к архиву OzonReportX_transfer_*.zip или к распакованной папке")
    args = parser.parse_args()

    if args.command == "export":
        export_data(REPO_ROOT, args.output)
    else:
        import_data(REPO_ROOT, args.archive)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
