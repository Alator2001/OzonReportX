"""
Перенос локальных данных OzonReportX на другой компьютер.

В git не хранятся ключи API, себестоимость, настройки и готовые отчёты.
Этот скрипт упаковывает их в один zip-архив и распаковывает на новом компьютере.

    python scripts/transfer_data.py export            -> OzonReportX_transfer_YYYYMMDD_HHMM.zip
    python scripts/transfer_data.py import <archive>  -> восстановление в папку проекта
"""
import argparse
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

# Файлы и папки, которые нужны программе, но исключены из git.
TRANSFER_FILES = [".env", "costs.xlsx", "wb_costs.xlsx", "margin_settings.json"]
TRANSFER_DIRS = ["reports", "wb reports", "ABC&XYZ reports", "balance reports", "stocks reports"]

SKIP_PREFIXES = ("~$", "~tmp_", ".")


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
            zf.write(path, path.relative_to(root).as_posix())
            count += 1
    print(f"Упаковано файлов: {count}")
    print(f"Архив: {archive}")
    print("ВНИМАНИЕ: архив содержит ключи API (.env). Не публикуйте его и удалите после переноса.")
    return archive


def import_data(root: Path, archive: Path) -> None:
    if not archive.is_file():
        raise SystemExit(f"Архив не найден: {archive}")
    backup_dir = root / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    restored = backed_up = 0
    with zipfile.ZipFile(archive) as zf:
        for member in zf.infolist():
            if member.is_dir():
                continue
            target = (root / member.filename).resolve()
            if root.resolve() not in target.parents:
                print(f"Пропущен небезопасный путь: {member.filename}")
                continue
            if target.exists():
                backup_path = backup_dir / member.filename
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target, backup_path)
                backed_up += 1
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            restored += 1
    print(f"Восстановлено файлов: {restored}")
    if backed_up:
        print(f"Существующие файлы ({backed_up}) сохранены в {backup_dir.name}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Перенос данных OzonReportX между компьютерами")
    sub = parser.add_subparsers(dest="command", required=True)
    exp = sub.add_parser("export", help="Упаковать ключи, себестоимость, настройки и отчёты")
    exp.add_argument("-o", "--output", type=Path, help="Путь к создаваемому архиву")
    imp = sub.add_parser("import", help="Восстановить данные из архива")
    imp.add_argument("archive", type=Path, help="Путь к архиву OzonReportX_transfer_*.zip")
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
