"""
Где лежат месячные отчёты:

    reports/Ozon/<год>/<Месяц Год>.xlsx       — «reports/Ozon/2026/Октябрь 2026.xlsx»
    reports/WB/<год>/WB <Месяц Год>.xlsx      — «reports/WB/2026/WB Октябрь 2026.xlsx»

До 3.5.0 все отчёты лежали прямо в reports/. Они переезжают на место сами при первом обращении к отчётам
(любая функция ниже вызывает migrate), поэтому старые данные, перенос данных с другого компьютера и распакованный
архив тоже попадают куда нужно. Имена файлов не меняются. Пишут и читают отчёты только через этот модуль.
"""

import re
import shutil
from pathlib import Path
from typing import Optional

REPORTS_DIR = "reports"
OZON_SUBDIR = "Ozon"
WB_SUBDIR = "WB"
WB_PREFIX = "WB "
MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
_YEAR = re.compile(r"\s(\d{4})\.xlsx$", re.IGNORECASE)


def _skip(path: Path) -> bool:
    # Временные и служебные файлы Excel и атомарной записи не трогаем.
    return path.name.startswith(("~", "."))


def _year(name: str) -> Optional[str]:
    match = _YEAR.search(name)
    return match.group(1) if match else None


def reports_root(root) -> Path:
    return Path(root) / REPORTS_DIR


def migrate(root) -> list[Path]:
    """Разложить отчёты по reports/<Ozon|WB>/<год>/. Возвращает новые пути перенесённых файлов.

    Берёт файлы из корня reports/ и из корня reports/Ozon, reports/WB. Файл без года в имени остаётся
    в папке магазина. Если на месте уже есть файл с таким именем, старый не трогается: перезаписывать отчёт нельзя.
    """
    base = reports_root(root)
    if not base.is_dir():
        return []
    moved = []
    for folder in (base, base / OZON_SUBDIR, base / WB_SUBDIR):
        if not folder.is_dir():
            continue
        for path in sorted(folder.glob("*.xlsx")):
            if not path.is_file() or _skip(path):
                continue
            store = base / (WB_SUBDIR if path.name.startswith(WB_PREFIX) else OZON_SUBDIR)
            year = _year(path.name)
            target_dir = store / year if year else store
            if target_dir == folder:
                continue
            target = target_dir / path.name
            if target.exists():
                print(f"⚠️ {path.name} не перенесён: в {target_dir.relative_to(base.parent)} уже есть файл с таким именем")
                continue
            target_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(path), str(target))
            moved.append(target)
    return moved


def ozon_reports_dir(root) -> Path:
    """Папка отчётов Ozon (внутри — папки по годам)."""
    migrate(root)
    return reports_root(root) / OZON_SUBDIR


def wb_reports_dir(root) -> Path:
    """Папка отчётов WB (внутри — папки по годам)."""
    migrate(root)
    return reports_root(root) / WB_SUBDIR


def ozon_report_path(root, year: int, month: int, create: bool = False) -> Path:
    path = ozon_reports_dir(root) / str(year) / f"{MONTHS[month - 1]} {year}.xlsx"
    if create:
        path.parent.mkdir(parents=True, exist_ok=True)
    return path


def wb_report_path(root, year: int, month: int, create: bool = False) -> Path:
    path = wb_reports_dir(root) / str(year) / f"{WB_PREFIX}{MONTHS[month - 1]} {year}.xlsx"
    if create:
        path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _files(folder: Path) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted((p for p in folder.rglob("*.xlsx") if p.is_file() and not _skip(p)), key=lambda p: p.name)


def ozon_report_files(root) -> list[Path]:
    """Все файлы отчётов Ozon во всех годах (без временных)."""
    return _files(ozon_reports_dir(root))


def wb_report_files(root) -> list[Path]:
    return _files(wb_reports_dir(root))
