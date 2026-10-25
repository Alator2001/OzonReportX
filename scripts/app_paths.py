"""
Где лежит код программы и где — данные пользователя.

Из исходников (python app_flet.py) это одна папка — корень репозитория, как и раньше.
В собранном приложении (PyInstaller: OzonReportX.app / OzonReportX.exe) код внутри сборки только для чтения,
поэтому отчёты, costs.xlsx, .env и настройки живут в «Документы/OzonReportX» или в папке, выбранной
в «Настройках» → «Папка с данными». Выбор хранится вне папки данных — в папке настроек программы
(macOS: ~/Library/Application Support/OzonReportX, Windows: %APPDATA%\\OzonReportX).
Переменная OZONREPORTX_HOME важнее всего.

Скрипты приложение запускает отдельными процессами. В сборке нет отдельного python, поэтому
script_command() запускает ту же сборку с путём к скрипту, а launcher.py выполняет этот скрипт.
"""

import os
import sys
from pathlib import Path
from typing import Optional

FROZEN = bool(getattr(sys, "frozen", False))
CODE_ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
SCRIPTS_DIR = CODE_ROOT / "scripts"
DEFAULT_DATA_ROOT = Path.home() / "Documents" / "OzonReportX"


def _config_dir() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "OzonReportX"
    if os.name == "nt":
        return Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming") / "OzonReportX"
    return Path.home() / ".config" / "ozonreportx"


DATA_FOLDER_SETTING = _config_dir() / "data_folder.txt"


def chosen_data_folder() -> Optional[Path]:
    """Папка, выбранная в «Настройках»; None — не выбрана или её больше нет (например, диск отключён)."""
    try:
        text = DATA_FOLDER_SETTING.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    folder = Path(text).expanduser() if text else None
    return folder if folder and folder.is_dir() else None


def save_data_folder(folder: Optional[Path]) -> None:
    """Запомнить папку с данными (None — вернуть «Документы/OzonReportX»). Действует со следующего запуска."""
    if folder is None:
        DATA_FOLDER_SETTING.unlink(missing_ok=True)
        return
    DATA_FOLDER_SETTING.parent.mkdir(parents=True, exist_ok=True)
    DATA_FOLDER_SETTING.write_text(str(Path(folder).expanduser().resolve()), encoding="utf-8")


def _data_root() -> tuple[Path, str]:
    """(папка, откуда взята): env — OZONREPORTX_HOME, chosen — выбрана в «Настройках», default — по умолчанию."""
    override = os.environ.get("OZONREPORTX_HOME")
    if override:
        return Path(override).expanduser(), "env"
    if not FROZEN:
        return CODE_ROOT, "default"
    chosen = chosen_data_folder()
    if chosen:
        return chosen, "chosen"
    return DEFAULT_DATA_ROOT, "default"


DATA_ROOT, DATA_ROOT_SOURCE = _data_root()
ENV_FILE = DATA_ROOT / ".env"


def python_executable() -> Path:
    """Чем запускать скрипты: сборка сама себя, иначе python из .venv, иначе текущий."""
    if FROZEN:
        return Path(sys.executable)
    for candidate in (CODE_ROOT / ".venv" / "Scripts" / "python.exe", CODE_ROOT / ".venv" / "bin" / "python"):
        if candidate.exists():
            return candidate
    return Path(sys.executable)


def script_command(name: str, *args: str) -> list[str]:
    return [str(python_executable()), str(SCRIPTS_DIR / name), *args]
