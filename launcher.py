"""
Точка входа собранного приложения (OzonReportX.app / OzonReportX.exe), см. packaging/build.py.

Без аргументов открывает окно программы. С путём к скрипту (`OzonReportX scripts/Monthly_sales_report.py --month 9 …`)
выполняет этот скрипт так, как его выполнил бы python: в сборке нет отдельного интерпретатора,
и приложение запускает свои скрипты этим же файлом (app_paths.script_command).
"""

import io
import multiprocessing
import os
import runpy
import sys
from pathlib import Path


def _utf8_stream(stream, fd):
    """В оконной сборке Windows sys.stdout бывает None, хотя родитель передал трубу."""
    if stream is None:
        try:
            stream = io.TextIOWrapper(os.fdopen(fd, "wb", buffering=0), encoding="utf-8", errors="replace")
        except OSError:
            stream = open(os.devnull, "w", encoding="utf-8")
    try:
        stream.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except (AttributeError, ValueError):
        pass
    return stream


def run_script(path: str, args: list[str]) -> None:
    sys.stdout = _utf8_stream(sys.stdout, 1)
    sys.stderr = _utf8_stream(sys.stderr, 2)
    script = Path(path).resolve()
    sys.argv = [str(script), *args]
    sys.path[:0] = [str(script.parent), str(script.parent.parent)]
    runpy.run_path(str(script), run_name="__main__")


def run_app() -> None:
    from scripts.app_paths import DATA_ROOT

    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    os.chdir(DATA_ROOT)
    # Скрипты, запущенные из окна, работают с той же папкой, даже если в «Настройках» уже выбрали другую
    # (она действует со следующего запуска программы).
    os.environ["OZONREPORTX_HOME"] = str(DATA_ROOT)
    import flet as ft
    import app_flet

    ft.run(app_flet.main)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    if len(sys.argv) > 1 and sys.argv[1].endswith(".py"):
        run_script(sys.argv[1], sys.argv[2:])
    else:
        run_app()
