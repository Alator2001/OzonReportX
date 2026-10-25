"""
Сборка OzonReportX в приложение: macOS — OzonReportX.app, Windows — папка с OzonReportX.exe.
Результат — архив в dist/: OzonReportX-<версия>-macOS-<arm64|x86_64>.zip или OzonReportX-<версия>-Windows.zip.

Запуск (на той системе, для которой собираем — PyInstaller не собирает под чужую ОС):
    .venv/bin/pip install -r config/requirements-build.txt
    .venv/bin/python packaging/build.py [--version 3.4.0]

Внутри — `flet pack` (PyInstaller + клиент Flet). Сборка всегда «папкой» (--onedir): приложение запускает свои
скрипты этим же файлом, и однофайловая сборка распаковывалась бы во временную папку при каждом запуске скрипта.
"""

import argparse
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PACKAGING = ROOT / "packaging"
DIST = ROOT / "dist"
NAME = "OzonReportX"
BUNDLE_ID = "io.github.alator2001.ozonreportx"
# Нужны только данные пакета (словари, шаблоны), PyInstaller их сам не находит.
COLLECT_DATA = ["pymorphy3", "pymorphy3_dicts_ru", "plotly"]
# Парсер выдачи — по желанию и ставит свой браузер; в сборку не входит.
EXCLUDE = ["playwright", "tkinter", "pytest"]


def read_version() -> str:
    match = re.search(r"Версия \*\*([0-9.]+)\*\*", (ROOT / "README.md").read_text(encoding="utf-8"))
    if not match:
        raise SystemExit("Не нашёл версию в README.md — передайте --version")
    return match.group(1)


def pack(version: str) -> Path:
    # flet pack очищает build/ и dist/ перед сборкой, поэтому version.txt — во временной папке.
    version_file = Path(tempfile.mkdtemp(prefix="ozonreportx-")) / "version.txt"
    version_file.write_text(version, encoding="utf-8")

    is_mac = sys.platform == "darwin"
    icon = PACKAGING / ("icon.icns" if is_mac else "icon.ico")
    # Скрипты приложение запускает как файлы, поэтому кладём исходники; hidden-import — чтобы PyInstaller
    # собрал все их зависимости (часть скриптов приложение не импортирует, а только запускает).
    hidden = [f"scripts.{p.stem}" for p in sorted((ROOT / "scripts").glob("*.py"))]
    extra = ["--onedir", f"--paths={ROOT}"]
    extra += [f"--collect-data={package}" for package in COLLECT_DATA]
    extra += [f"--exclude-module={module}" for module in EXCLUDE]

    cmd = [
        sys.executable, "-c", "import sys; from flet.cli import main; sys.argv[0] = 'flet'; main()", "pack", "launcher.py",
        "--name", NAME, "--icon", str(icon), "--distpath", str(DIST), "-y",
        "--product-name", NAME, "--product-version", version, "--file-version", version,
        "--file-description", "Отчёты и аналитика для продавца Ozon", "--copyright", "OzonReportX",
        "--bundle-id", BUNDLE_ID,
        "--add-data", f"{ROOT / 'scripts'}:scripts", f"{ROOT / 'img'}:img", f"{version_file}:.",
        "--hidden-import", *hidden,
    ]
    if not is_mac:
        cmd.append("--onedir")
    # flet pack принимает флаги PyInstaller только в виде «--pyinstaller-build-args=--флаг=значение».
    cmd += [f"--pyinstaller-build-args={arg}" for arg in extra]
    print("$", " ".join(cmd), flush=True)
    subprocess.run(cmd, cwd=ROOT, check=True)

    if is_mac:
        app = DIST / f"{NAME}.app"
        # Без подписи разработчика Apple приложение не запустится на Apple Silicon вовсе; ad-hoc подпись
        # это снимает (предупреждение Gatekeeper о неизвестном разработчике остаётся).
        subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(app)], check=True)
        archive = DIST / f"{NAME}-{version}-macOS-{platform.machine()}.zip"
        subprocess.run(["ditto", "-c", "-k", "--keepParent", str(app), str(archive)], check=True)
    else:
        folder = DIST / NAME
        archive = DIST / f"{NAME}-{version}-Windows.zip"
        shutil.make_archive(str(archive.with_suffix("")), "zip", DIST, folder.name)
    return archive


def main() -> None:
    parser = argparse.ArgumentParser(description="Сборка OzonReportX для текущей ОС.")
    parser.add_argument("--version", default=None, help="Номер версии; по умолчанию — из README.md")
    args = parser.parse_args()
    archive = pack(args.version or read_version())
    size_mb = archive.stat().st_size / 1024 / 1024
    print(f"✅ Готово: {archive} ({size_mb:.0f} МБ)")


if __name__ == "__main__":
    main()
