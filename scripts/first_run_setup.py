import os
import sys
import subprocess
from pathlib import Path


def print_step(title: str):
    print(f"\n=== {title} ===")


def run(cmd, cwd=None, quiet=False):
    if quiet:
        result = subprocess.run(cmd, cwd=cwd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        result = subprocess.run(cmd, cwd=cwd)
    if result.returncode != 0:
        raise RuntimeError(f"Команда завершилась с кодом {result.returncode}: {' '.join(map(str, cmd))}")


def ensure_venv(repo_root: Path) -> tuple[Path, bool]:
    venv_dir = repo_root / ".venv"
    created = False
    if not venv_dir.exists():
        print_step("Создание виртуального окружения (.venv)")
        run([sys.executable, "-m", "venv", str(venv_dir)])
        created = True
    venv_python = venv_dir / "Scripts" / "python.exe"
    if not venv_python.exists():
        # На всякий случай поддержим unix-пути
        venv_python = venv_dir / "bin" / "python"
    if not venv_python.exists():
        raise RuntimeError("Не найден исполняемый файл Python в .venv")
    return venv_python, created


def ensure_deps(venv_python: Path, repo_root: Path):
    print_step("Обновление pip и установка зависимостей")
    run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip"], cwd=repo_root, quiet=True)
    config_dir = Path(__file__).resolve().parent
    req = config_dir / "requirements.txt"
    if req.exists():
        run([str(venv_python), "-m", "pip", "install", "-r", str(req)], cwd=repo_root, quiet=True)
    else:
        run([str(venv_python), "-m", "pip", "install",
             "requests", "pandas", "openpyxl", "python-dateutil", "python-dotenv"], cwd=repo_root, quiet=True)


def prompt_yes_no(message: str, default_yes: bool = True) -> bool:
    suffix = "[Y/n]" if default_yes else "[y/N]"
    while True:
        answer = input(f"{message} {suffix} ").strip().lower()
        if not answer:
            return default_yes
        if answer in ("y", "yes", "д", "да"):
            return True
        if answer in ("n", "no", "н", "нет"):
            return False
        print("Пожалуйста, ответьте 'y' или 'n'.")


def ensure_env(repo_root: Path) -> bool:
    env_path = repo_root / ".env"
    if env_path.exists():
        return False
    print_step("Создание .env")
    client_id = input("Введите OZON_CLIENT_ID: ").strip()
    api_key = input("Введите OZON_API_KEY: ").strip()
    env_path.write_text(f"OZON_CLIENT_ID={client_id}\nOZON_API_KEY={api_key}\n", encoding="utf-8")
    print("Файл .env создан")
    return True


def ensure_costs(repo_root: Path) -> bool:
    costs_xlsx = repo_root / "costs.xlsx"
    costs_csv = repo_root / "costs.csv"
    created = False
    if costs_xlsx.exists() or costs_csv.exists():
        print_step("Файл себестоимости найден")
    else:
        print_step("Создание шаблона себестоимости costs.xlsx")
        try:
            import pandas as pd
            df = pd.DataFrame(columns=["артикул", "себестоимость"])
            df.to_excel(costs_xlsx, index=False)
            created = True
            print("Создан costs.xlsx. Заполните артикулы и себестоимость.")
        except Exception as e:
            # Резервный вариант — CSV, если нет pandas/openpyxl
            costs_csv.write_text("артикул,себестоимость\n", encoding="utf-8")
            print(f"Не удалось создать Excel ({e}). Создан резервный costs.csv.")
    # Диалог по себестоимости — только если файл был создан сейчас
    if created:
        print_step("Заполнение себестоимости")
        print("Укажите себестоимость для своих артикулов в файле costs.xlsx (или costs.csv).")
        if prompt_yes_no("Открыть файл себестоимости сейчас?", default_yes=True):
            try:
                target = str(costs_xlsx if costs_xlsx.exists() else costs_csv)
                if os.name == 'nt':
                    os.startfile(target)  # type: ignore[attr-defined]
                elif sys.platform == 'darwin':
                    run(["open", target])
                else:
                    run(["xdg-open", target])
            except Exception as e:
                print(f"Не удалось открыть файл автоматически: {e}")
    return created


def ensure_reports_dir(repo_root: Path):
    reports = repo_root / "reports"
    reports.mkdir(parents=True, exist_ok=True)


def run_report(venv_python: Path, repo_root: Path):
    print_step("Запуск формирования отчёта")
    main_script = repo_root / "scripts" / "Monthly_sales_report.py"
    run([str(venv_python), str(main_script)], cwd=repo_root)

def select_menu_option():
    print_step("Меню выбора отчёта")
    print("1. Месячный отчёт по продажам")
    print("2. ABC&XYZ-анализ")
    print("3. Выход")   
    while True:
        choice = input("Выберите опцию (1-3): ").strip()
        if choice in ("1", "2", "3"):
            return choice
        print("Пожалуйста, выберите корректную опцию (1, 2 или 3).")


def main():
    choice = select_menu_option()
    if choice == "1":
        print_step("Выбран Месячный отчёт по продажам.")
        repo_root = Path(__file__).resolve().parent.parent
        print_step("Мастер настройки и запуска генерации отчёта Ozon")
        try:
            venv_python, venv_created = ensure_venv(repo_root)
            ensure_deps(venv_python, repo_root)
            env_created = ensure_env(repo_root)
            costs_created = ensure_costs(repo_root)
            ensure_reports_dir(repo_root)
            # Спрашиваем разрешение на запуск только при самом первом конфигурировании
            if venv_created or env_created or costs_created:
                if prompt_yes_no("Можно начинать формирование отчёта?", default_yes=True):
                    run_report(venv_python, repo_root)
                else:
                    print("Окей, запуск отчёта отменён. Вы можете запустить позже: run.bat или python config/first_run_setup.py")
            else:
                run_report(venv_python, repo_root)
        except KeyboardInterrupt:
            print("\nОперация прервана пользователем.")
        except Exception as e:
            print(f"\nОшибка: {e}")
            sys.exit(1)
    elif choice == "2":
        print("Выбран ABC&XYZ-анализ.")
        print("Пока в разработке. Но скоро появится")
        return
    if choice == "3":
        print("Выход из программы.")
        return
    


if __name__ == "__main__":
    main()


