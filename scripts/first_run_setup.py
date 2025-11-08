import os
import sys
import subprocess
from pathlib import Path


def print_step(title: str):
    print(f"\n=== {title} ===")

def ensure_auto_update_package(venv_python: Path, repo_root: Path):
    """Проверяет и устанавливает необходимые пакеты для обновления"""
    required_packages = ['requests', 'packaging']
    
    for package in required_packages:
        try:
            result = subprocess.run(
                [str(venv_python), "-c", f"import {package}"],
                cwd=repo_root,
                capture_output=True,
                timeout=5
            )
            if result.returncode != 0:
                subprocess.check_call(
                    [str(venv_python), "-m", "pip", "install", package],
                    cwd=repo_root,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
        except:
            return False
    return True


def check_for_updates(venv_python: Path, repo_root: Path):
    """Проверка и установка обновлений"""
    print_step("Проверка обновлений")
    
    if not ensure_auto_update_package(venv_python, repo_root):
        print("⚠ Не удалось установить необходимые пакеты")
        return
    
    try:
        auto_update_file = repo_root / "scripts" / "_auto_update.py"
        
        if not auto_update_file.exists():
            print("⚠ Файл scripts/_auto_update.py не найден")
            return
        
        result = subprocess.run(
            [str(venv_python), str(auto_update_file)],
            cwd=repo_root,
            timeout=60
        )
        
        if result.returncode == 0:
            print("✓ Проверка обновлений завершена")
            
    except subprocess.TimeoutExpired:
        print("⚠ Превышено время ожидания")
    except Exception as e:
        print(f"⚠ Ошибка: {e}")

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
    """Проверяет наличие файла себестоимости"""
    costs_xlsx = repo_root / "costs.xlsx"
    costs_csv = repo_root / "costs.csv"

    if costs_xlsx.exists() or costs_csv.exists():
        print_step("Файл себестоимости найден")
        return False
    
    print_step("Создание шаблона себестоимости costs.xlsx")
    
    try:
        import subprocess
        result = subprocess.run(
            [str(venv_python), "-c", 
             "import pandas as pd; "
             "df = pd.DataFrame(columns=['артикул', 'себестоимость']); "
             f"df.to_excel('{costs_xlsx}', index=False)"],
            cwd=repo_root,
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            print("Создан costs.xlsx. Заполните артикулы и себестоимость.")
            created = True
        else:
            raise Exception("pandas не установлен")
            
    except Exception as e:
        # Резервный вариант — CSV
        costs_csv.write_text("артикул,себестоимость\n", encoding="utf-8")
        print(f"Не удалось создать Excel ({e}). Создан резервный costs.csv.")
        created = True
    
    # Диалог по себестоимости
    if created:
        print_step("Заполнение себестоимости")
        print("Укажите себестоимость для своих артикулов в файле costs.xlsx (или costs.csv).")
        if prompt_yes_no("Открыть файл себестоимости сейчас?", default_yes=True):
            try:
                target = str(costs_xlsx if costs_xlsx.exists() else costs_csv)
                if os.name == 'nt':
                    os.startfile(target)
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
    repo_root = Path(__file__).resolve().parent.parent
    print_step("Мастер настройки и запуска генерации отчёта Ozon")
    
    venv_python, venv_created = ensure_venv(repo_root)
    ensure_deps(venv_python, repo_root)
    
    # Проверка обновлений сразу после установки зависимостей
    check_for_updates(venv_python, repo_root)
    
    choice = select_menu_option()
    if choice == "1":
        print_step("Выбран Месячный отчёт по продажам.")
        print_step("Мастер настройки и запуска генерации отчёта Ozon")
        try:
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


