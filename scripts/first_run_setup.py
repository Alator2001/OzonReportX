import os
import sys
import subprocess
from pathlib import Path
import argparse
# Импорт утилит как локального модуля при запуске по пути (scripts/first_run_setup.py)
try:
    from scripts.utils import print_step, prompt_yes_no, set_prompt_force  # type: ignore
except ModuleNotFoundError:
    sys.path.append(str(Path(__file__).resolve().parent))
    from utils import print_step, prompt_yes_no, set_prompt_force  # type: ignore

 

def ensure_auto_update_package(venv_python: Path, repo_root: Path):
    """Требуемые пакеты для автообновления уже ставятся через requirements.txt в ensure_deps."""
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
    # Определим каталог виртуального окружения и файл-маркер
    venv_dir = Path(venv_python).resolve().parent.parent
    bootstrap_marker = venv_dir / ".bootstrap_done"
    if bootstrap_marker.exists():
        print("Зависимости уже установлены (найден .bootstrap_done).")
        return
    # Обновляем pip и ставим зависимости одним вызовом
    run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip"], cwd=repo_root, quiet=True)
    config_dir = Path(__file__).resolve().parent
    req = config_dir / "requirements.txt"
    if req.exists():
        run([str(venv_python), "-m", "pip", "install", "-r", str(req)], cwd=repo_root, quiet=True)
    else:
        # Резервный список, если requirements.txt отсутствует
        run([str(venv_python), "-m", "pip", "install",
             "requests", "pandas", "openpyxl", "python-dateutil", "python-dotenv", "packaging"], cwd=repo_root, quiet=True)
    # Создаём маркер успешной установки
    try:
        bootstrap_marker.write_text("ok", encoding="utf-8")
    except Exception:
        pass

 


def ensure_env(repo_root: Path) -> bool:
    env_path = repo_root / ".env"
    if env_path.exists():
        return False
    print_step("Создание .env")
    client_id = input("Введите OZON_CLIENT_ID: ").strip()
    api_key = input("Введите OZON_API_KEY: ").strip()
    
    env_content = f"OZON_CLIENT_ID={client_id}\nOZON_API_KEY={api_key}\n"
    
    # Опционально: Performance API для автоматического получения затрат на рекламу
    if prompt_yes_no("Добавить Performance API credentials для автоматического получения затрат на рекламу? (опционально)", default_yes=False):
        print("\nДля получения данных о рекламных кампаниях нужны отдельные ключи из раздела 'Продвижение' → 'API' в кабинете Ozon.")
        perf_client_id = input("Введите OZON_PERF_CLIENT_ID (или Enter для пропуска): ").strip()
        perf_api_key = input("Введите OZON_PERF_API_KEY (или Enter для пропуска): ").strip()
        if perf_client_id and perf_api_key:
            env_content += f"OZON_PERF_CLIENT_ID={perf_client_id}\nOZON_PERF_API_KEY={perf_api_key}\n"
            print("✅ Performance API credentials добавлены")
        else:
            print("ℹ️ Performance API credentials не добавлены (можно добавить позже в .env)")
    
    env_path.write_text(env_content, encoding="utf-8")
    print("Файл .env создан")
    return True


def ensure_costs(venv_python, repo_root: Path) -> bool:
    """Проверяет наличие файла себестоимости"""
    costs_xlsx = repo_root / "costs.xlsx"

    if costs_xlsx.exists():
        print_step("Файл себестоимости найден")
        return False
    
    print_step("Создание шаблона себестоимости costs.xlsx")
    

    import subprocess
    # Путь для подстановки в однострочную команду Python
    _path = str(costs_xlsx).replace("\\", "\\\\")
    create_cmd = (
        "import sys; "
        "path=r\"" + _path + "\"; "
        "try:\n"
        "    import pandas as pd\n"
        "    df = pd.DataFrame(columns=['артикул', 'себестоимость'])\n"
        "    df.to_excel(path, index=False)\n"
        "except Exception:\n"
        "    from openpyxl import Workbook\n"
        "    wb = Workbook()\n"
        "    ws = wb.active\n"
        "    ws.append(['артикул', 'себестоимость'])\n"
        "    wb.save(path)\n"
    )
    result = subprocess.run(
        [str(venv_python), "-c", create_cmd],
        cwd=repo_root,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        timeout=10
    )
    
    if result.returncode == 0:
        print("Создан costs.xlsx. Заполните артикулы и себестоимость.")
        created = True
    else:
        raise Exception("Не удалось создать файл себестоимости (ни через pandas, ни через openpyxl)")
            
    
    # Диалог по себестоимости
    if created:
        print_step("Заполнение себестоимости")
        print("Укажите себестоимость для своих артикулов в файле costs.xlsx")
        if prompt_yes_no("Открыть файл себестоимости сейчас?", default_yes=True):
            try:
                target = str(costs_xlsx)
                if costs_xlsx.exists():
                    if os.name == 'nt':
                        os.startfile(target)
                    elif sys.platform == 'darwin':
                        run(["open", target])
                    else:
                        run(["xdg-open", target])
                else:
                    print("Файл не найден.")
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


# Названия месяцев для имён файлов отчётов («Месяц Год.xlsx»), как в Monthly_sales_report
MONTHS_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]


def ask_abc_xyz_date_range():
    """Запрашивает у пользователя диапазон месяцев для ABC&XYZ-анализа. Возвращает (from_month, from_year, to_month, to_year)."""
    print("Укажите диапазон месяцев для анализа (по одному месяцу — введите его дважды).")
    while True:
        try:
            from_part = input("Месяц и год начала (например 10 2025): ").strip().split()
            to_part = input("Месяц и год конца (например 12 2025): ").strip().split()
            if len(from_part) >= 2 and len(to_part) >= 2:
                from_month = int(from_part[0])
                from_year = int(from_part[1])
                to_month = int(to_part[0])
                to_year = int(to_part[1])
                if not (1 <= from_month <= 12 and 2000 <= from_year <= 2100):
                    print("Некорректное начало: месяц 1–12, год 2000–2100.")
                    continue
                if not (1 <= to_month <= 12 and 2000 <= to_year <= 2100):
                    print("Некорректный конец: месяц 1–12, год 2000–2100.")
                    continue
                if (from_year, from_month) <= (to_year, to_month):
                    return from_month, from_year, to_month, to_year
                print("Начало периода должно быть не позже конца.")
            else:
                print("Введите два числа через пробел: месяц и год (например 10 2025).")
        except ValueError:
            print("Некорректный ввод. Введите месяц и год числами через пробел.")


def run_abc_xyz(
    venv_python: Path,
    repo_root: Path,
    from_month: int,
    from_year: int,
    to_month: int,
    to_year: int,
):
    """
    Проверяет наличие отчётов за запрошенный диапазон в reports; недостающие генерирует.
    Затем объединяет заказы из отчётов за этот диапазон в папку «ABC&XYZ reports».
    """
    ensure_reports_dir(repo_root)
    reports_dir = repo_root / "reports"
    main_script = repo_root / "scripts" / "Monthly_sales_report.py"
    abc_script = repo_root / "scripts" / "ABC_XYZ_analytics_report.py"

    # Список (year, month) от начала до конца включительно
    start_ym = from_year * 12 + (from_month - 1)
    end_ym = to_year * 12 + (to_month - 1)
    months_to_need = []
    for i in range(start_ym, end_ym + 1):
        y, m = i // 12, (i % 12) + 1
        months_to_need.append((y, m))

    missing = []
    for y, m in months_to_need:
        fname = f"{MONTHS_RU[m - 1]} {y}.xlsx"
        if not (reports_dir / fname).exists():
            missing.append((y, m))

    if missing:
        print_step("Генерация недостающих месячных отчётов")
        for y, m in missing:
            label = f"{MONTHS_RU[m - 1]} {y}"
            print(f"  Формируется отчёт за {label}...")
            run(
                [str(venv_python), str(main_script), "--month", str(m), "--year", str(y)],
                cwd=repo_root,
            )

    print_step("ABC&XYZ-анализ: объединение заказов из отчётов за выбранный период")
    run(
        [
            str(venv_python), str(abc_script),
            "-i", str(reports_dir),
            "--output_dir", "ABC&XYZ reports",
            "--from-month", str(from_month), "--from-year", str(from_year),
            "--to-month", str(to_month), "--to-year", str(to_year),
        ],
        cwd=repo_root,
    )


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
    
    parser = argparse.ArgumentParser(add_help=False)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--yes", action="store_true", help="Автоматически отвечать Да на все вопросы")
    group.add_argument("--no", action="store_true", help="Автоматически отвечать Нет на все вопросы")
    args, _unknown = parser.parse_known_args()
    if args.yes:
        set_prompt_force(True)
    elif args.no:
        set_prompt_force(False)
    else:
        set_prompt_force(None)
    
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
            costs_created = ensure_costs(venv_python, repo_root)
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
        print_step("Выбран ABC&XYZ-анализ.")
        try:
            ensure_env(repo_root)
            ensure_costs(venv_python, repo_root)
            ensure_reports_dir(repo_root)
            from_month, from_year, to_month, to_year = ask_abc_xyz_date_range()
            run_abc_xyz(venv_python, repo_root, from_month, from_year, to_month, to_year)
        except KeyboardInterrupt:
            print("\nОперация прервана пользователем.")
        except Exception as e:
            print(f"\nОшибка: {e}")
            sys.exit(1)
        return
    if choice == "3":
        print("Выход из программы.")
        return
    


if __name__ == "__main__":
    main()