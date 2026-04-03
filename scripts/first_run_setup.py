import os
import sys
import subprocess
from pathlib import Path
import argparse

try:
    from scripts.utils import print_step, prompt_yes_no, set_prompt_force, log_verbose, VERBOSE  # type: ignore
except ModuleNotFoundError:
    sys.path.append(str(Path(__file__).resolve().parent))
    from utils import print_step, prompt_yes_no, set_prompt_force, log_verbose, VERBOSE  # type: ignore


def ensure_auto_update_package(venv_python: Path, repo_root: Path):
    """Зависимости автообновления ставятся через requirements.txt."""
    return True


def check_for_updates(venv_python: Path, repo_root: Path):
    """Проверка и установка обновлений."""
    log_verbose("Проверка обновлений...")
    if not ensure_auto_update_package(venv_python, repo_root):
        print("Не удалось установить необходимые пакеты для автообновления.")
        return
    try:
        auto_update_file = repo_root / "scripts" / "_auto_update.py"
        if not auto_update_file.exists():
            return
        result = subprocess.run(
            [str(venv_python), str(auto_update_file)],
            cwd=repo_root,
            timeout=60,
            capture_output=not VERBOSE,
        )
        if result.returncode == 0:
            log_verbose("Проверка обновлений завершена.")
    except subprocess.TimeoutExpired:
        log_verbose("Превышено время ожидания проверки обновлений.")
    except Exception as e:
        print(f"Ошибка проверки обновлений: {e}")


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
        print_step("Создание окружения .venv")
        run([sys.executable, "-m", "venv", str(venv_dir)])
        created = True
    venv_python = venv_dir / "Scripts" / "python.exe"
    if not venv_python.exists():
        venv_python = venv_dir / "bin" / "python"
    if not venv_python.exists():
        raise RuntimeError("Не найден исполняемый файл Python в .venv")
    return venv_python, created


def ensure_deps(venv_python: Path, repo_root: Path):
    venv_dir = Path(venv_python).resolve().parent.parent
    bootstrap_marker = venv_dir / ".bootstrap_done"
    if bootstrap_marker.exists():
        log_verbose("Зависимости уже установлены.")
        return
    print_step("Установка зависимостей")
    run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip"], cwd=repo_root, quiet=True)
    config_dir = Path(__file__).resolve().parent
    req = config_dir / "requirements.txt"
    if req.exists():
        run([str(venv_python), "-m", "pip", "install", "-r", str(req)], cwd=repo_root, quiet=True)
    else:
        run(
            [
                str(venv_python), "-m", "pip", "install",
                "requests", "pandas", "openpyxl", "python-dateutil", "python-dotenv", "packaging",
            ],
            cwd=repo_root,
            quiet=True,
        )
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
    if prompt_yes_no("Добавить Performance API credentials для рекламы? (опционально)", default_yes=False):
        print("\nДля рекламы нужны отдельные ключи из раздела Продвижение -> API в кабинете Ozon.")
        perf_client_id = input("Введите OZON_PERF_CLIENT_ID (или Enter для пропуска): ").strip()
        perf_api_key = input("Введите OZON_PERF_API_KEY (или Enter для пропуска): ").strip()
        if perf_client_id and perf_api_key:
            env_content += f"OZON_PERF_CLIENT_ID={perf_client_id}\nOZON_PERF_API_KEY={perf_api_key}\n"
            print("Performance API credentials добавлены.")
        else:
            print("Performance API credentials пропущены.")

    env_path.write_text(env_content, encoding="utf-8")
    print("Файл .env создан.")
    return True


def ensure_costs(venv_python, repo_root: Path) -> bool:
    costs_xlsx = repo_root / "costs.xlsx"
    if costs_xlsx.exists():
        print_step("Файл себестоимости найден")
        return False

    print_step("Создание шаблона себестоимости costs.xlsx")
    create_cmd = (
        "import sys; "
        f"path=r'{str(costs_xlsx)}'; "
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
        timeout=10,
    )
    if result.returncode != 0:
        raise RuntimeError("Не удалось создать файл costs.xlsx")

    print("Создан costs.xlsx. Заполните артикулы и себестоимость.")
    if prompt_yes_no("Открыть costs.xlsx сейчас?", default_yes=True):
        open_local_file(costs_xlsx)
    return True


def ensure_reports_dir(repo_root: Path):
    (repo_root / "reports").mkdir(parents=True, exist_ok=True)


def open_local_file(path: Path):
    if not path.exists():
        print(f"Файл не найден: {path.name}")
        return
    target = str(path)
    try:
        if os.name == "nt":
            os.startfile(target)
        elif sys.platform == "darwin":
            run(["open", target])
        else:
            run(["xdg-open", target])
    except Exception as e:
        print(f"Не удалось открыть файл автоматически: {e}")


def read_env_flags(repo_root: Path) -> dict:
    env_path = repo_root / ".env"
    flags = {
        "ozon_client_id": False,
        "ozon_api_key": False,
        "perf_client_id": False,
        "perf_api_key": False,
    }
    if not env_path.exists():
        return flags
    try:
        lines = env_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    except Exception:
        return flags
    for line in lines:
        raw = line.strip()
        if not raw or raw.startswith("#") or "=" not in raw:
            continue
        key, value = raw.split("=", 1)
        key = key.strip()
        value = value.strip()
        if key == "OZON_CLIENT_ID" and value:
            flags["ozon_client_id"] = True
        elif key == "OZON_API_KEY" and value:
            flags["ozon_api_key"] = True
        elif key == "OZON_PERF_CLIENT_ID" and value:
            flags["perf_client_id"] = True
        elif key == "OZON_PERF_API_KEY" and value:
            flags["perf_api_key"] = True
    return flags


def list_report_files(repo_root: Path, folder_name: str) -> list[Path]:
    folder = repo_root / folder_name
    if not folder.exists():
        return []
    return sorted(folder.glob("*.xlsx"), key=lambda p: p.name)


def print_system_status(repo_root: Path):
    env_flags = read_env_flags(repo_root)
    ozon_ok = env_flags["ozon_client_id"] and env_flags["ozon_api_key"]
    perf_ok = env_flags["perf_client_id"] and env_flags["perf_api_key"]
    costs_exists = (repo_root / "costs.xlsx").exists()
    reports_count = len(list_report_files(repo_root, "reports"))

    print("Статус системы")
    print(f"[{'OK' if ozon_ok else 'WARN'}] Ozon API {'настроен' if ozon_ok else 'не настроен'}")
    print(f"[{'OK' if costs_exists else 'WARN'}] costs.xlsx {'найден' if costs_exists else 'не найден'}")
    print(f"[OK] Monthly reports: {reports_count}")
    print(f"[{'OK' if perf_ok else 'WARN'}] Performance API {'настроен' if perf_ok else 'не настроен'}")


def choose_menu(title: str, options: list[str], back_label: str = "Назад") -> str:
    print()
    print("  " + "=" * 50)
    print(f"  {title}")
    print("  " + "=" * 50)
    print()
    for idx, option in enumerate(options, start=1):
        print(f"  {idx}. {option}")
    print(f"  0. {back_label}")
    print()
    valid = {str(i) for i in range(0, len(options) + 1)}
    while True:
        answer = input("  Введите номер: ").strip().lower()
        if answer in valid:
            return answer
        if answer in ("q", "quit", "exit"):
            return "0"
        print("  Введите корректный номер из списка.")


def print_main_screen(repo_root: Path):
    print()
    print("  " + "=" * 50)
    print("  OzonReportX")
    print("  " + "=" * 50)
    print()
    print_system_status(repo_root)
    print()
    print("Разделы")
    print("1. Отчёты")
    print("2. Цены и акции")
    print("3. Поставки")
    print("4. Финансы")
    print("5. Настройки")
    print("6. Выход")
    print()


def select_main_menu_option() -> str:
    while True:
        choice = input("Введите номер раздела (1-6) или q для выхода: ").strip().lower()
        if choice in ("q", "quit", "exit", "6"):
            return "6"
        if choice in ("1", "2", "3", "4", "5"):
            return choice
        print("Введите число от 1 до 6.")


def run_report(venv_python: Path, repo_root: Path):
    print_step("Запуск формирования отчёта")
    main_script = repo_root / "scripts" / "Monthly_sales_report.py"
    run([str(venv_python), str(main_script)], cwd=repo_root)


MONTHS_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]


def ask_abc_xyz_date_range():
    print("Укажите диапазон месяцев для анализа.")
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
                    print("Некорректное начало периода.")
                    continue
                if not (1 <= to_month <= 12 and 2000 <= to_year <= 2100):
                    print("Некорректный конец периода.")
                    continue
                if (from_year, from_month) <= (to_year, to_month):
                    return from_month, from_year, to_month, to_year
                print("Начало периода должно быть не позже конца.")
            else:
                print("Введите месяц и год через пробел.")
        except ValueError:
            print("Некорректный ввод. Используйте числа.")


def run_abc_xyz(
    venv_python: Path,
    repo_root: Path,
    from_month: int,
    from_year: int,
    to_month: int,
    to_year: int,
):
    ensure_reports_dir(repo_root)
    reports_dir = repo_root / "reports"
    main_script = repo_root / "scripts" / "Monthly_sales_report.py"
    abc_script = repo_root / "scripts" / "ABC_XYZ_analytics_report.py"

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
            print(f"Формируется отчёт за {label}...")
            run([str(venv_python), str(main_script), "--month", str(m), "--year", str(y)], cwd=repo_root)

    print_step("ABC/XYZ-анализ")
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


def run_recommended_prices(venv_python: Path, repo_root: Path):
    script = repo_root / "scripts" / "recommended_prices.py"
    if not script.exists():
        print("Скрипт recommended_prices.py не найден.")
        return
    run([str(venv_python), str(script)], cwd=repo_root)


def run_update_prices(venv_python: Path, repo_root: Path):
    script = repo_root / "scripts" / "update_prices.py"
    if not script.exists():
        print("Скрипт update_prices.py не найден.")
        return
    run([str(venv_python), str(script)], cwd=repo_root)


def ensure_base_requirements(venv_python: Path, repo_root: Path, need_costs: bool = False):
    ensure_env(repo_root)
    if need_costs:
        ensure_costs(venv_python, repo_root)
    ensure_reports_dir(repo_root)


def run_monthly_report_flow(venv_python: Path, repo_root: Path, venv_created: bool):
    print_step("Создать месячный отчёт")
    env_created = ensure_env(repo_root)
    costs_created = ensure_costs(venv_python, repo_root)
    ensure_reports_dir(repo_root)
    print("Что будет сделано:")
    print("- загрузка FBS/FBO заказов")
    print("- расчёт показателей")
    print("- добавление бизнес-метрик в Excel")
    print("- сохранение отчёта в папку reports")
    if venv_created or env_created or costs_created:
        if not prompt_yes_no("Продолжить создание месячного отчёта?", default_yes=True):
            print("Операция отменена.")
            return
    run_report(venv_python, repo_root)


def show_report_files(repo_root: Path, folder_name: str, title: str):
    print_step(title)
    files = list_report_files(repo_root, folder_name)
    if not files:
        print(f"Отчёты в папке {folder_name} не найдены.")
        return
    for idx, file in enumerate(files, start=1):
        print(f"{idx}. {file.name}")


def run_abc_xyz_flow(venv_python: Path, repo_root: Path):
    print_step("Построить ABC/XYZ-анализ")
    ensure_base_requirements(venv_python, repo_root, need_costs=True)
    from_month, from_year, to_month, to_year = ask_abc_xyz_date_range()
    run_abc_xyz(venv_python, repo_root, from_month, from_year, to_month, to_year)


def run_fbo_supply_flow(venv_python: Path, repo_root: Path):
    print_step("Рассчитать поставку FBO")
    ensure_env(repo_root)
    ensure_reports_dir(repo_root)
    fbo_script = repo_root / "scripts" / "fbo_supply_report.py"
    if not fbo_script.exists():
        print("Модуль расчёта поставок FBO не найден.")
        return
    print("Будет рассчитано:")
    print("- среднедневные продажи")
    print("- текущий остаток FBO")
    print("- дефицит и рекомендуемая поставка")
    if prompt_yes_no("Продолжить расчёт поставки FBO?", default_yes=True):
        run([str(venv_python), str(fbo_script)], cwd=repo_root)
    else:
        print("Операция отменена.")


def run_balance_report_flow(venv_python: Path, repo_root: Path):
    print_step("Скачать отчёт по балансу")
    ensure_env(repo_root)
    balance_script = repo_root / "scripts" / "balance_report.py"
    if not balance_script.exists():
        print("Модуль отчёта о балансе не найден.")
        return
    print("Ограничение Ozon: период не более 30 дней.")
    if prompt_yes_no("Продолжить выгрузку отчёта по балансу?", default_yes=True):
        run([str(venv_python), str(balance_script)], cwd=repo_root)
    else:
        print("Операция отменена.")


def run_price_management_flow(venv_python: Path, repo_root: Path):
    print_step("Цены и акции")
    ensure_costs(venv_python, repo_root)
    ensure_env(repo_root)
    ensure_reports_dir(repo_root)
    price_management_script = repo_root / "scripts" / "price_management.py"
    if not price_management_script.exists():
        print("Модуль управления ценами не найден.")
        return
    run([str(venv_python), str(price_management_script)], cwd=repo_root)


def show_reports_menu(venv_python: Path, repo_root: Path, venv_created: bool):
    while True:
        choice = choose_menu(
            "ОТЧЁТЫ",
            [
                "Создать месячный отчёт",
                "Открыть список готовых отчётов",
                "Построить ABC/XYZ-анализ",
            ],
        )
        if choice == "0":
            return
        if choice == "1":
            run_monthly_report_flow(venv_python, repo_root, venv_created)
        elif choice == "2":
            show_report_files(repo_root, "reports", "Готовые месячные отчёты")
        elif choice == "3":
            run_abc_xyz_flow(venv_python, repo_root)


def show_pricing_menu(venv_python: Path, repo_root: Path):
    while True:
        choice = choose_menu(
            "ЦЕНЫ И АКЦИИ",
            [
                "Открыть раздел управления ценами и акциями",
                "Рассчитать рекомендованные цены",
                "Обновить минимальные цены на Ozon",
            ],
        )
        if choice == "0":
            return
        if choice == "1":
            run_price_management_flow(venv_python, repo_root)
        elif choice == "2":
            ensure_costs(venv_python, repo_root)
            ensure_env(repo_root)
            ensure_reports_dir(repo_root)
            run_recommended_prices(venv_python, repo_root)
        elif choice == "3":
            ensure_costs(venv_python, repo_root)
            ensure_env(repo_root)
            run_update_prices(venv_python, repo_root)


def show_supply_menu(venv_python: Path, repo_root: Path):
    while True:
        choice = choose_menu(
            "ПОСТАВКИ",
            [
                "Рассчитать поставку FBO",
                "Открыть последние отчёты по поставкам",
            ],
        )
        if choice == "0":
            return
        if choice == "1":
            run_fbo_supply_flow(venv_python, repo_root)
        elif choice == "2":
            show_report_files(repo_root, "stocks reports", "Отчёты по поставкам")


def show_finance_menu(venv_python: Path, repo_root: Path):
    while True:
        choice = choose_menu(
            "ФИНАНСЫ",
            [
                "Скачать отчёт по балансу",
                "Показать последние выгрузки",
            ],
        )
        if choice == "0":
            return
        if choice == "1":
            run_balance_report_flow(venv_python, repo_root)
        elif choice == "2":
            show_report_files(repo_root, "balance reports", "Последние выгрузки баланса")


def show_settings_menu(venv_python: Path, repo_root: Path):
    while True:
        choice = choose_menu(
            "НАСТРОЙКИ",
            [
                "Проверить статус системы",
                "Создать или проверить .env",
                "Открыть costs.xlsx",
                "Проверить обновления",
            ],
        )
        if choice == "0":
            return
        if choice == "1":
            print_system_status(repo_root)
        elif choice == "2":
            ensure_env(repo_root)
        elif choice == "3":
            ensure_costs(venv_python, repo_root)
            open_local_file(repo_root / "costs.xlsx")
        elif choice == "4":
            check_for_updates(venv_python, repo_root)


def main():
    repo_root = Path(__file__).resolve().parent.parent

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
    check_for_updates(venv_python, repo_root)

    while True:
        print_main_screen(repo_root)
        choice = select_main_menu_option()
        if choice == "6":
            print("\nВыход из программы.")
            return
        try:
            if choice == "1":
                show_reports_menu(venv_python, repo_root, venv_created)
            elif choice == "2":
                run_price_management_flow(venv_python, repo_root)
            elif choice == "3":
                show_supply_menu(venv_python, repo_root)
            elif choice == "4":
                show_finance_menu(venv_python, repo_root)
            elif choice == "5":
                show_settings_menu(venv_python, repo_root)
        except KeyboardInterrupt:
            print("\nОперация прервана пользователем.")
        except Exception as e:
            print(f"\nОшибка: {e}")
            import traceback
            print(traceback.format_exc())


if __name__ == "__main__":
    main()
