# -*- coding: utf-8 -*-
"""
Расчёт поставок FBO: на основе ABC&XYZ за последние 3 месяца строит отчёт
с артикулом, классом ABC/XYZ, дневными продажами за 90 дней, остатком FBO и метриками спроса.
Сохраняет в папку stocks reports с именем «Расчёт поставок {Месяц} {Год}.xlsx».
"""

import os
import sys
from pathlib import Path
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import numpy as np

script_dir = Path(__file__).resolve().parent
repo_root = script_dir.parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

MONTHS_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

ABC_XYZ_DIR = "ABC&XYZ reports"
REPORTS_DIR = "reports"
STOCKS_REPORTS_DIR = "stocks reports"


def _last_three_months_from_now() -> Tuple[int, int, int, int]:
    """Возвращает (from_month, from_year, to_month, to_year) за последние 3 месяца.
    Например, если сейчас январь 2025 → (10, 2024, 12, 2024).
    """
    today = date.today()
    # Конец периода: последний день предыдущего месяца (чтобы 3 полных месяца: окт, ноя, дек)
    end = today.replace(day=1) - timedelta(days=1)
    # Начало: 3 месяца назад от конца
    start_year = end.year
    start_month = end.month - 2
    if start_month <= 0:
        start_month += 12
        start_year -= 1
    return start_month, start_year, end.month, end.year


def _ensure_abc_xyz_report(repo_root: Path, from_month: int, from_year: int, to_month: int, to_year: int) -> Optional[Path]:
    """Проверяет наличие ABC&XYZ отчёта за период; при отсутствии запускает построение. Возвращает путь к файлу."""
    abc_dir = repo_root / ABC_XYZ_DIR
    abc_dir.mkdir(parents=True, exist_ok=True)
    early_name = f"{MONTHS_RU[from_month - 1]} {from_year}"
    late_name = f"{MONTHS_RU[to_month - 1]} {to_year}"
    expected_name = f"{early_name}-{late_name}.xlsx"
    expected_path = abc_dir / expected_name
    if expected_path.exists():
        return expected_path
    # Запуск построения ABC&XYZ
    reports_dir = repo_root / REPORTS_DIR
    reports_dir.mkdir(parents=True, exist_ok=True)
    main_script = repo_root / "scripts" / "Monthly_sales_report.py"
    abc_script = repo_root / "scripts" / "ABC_XYZ_analytics_report.py"
    venv_python = repo_root / ".venv" / "Scripts" / "python.exe"
    if not venv_python.exists():
        venv_python = repo_root / ".venv" / "bin" / "python"
    if not venv_python.exists():
        venv_python = Path(sys.executable)
    import subprocess
    # Недостающие месячные отчёты
    months_need = []
    for y in range(from_year, to_year + 1):
        for m in range(1, 13):
            if (y, m) < (from_year, from_month):
                continue
            if (y, m) > (to_year, to_month):
                break
            months_need.append((y, m))
    for y, m in months_need:
        fname = f"{MONTHS_RU[m - 1]} {y}.xlsx"
        if not (reports_dir / fname).exists() and main_script.exists():
            subprocess.run(
                [str(venv_python), str(main_script), "--month", str(m), "--year", str(y)],
                cwd=repo_root,
                check=False,
            )
    if abc_script.exists():
        subprocess.run(
            [
                str(venv_python), str(abc_script),
                "-i", str(reports_dir),
                "--output_dir", ABC_XYZ_DIR,
                "--from-month", str(from_month), "--from-year", str(from_year),
                "--to-month", str(to_month), "--to-year", str(to_year),
            ],
            cwd=repo_root,
            check=False,
        )
    return expected_path if expected_path.exists() else None


def _normalize_artikul(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _artikul_canonical(s: str) -> str:
    """Единый ключ артикула для сопоставления (12345.0 и 12345 -> 12345)."""
    if not s or not isinstance(s, str):
        return (s or "").strip()
    s = s.strip()
    if s.lower() == "nan":
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return s
    except (ValueError, TypeError):
        return s


def _load_abc_xyz_itog_and_orders(abc_path: Path) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Загружает листы Итог и Заказы из ABC&XYZ отчёта."""
    itog = pd.read_excel(abc_path, sheet_name="Итог")
    orders = pd.read_excel(abc_path, sheet_name="Заказы")
    return itog, orders


def _daily_sales_90_from_orders(orders_df: pd.DataFrame, from_date: date, to_date: date) -> Dict[str, List[int]]:
    """
    Строит по каждому артикулу список продаж по дням за период [from_date, to_date] (90 дней).
    Возвращает {artikul: [q_day1, q_day2, ..., q_day90]}.
    """
    if "Артикул" not in orders_df.columns or "Дата отгрузки" not in orders_df.columns or "Количество шт." not in orders_df.columns:
        return {}
    orders_df = orders_df.copy()
    orders_df["Дата отгрузки"] = pd.to_datetime(orders_df["Дата отгрузки"], errors="coerce").dt.date
    orders_df = orders_df.dropna(subset=["Дата отгрузки"])
    orders_df["Количество шт."] = pd.to_numeric(orders_df["Количество шт."], errors="coerce").fillna(0).astype(int)
    # Ровно 90 дней: from_date .. from_date+89
    day_range = [from_date + timedelta(days=i) for i in range(90)]
    date_to_idx = {d: i for i, d in enumerate(day_range)}
    artikul_days: Dict[str, List[int]] = {}
    for _, row in orders_df.iterrows():
        art = _artikul_canonical(str(row.get("Артикул", "") or ""))
        if not art:
            continue
        dt = row.get("Дата отгрузки")
        if dt not in date_to_idx:
            continue
        idx = date_to_idx[dt]
        q = int(row.get("Количество шт.", 0) or 0)
        if art not in artikul_days:
            artikul_days[art] = [0] * len(day_range)
        artikul_days[art][idx] = artikul_days[art][idx] + q
    return artikul_days


def _metrics_90(daily: List[int]) -> Tuple[float, float, float]:
    """avg_daily_sales, std_daily_sales, zero_days_ratio для списка из 90 дней."""
    n = len(daily) or 90
    if n == 0:
        return 0.0, 0.0, 0.0
    arr = np.array(daily, dtype=float)
    total = arr.sum()
    avg = total / n
    std = float(np.std(arr)) if n > 1 else 0.0
    zeros = int((arr == 0).sum())
    zero_ratio = zeros / n
    return round(avg, 4), round(std, 4), round(zero_ratio, 4)


def run(repo_root_path: Optional[Path] = None) -> bool:
    repo_root_path = repo_root_path or repo_root
    from_month, from_year, to_month, to_year = _last_three_months_from_now()
    early_name = f"{MONTHS_RU[from_month - 1]} {from_year}"
    late_name = f"{MONTHS_RU[to_month - 1]} {to_year}"
    print(f"· Период ABC&XYZ: {early_name} — {late_name} (последние 3 месяца)")

    abc_path = _ensure_abc_xyz_report(repo_root_path, from_month, from_year, to_month, to_year)
    if not abc_path or not abc_path.exists():
        print("❌ Не удалось получить или построить ABC&XYZ отчёт за этот период.")
        return False

    print(f"· Используется ABC&XYZ: {abc_path.name}")
    itog_df, orders_df = _load_abc_xyz_itog_and_orders(abc_path)

    # Диапазон 90 дней: конец периода = последний день to_month
    to_date = date(to_year, to_month, 1) + timedelta(days=32)
    to_date = to_date.replace(day=1) - timedelta(days=1)
    from_date = to_date - timedelta(days=89)
    daily_sales = _daily_sales_90_from_orders(orders_df, from_date, to_date)

    # Колонки Итог: Артикул, Оценка по ABC, Оценка по XYZ
    if "Артикул" not in itog_df.columns or "Оценка по ABC" not in itog_df.columns or "Оценка по XYZ" not in itog_df.columns:
        print("❌ В ABC&XYZ отчёте нет ожидаемых колонок (Артикул, Оценка по ABC, Оценка по XYZ).")
        return False

    offer_ids = []
    for _, row in itog_df.iterrows():
        art = _normalize_artikul(row.get("Артикул"))
        if art:
            offer_ids.append(art)

    from recommended_prices import get_fbo_stocks_by_offer_ids, _normalize_offer_id

    # Нормализуем артикулы как в API (12345.0 -> 12345), чтобы совпадали ключи
    offer_ids_norm = [_normalize_offer_id(str(a).strip()) for a in offer_ids if _normalize_offer_id(str(a).strip())]
    offer_ids_norm = list(dict.fromkeys(offer_ids_norm))

    print("· Запрос остатков FBO (Ozon API /v4/product/info/stocks)...")
    stocks = get_fbo_stocks_by_offer_ids(offer_ids_norm)

    rows = []
    for _, row in itog_df.iterrows():
        art_raw = row.get("Артикул")
        art = _normalize_offer_id(str(art_raw).strip()) if art_raw is not None else ""
        if not art:
            continue
        abc_class = str(row.get("Оценка по ABC", "") or "").strip()
        xyz_class = str(row.get("Оценка по XYZ", "") or "").strip()
        daily = daily_sales.get(art, [0] * 90)
        if len(daily) < 90:
            daily = (daily + [0] * 90)[:90]
        avg_daily, std_daily, zero_ratio = _metrics_90(daily)
        available_fbo = stocks.get(art)
        sales_by_day_str = ",".join(str(q) for q in daily)
        rows.append({
            "Артикул": art,
            "abc_class": abc_class,
            "xyz_class": xyz_class,
            "sales_by_day": sales_by_day_str,
            "Активный сток FBO": available_fbo if available_fbo is not None else 0,
            "avg_daily_sales": avg_daily,
            "std_daily_sales": std_daily,
            "zero_days_ratio": zero_ratio,
        })

    df = pd.DataFrame(rows)

    out_dir = repo_root_path / STOCKS_REPORTS_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    now = date.today()
    out_name = f"Расчёт поставок {MONTHS_RU[now.month - 1]} {now.year}.xlsx"
    out_path = out_dir / out_name
    df.to_excel(out_path, index=False, sheet_name="Главная")
    print(f"✅ Отчёт сохранён: {out_path}")
    return True


def main():
    import argparse
    p = argparse.ArgumentParser(description="Расчёт поставок FBO по ABC&XYZ за 3 месяца.")
    p.add_argument("--repo", type=Path, default=None, help="Корень репозитория")
    args = p.parse_args()
    ok = run(args.repo)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
