"""
Скачивание отчёта о балансе (Ozon API v1/finance/balance).
Соответствует разделу Финансы → Баланс в ЛК. Сохраняется в папку «balance reports».
Период не более 30 дней.
"""

import os
import sys
import json
import argparse
from datetime import datetime, timedelta
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv()
CLIENT_ID = os.getenv("OZON_CLIENT_ID")
API_KEY = os.getenv("OZON_API_KEY")

if not CLIENT_ID or not API_KEY:
    print("❌ Укажите OZON_CLIENT_ID и OZON_API_KEY в файле .env")
    sys.exit(1)

HEADERS = {
    "Client-Id": CLIENT_ID,
    "Api-Key": API_KEY,
    "Content-Type": "application/json",
}

BALANCE_REPORTS_DIR = "balance reports"
MAX_DAYS = 30


def _extract_star_products_value(data: dict) -> float:
    """Из ответа balance API извлекает сумму по «Звёздным товарам» (расход на продвижение)."""
    cf = data.get("cashflows") or {}
    total = 0.0
    # Вариант 1: отдельное поле star_products
    sp = cf.get("star_products")
    if sp is not None and isinstance(sp, dict):
        amt = sp.get("amount")
        if isinstance(amt, dict) and "value" in amt:
            return float(amt.get("value") or 0)
    # Вариант 2: в services по имени (например «Звёздные товары», stars)
    for s in cf.get("services") or []:
        name = (s.get("name") or "").lower()
        if "star" in name or "звезд" in name or "star_products" in name:
            amt = s.get("amount")
            if isinstance(amt, dict) and "value" in amt:
                total += float(amt.get("value") or 0)
    return total


def _extract_product_placement_in_ozon_warehouses_value(data: dict) -> float:
    """Из ответа balance API извлекает сумму по хранению FBO."""
    cf = data.get("cashflows") or {}
    total = 0.0

    placement = cf.get("product_placement_in_ozon_warehouses")
    if placement is not None and isinstance(placement, dict):
        amt = placement.get("amount")
        if isinstance(amt, dict) and "value" in amt:
            return float(amt.get("value") or 0)

    for s in cf.get("services") or []:
        name = (s.get("name") or "").lower()
        if (
            "product_placement_in_ozon_warehouses" in name
            or "placement" in name
            or "warehouse" in name
            or "хранен" in name
            or "размещ" in name
        ):
            amt = s.get("amount")
            if isinstance(amt, dict) and "value" in amt:
                total += float(amt.get("value") or 0)
    return total


def get_star_products_for_month(month: int, year: int) -> float:
    """
    Запрашивает отчёт о балансе за месяц (с разбивкой по 30 дней) и возвращает
    суммарные расходы по «Звёздным товарам» за период.
    """
    from calendar import monthrange
    first = f"{year}-{month:02d}-01"
    last_day = monthrange(year, month)[1]
    total = 0.0
    # Первый отрезок: 01 — min(30, last_day)
    end1 = min(30, last_day)
    to1 = f"{year}-{month:02d}-{end1:02d}"
    try:
        data1 = get_balance_report(first, to1)
        total += _extract_star_products_value(data1)
    except Exception:
        pass
    # Если в месяце 31 день — отдельный запрос за 31-е
    if last_day == 31:
        try:
            data2 = get_balance_report(f"{year}-{month:02d}-31", f"{year}-{month:02d}-31")
            total += _extract_star_products_value(data2)
        except Exception:
            pass
    return total


def get_product_placement_in_ozon_warehouses_for_month(month: int, year: int) -> float:
    """
    Запрашивает отчёт о балансе за месяц (с разбивкой по 30 дней) и возвращает
    суммарные расходы по хранению FBO за период.
    """
    from calendar import monthrange

    first = f"{year}-{month:02d}-01"
    last_day = monthrange(year, month)[1]
    total = 0.0

    end1 = min(30, last_day)
    to1 = f"{year}-{month:02d}-{end1:02d}"
    try:
        data1 = get_balance_report(first, to1)
        total += _extract_product_placement_in_ozon_warehouses_value(data1)
    except Exception:
        pass

    if last_day == 31:
        try:
            data2 = get_balance_report(f"{year}-{month:02d}-31", f"{year}-{month:02d}-31")
            total += _extract_product_placement_in_ozon_warehouses_value(data2)
        except Exception:
            pass

    return total


def get_balance_report(date_from: str, date_to: str) -> dict:
    """Запрашивает отчёт о балансе за период. date_from/date_to в формате YYYY-MM-DD (без времени)."""
    url = "https://api-seller.ozon.ru/v1/finance/balance"
    payload = {
        "date_from": date_from,
        "date_to": date_to,
    }
    resp = requests.post(url, headers=HEADERS, json=payload, timeout=60)
    if resp.status_code != 200:
        try:
            err = resp.json()
            msg = err.get("message", resp.text)
            code = err.get("code", resp.status_code)
        except Exception:
            msg = resp.text or resp.reason
            code = resp.status_code
        raise RuntimeError(f"API ошибка {code}: {msg}")
    return resp.json()


def ensure_balance_reports_dir(repo_root: Path) -> Path:
    out_dir = repo_root / BALANCE_REPORTS_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def _money(obj: dict) -> str:
    if not obj or not isinstance(obj, dict):
        return "—"
    val = obj.get("value")
    code = obj.get("currency_code", "")
    if val is None:
        return "—"
    return f"{val} {code}".strip()


def save_report(data: dict, date_from: str, date_to: str, repo_root: Path) -> Path:
    """Сохраняет отчёт: JSON и Excel (сводка по балансу и cashflows)."""
    out_dir = ensure_balance_reports_dir(repo_root)
    base_name = f"{date_from}_to_{date_to}"
    json_path = out_dir / f"{base_name}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    try:
        import pandas as pd
        total = data.get("total") or {}
        cashflows = data.get("cashflows") or {}

        rows_total = []
        for key, label in [
            ("opening_balance", "Входящий баланс"),
            ("closing_balance", "Исходящий баланс"),
            ("accrued", "Начислено"),
        ]:
            val = total.get(key)
            rows_total.append({"Показатель": label, "Значение": _money(val) if isinstance(val, dict) else val})
        payments = total.get("payments") or []
        for i, p in enumerate(payments):
            rows_total.append({"Показатель": f"Выплата {i + 1}", "Значение": _money(p) if isinstance(p, dict) else p})

        sales = cashflows.get("sales") or {}
        returns = cashflows.get("returns") or {}
        rows_cf = [
            {"Раздел": "Продажи", "Сумма": _money(sales.get("amount")), "Комиссия": _money(sales.get("fee"))},
            {"Раздел": "Возвраты", "Сумма": _money(returns.get("amount")), "Комиссия": _money(returns.get("fee"))},
        ]
        services = cashflows.get("services") or []
        for s in services:
            name = s.get("name", "Услуга")
            rows_cf.append({"Раздел": name, "Сумма": _money(s.get("amount")), "Комиссия": "—"})

        excel_path = out_dir / f"{base_name}.xlsx"
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            pd.DataFrame(rows_total).to_excel(writer, sheet_name="Баланс", index=False)
            pd.DataFrame(rows_cf).to_excel(writer, sheet_name="Движения", index=False)
        print(f"   Excel: {excel_path.name}")
    except Exception as e:
        print(f"   (Excel не создан: {e})")
    return json_path


def ask_date_range():
    """Интерактивный ввод периода (макс. 30 дней)."""
    while True:
        try:
            from_str = input("Дата начала (YYYY-MM-DD): ").strip()
            to_str = input("Дата окончания (YYYY-MM-DD): ").strip()
            d_from = datetime.strptime(from_str, "%Y-%m-%d")
            d_to = datetime.strptime(to_str, "%Y-%m-%d")
            if d_from > d_to:
                print("⚠️ Дата начала не может быть позже даты окончания.")
                continue
            delta = (d_to - d_from).days + 1
            if delta > MAX_DAYS:
                print(f"⚠️ Период не более {MAX_DAYS} дней.")
                continue
            return from_str, to_str
        except ValueError:
            print("⚠️ Введите даты в формате YYYY-MM-DD.")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Скачать отчёт о балансе (Ozon). Период до 30 дней."
    )
    parser.add_argument("--date_from", default=None, help="Начало периода YYYY-MM-DD")
    parser.add_argument("--date_to", default=None, help="Конец периода YYYY-MM-DD")
    args = parser.parse_args(argv)

    if args.date_from and args.date_to:
        try:
            d_from = datetime.strptime(args.date_from, "%Y-%m-%d")
            d_to = datetime.strptime(args.date_to, "%Y-%m-%d")
            if d_from > d_to or (d_to - d_from).days + 1 > MAX_DAYS:
                print(f"❌ Некорректный период (макс. {MAX_DAYS} дней).")
                sys.exit(1)
            date_from, date_to = args.date_from, args.date_to
        except ValueError:
            print("❌ Даты в формате YYYY-MM-DD.")
            sys.exit(1)
    else:
        print("\n📋 Отчёт о балансе (Финансы → Баланс)")
        print(f"   Период не более {MAX_DAYS} дней.\n")
        date_from, date_to = ask_date_range()

    repo_root = Path(__file__).resolve().parent.parent
    print(f"\n⏳ Запрос отчёта за {date_from} – {date_to}...")
    try:
        data = get_balance_report(date_from, date_to)
    except RuntimeError as e:
        print(f"❌ {e}")
        sys.exit(1)
    except requests.RequestException as e:
        print(f"❌ Ошибка сети: {e}")
        sys.exit(1)

    total = data.get("total") or {}
    print(f"   Входящий баланс: {_money(total.get('opening_balance'))}")
    print(f"   Исходящий баланс: {_money(total.get('closing_balance'))}")

    json_path = save_report(data, date_from, date_to, repo_root)
    print(f"✅ Отчёт сохранён в папку «{BALANCE_REPORTS_DIR}»:")
    print(f"   JSON: {json_path.name}")


if __name__ == "__main__":
    main()
