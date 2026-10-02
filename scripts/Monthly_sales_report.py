import os
import json
import sys
import time
import shutil
import importlib
from contextlib import ExitStack, closing
from datetime import datetime, timedelta, timezone
from decimal import Decimal
from typing import List, Dict, Any, Optional

import pandas as pd
import requests
from dateutil.parser import isoparse
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from dotenv import load_dotenv
try:
    from scripts.report_http import ReportSession
except ModuleNotFoundError:
    from report_http import ReportSession
try:
    from scripts.file_io import excel_writer, atomic_output_path, save_workbook_atomic, read_costs_dataframe
except ModuleNotFoundError:
    from file_io import excel_writer, atomic_output_path, save_workbook_atomic, read_costs_dataframe

# 🔐 Загружаем данные для авторизации из переменных окружения
load_dotenv()
CLIENT_ID = os.getenv('OZON_CLIENT_ID')
API_KEY = os.getenv('OZON_API_KEY')

if not CLIENT_ID or not API_KEY:
    raise RuntimeError("Отсутствуют переменные OZON_CLIENT_ID или OZON_API_KEY. Укажите их в .env или окружении.")

HEADERS = {
    'Client-Id': CLIENT_ID,
    'Api-Key': API_KEY,
    'Content-Type': 'application/json'
}

def create_session() -> requests.Session:
    session = ReportSession()
    retry = Retry(
        total=5,
        backoff_factor=0.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["POST", "GET"]),
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    # 429 for Seller API is handled above the adapter, with one shared pacing gate.
    seller_retry = Retry(total=5, backoff_factor=0.5,
                         status_forcelist=(500, 502, 503, 504),
                         allowed_methods=frozenset(["POST", "GET"]),
                         respect_retry_after_header=False)
    session.mount("https://api-seller.ozon.ru/", HTTPAdapter(max_retries=seller_retry))
    return session

def get_custom_date_range():
    while True:
        try:
            month = int(input("Введите номер месяца (1–12): ").strip())
            year = int(input("Введите год (например, 2025): ").strip())

            if 1 <= month <= 12 and 2000 <= year <= 2100:
                break
            else:
                print("⚠️ Введите корректный месяц (1–12) и год (2000–2100).")
        except ValueError:
            print("❌ Некорректный ввод. Попробуйте снова.")

    from datetime import datetime, timedelta
    from calendar import monthrange

    first_day = datetime(year, month, 1)
    last_day = datetime(year, month, monthrange(year, month)[1])
    date_from = first_day.strftime('%Y-%m-%dT00:00:00Z')
    date_to = last_day.strftime('%Y-%m-%dT23:59:59Z')
    return date_from, date_to, month, year



def _normalize_articul_key(s: str) -> str:
    """Приводит артикул к одному виду для сопоставления (Excel даёт 12345.0, API — 12345)."""
    if not s or not isinstance(s, str):
        return (s or "").strip()
    s = s.strip()
    if s.lower() == 'nan':
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
        return s
    except (ValueError, TypeError):
        return s


# 📄 Загрузка карты себестоимости из внешнего файла
def load_cost_map():
    script_dir = os.path.dirname(__file__)
    repo_root = os.path.abspath(os.path.join(script_dir, '..'))

    candidates = [
        os.path.join(repo_root, 'costs.xlsx'),
    ]

    for path in candidates:
        if os.path.exists(path):
            try:
                if path.endswith('.xlsx'):
                    df = read_costs_dataframe(path)
                else:
                    df = pd.read_csv(path)

                # Нормализуем имена столбцов
                lower_cols = {c.lower(): c for c in df.columns}
                # Поддерживаемые варианты названий
                key_col = None
                cost_col = None

                for variant in ['prefix', 'префикс', 'код', 'артикул', 'offer_id']:
                    if variant in lower_cols:
                        key_col = lower_cols[variant]
                        break
                for variant in ['cost', 'себестоимость', 'цена', 'стоимость']:
                    if variant in lower_cols:
                        cost_col = lower_cols[variant]
                        break

                if not key_col or not cost_col:
                    print(f"⚠️ Файл {os.path.basename(path)} найден, но столбцы не распознаны. \nОжидаются столбцы: 'prefix'/'префикс'/'код'/'артикул' и 'cost'/'себестоимость'.")
                    continue

                mapping = {}
                for _, row in df.iterrows():
                    raw = row.get(key_col, '')
                    key = _normalize_articul_key(str(raw).strip() if raw is not None else '')
                    if not key:
                        continue
                    try:
                        value = float(row.get(cost_col, 0) or 0)
                    except Exception:
                        continue
                    mapping[key] = value

                print(f"🧾 Загружена карта себестоимости из {os.path.basename(path)}: {len(mapping)} записей")
                return mapping
            except Exception as e:
                print(f"⚠️ Не удалось прочитать {os.path.basename(path)}: {e}")

    print("ℹ️ Файл себестоимости не найден (costs.xlsx). Будет использовано значение 0.")
    return {}

# Импорт функций для работы с Performance API
try:
    from scripts.performance_api import get_cpc_campaigns_for_month, get_campaigns_data_for_excel  # type: ignore
except ImportError:
    import sys
    from pathlib import Path
    sys.path.append(str(Path(__file__).resolve().parent))
    from performance_api import get_cpc_campaigns_for_month, get_campaigns_data_for_excel  # type: ignore

# Импорт функций для получения расходов из отчёта о балансе
try:
    from scripts.balance_report import get_star_products_for_month, get_product_placement_in_ozon_warehouses_for_month, get_monthly_balance_reports, summarize_month as summarize_balance_month  # type: ignore
except ImportError:
    from pathlib import Path
    import sys
    sys.path.append(str(Path(__file__).resolve().parent))
    from balance_report import get_star_products_for_month, get_product_placement_in_ozon_warehouses_for_month, get_monthly_balance_reports, summarize_month as summarize_balance_month  # type: ignore

# 📥 Получаем список заказов FBS (Fulfillment by Seller)
def _fetch_fbs_page(session: requests.Session, date_from: str, date_to: str, status: str, limit: int, offset: int) -> List[Dict[str, Any]]:
    url = 'https://api-seller.ozon.ru/v3/posting/fbs/list'
    payload = {
        "filter": {
            "since": date_from,
            "to": date_to,
            "status": status
        },
        "limit": limit,
        "offset": offset,
        "with": {
            "analytics_data": True,
            "financial_data": True
        }
    }
    resp = session.post(url, headers=HEADERS, json=payload, timeout=(10, 60))
    resp.raise_for_status()
    data = resp.json()
    postings = data.get("result", {}).get("postings", [])
    for p in postings:
        p["__schema"] = "FBS"
    return postings

# 📥 Получаем список заказов FBS (Fulfillment by Seller)
def get_orders(date_from, date_to, session: Optional[requests.Session] = None):
    if session is None:
        with create_session() as owned_session:
            return get_orders(date_from, date_to, session=owned_session)
    result = []
    limit = 100
    for status in ['awaiting_packaging', 'awaiting_deliver', 'delivering', 'delivered', 'cancelled']:
        offset = 0
        while True:
            try:
                postings = _fetch_fbs_page(session, date_from, date_to, status, limit, offset)
            except Exception as exc:
                raise RuntimeError(f"Не удалось загрузить FBS: статус {status}, offset {offset}. {exc}") from exc
            if not postings:
                break
            result.extend(postings)
            offset += limit
    return result

# 📥 Получаем список заказов FBO (Fulfillment by Ozon)
def _fetch_fbo_page(session: requests.Session, date_from: str, date_to: str, limit: int, cursor: str) -> Dict[str, Any]:
    url = 'https://api-seller.ozon.ru/v3/posting/fbo/list'
    payload = {
        "sort_dir": "ASC",
        "filter": {
            "since": date_from,
            "to": date_to,
            "statuses": ['awaiting_packaging', 'awaiting_deliver', 'delivering', 'delivered', 'cancelled']
        },
        "limit": limit,
        "cursor": cursor,
        "with": {
            "analytics_data": True,
            "financial_data": True
        }
    }
    resp = session.post(url, headers=HEADERS, json=payload, timeout=(10, 60))
    resp.raise_for_status()
    data = resp.json()
    if (not isinstance(data, dict) or not isinstance(data.get("postings"), list)
            or not isinstance(data.get("has_next"), bool)):
        raise ValueError("Некорректный ответ /v3/posting/fbo/list: ожидаются postings и has_next.")
    postings = data["postings"]
    for p in postings:
        if not isinstance(p, dict):
            raise ValueError("Некорректное отправление в ответе FBO.")
        p["__schema"] = "FBO"
        # Keep the common report representation shared with FBS.
        for product in p.get("products", []) or []:
            price = product.get("price")
            if isinstance(price, dict):
                product["price"] = price["amount"]
                product["currency_code"] = price.get("currency", "")
        for product in (p.get("financial_data") or {}).get("products", []) or []:
            commission = product.get("commission")
            if isinstance(commission, dict):
                product["commission_amount"] = commission.get("amount", 0)
                product["commission_percent"] = commission.get("percent", 0)
                product["currency_code"] = commission.get("currency", "")
    return data

def get_fbo_orders(date_from, date_to, session: Optional[requests.Session] = None):
    if session is None:
        with create_session() as owned_session:
            return get_fbo_orders(date_from, date_to, session=owned_session)
    result = []
    limit = 100
    cursor = ""
    seen_cursors = {cursor}
    while True:
        try:
            page = _fetch_fbo_page(session, date_from, date_to, limit, cursor)
            result.extend(page["postings"])
            if not page["has_next"]:
                break
            next_cursor = page.get("cursor")
            if not isinstance(next_cursor, str) or not next_cursor or next_cursor in seen_cursors:
                raise ValueError("Ozon не вернул новый cursor при has_next=true; неполный отчёт не сохранён.")
            seen_cursors.add(next_cursor)
            cursor = next_cursor
        except Exception as exc:
            raise RuntimeError(f"Не удалось загрузить FBO (v3), страница {len(seen_cursors)}. {exc}") from exc
    return result

def _accrual_money(value):
    """Read signed money without silently turning malformed amounts into zero."""
    if not isinstance(value, dict) or value.get("amount") is None:
        raise ValueError("В начислении отсутствует сумма amount.")
    amount = Decimal(str(value["amount"]))
    if not amount.is_finite():
        raise ValueError("Некорректная сумма начисления.")
    if value.get("currency") != "RUB":
        raise ValueError("Начисление не в RUB: суммирование разных валют не поддерживается.")
    return amount


def get_transactions_by_posting(date_from, date_to, session: Optional[requests.Session] = None):
    """Load daily accruals once, preserving the report's transaction field names.

    /v3/finance/transaction/list was retired on 2026-09-08. Daily accruals
    provide signed totals and sale commission directly, including adjustments.
    The monthly report uses whole UTC calendar days, as does this endpoint.
    """
    if session is None:
        with create_session() as owned_session:
            return get_transactions_by_posting(date_from, date_to, owned_session)
    day = isoparse(date_from).date()
    end = isoparse(date_to).date()
    if day > end:
        raise ValueError("Начало периода позже конца.")
    end = min(end, datetime.now(timezone.utc).date())
    result = {}
    while day <= end:
        date = day.isoformat()
        last_id = ""
        seen_cursors = {last_id}
        while True:
            try:
                response = session.post(
                    "https://api-seller.ozon.ru/v1/finance/accrual/by-day",
                    headers=HEADERS, json={"date": date, "last_id": last_id}, timeout=(10, 60),
                )
                response.raise_for_status()
                data = response.json()
                if (not isinstance(data, dict) or not isinstance(data.get("accruals"), list)
                        or not isinstance(data.get("last_id"), str)):
                    raise ValueError("Некорректный ответ начислений: ожидаются accruals и last_id.")
                for accrual in data["accruals"]:
                    number = accrual.get("unit_number")
                    if not number:
                        continue  # Account-wide charges have no posting to attach to.
                    if not isinstance(number, str):
                        raise ValueError("Некорректный номер отправления в начислении.")
                    sale = Decimal(0)
                    commission = Decimal(0)
                    has_commission_data = False
                    for product in (accrual.get("posting") or {}).get("products", []) or []:
                        details = product.get("commission")
                        if details:
                            has_commission_data = True
                            sale += _accrual_money(details["seller_price"])
                            commission += _accrual_money(details["sale_commission"])
                    result.setdefault(number, []).append({
                        "amount": _accrual_money(accrual["total_amount"]),
                        "sale_commission": commission,
                        "accruals_for_sale": sale,
                        # A return/reversal accrual carries a negative seller_price (Ozon books it
                        # as its own "Возврат" entry). The posting's own status stays "delivered"
                        # forever in this case — Ozon does not reflect post-delivery returns there —
                        # so this is the only reliable signal that the item came back.
                        "is_return": sale < 0,
                        # Whether this accrual actually carried commission data at all, as opposed to
                        # being a delivery/service-fee-only entry (commission: null). Distinguishes a
                        # sale genuinely netted to zero by a same-period return from a posting whose
                        # sale accrual simply fell in a different month.
                        "has_commission_data": has_commission_data,
                    })
                next_id = data["last_id"]
                if not next_id:
                    break
                if next_id in seen_cursors:
                    raise ValueError("Ozon повторил last_id; загрузка начислений не завершена.")
                seen_cursors.add(next_id)
                last_id = next_id
            except Exception as exc:
                raise RuntimeError(f"Не удалось загрузить начисления за {date}. {exc}") from exc
        print(f"💳 Начисления загружены за {date}", flush=True)
        day += timedelta(days=1)
    return result


def get_transactions(posting_number, date_from, date_to, session: Optional[requests.Session] = None):
    return get_transactions_by_posting(date_from, date_to, session).get(posting_number, [])

# 📊 Преобразуем данные в Excel
def _ensure_reports_dir_and_check_space(reports_dir: str, min_free_mb: int = 20) -> None:
    os.makedirs(reports_dir, exist_ok=True)
    try:
        usage = shutil.disk_usage(reports_dir)
        free_mb = usage.free // (1024 * 1024)
    except OSError:
        # Если не удалось определить, продолжаем без жёсткой блокировки
        return
    if free_mb < min_free_mb:
        raise RuntimeError(f"Недостаточно места на диске: доступно {free_mb} МБ, требуется ≥ {min_free_mb} МБ")

def _artikul_to_number(v):
    """Преобразует значение артикула в число, если возможно; иначе возвращает как есть."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return v
    s = str(v).strip()
    if not s:
        return v
    try:
        n = float(s.replace(",", "."))
        return int(n) if n == int(n) else n
    except (ValueError, TypeError):
        return v


def _safe_save_excel(df: pd.DataFrame, output_file: str, sheet_name: str = "Sheet1") -> str:
    # Пишем во временный файл и затем атомарно заменяем
    try:
        with atomic_output_path(output_file) as tmp_path:
            with excel_writer(tmp_path) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                if "Артикул" in df.columns:
                    col_idx = list(df.columns).index("Артикул") + 1
                    ws = writer.sheets[sheet_name]
                    for row in range(2, len(df) + 2):
                        ws.cell(row=row, column=col_idx).number_format = "0"
    except PermissionError as exc:
        raise RuntimeError(f"Файл занят другим процессом: {output_file}. Закройте его и повторите.") from exc
    return output_file

# 📊 Преобразуем данные в Excel
def to_excel(postings, date_from, date_to, month, year, output_file=None, session: Optional[requests.Session] = None):
    from datetime import datetime
    import pandas as pd
    session = session or create_session()

    rows = []
    total_posts = max(len(postings or []), 1)

    # Название месяца на русском в родительном падеже (Сентябрь → сентября)
    months = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]
    month_name = months[month-1]

    # Формируем путь и имя файла в папке ../reports относительно этого скрипта
    if not output_file:
        script_dir = os.path.dirname(__file__)
        reports_dir = os.path.abspath(os.path.join(script_dir, '..', 'reports'))
        _ensure_reports_dir_and_check_space(reports_dir)
        output_file = os.path.join(reports_dir, f"{month_name} {year}.xlsx")


    # карта себестоимости: ключ может быть точным offer_id или префиксом
    cost_map = load_cost_map()
    # Ozon's sale-with-commission accrual typically lands 2-19 days after shipment (median ~6,
    # verified against real August data) — not just for orders shipped at month-end. Postings stay
    # scoped to the calendar month (date_from/date_to below), but accruals are fetched with a lookahead
    # past month-end so a posting shipped this month isn't reported as "ожидает расчёта" just because
    # Ozon's own commission bookkeeping runs behind the shipment date. get_transactions_by_posting
    # already clamps to "today", so this is a no-op when the report is generated soon after month-end.
    ACCRUAL_LOOKAHEAD_DAYS = 21
    accrual_date_to = (isoparse(date_to) + timedelta(days=ACCRUAL_LOOKAHEAD_DAYS)).isoformat()
    transactions_by_posting = get_transactions_by_posting(date_from, accrual_date_to, session) if postings else {}

    for idx, post in enumerate(postings, start=1):
        posting_number = post.get("posting_number", "")
        status = str(post.get("status") or "").strip().lower()
        schema = post.get("__schema", "")
        
        # Дата отгрузки - для FBS используется shipment_date, для FBO может быть другое поле
        if schema == "FBO":
            # Для FBO заказов пробуем различные поля с датами (в порядке приоритета)
            date = (post.get("in_process_at") or 
                   post.get("shipment_date") or 
                   post.get("created_at") or 
                   post.get("date") or
                   post.get("in_process_at_date") or
                   post.get("shipment_date_time") or "")
        else:
            # Для FBS используем shipment_date
            date = post.get("shipment_date", "")
        
        # Оставляем только дату без времени (YYYY-MM-DD)
        if date and isinstance(date, str):
            if "T" in date:
                date = date.split("T")[0]
            elif " " in date:
                date = date.split(" ")[0]
        
        items = post.get("products", []) or []

        # Если в заказе нет товаров — пропускаем
        if not items:
            continue

        # Количество (сумма по позициям)
        quantity_total = sum(int(it.get("quantity", 0) or 0) for it in items)

        # Заголовок строки — первая позиция
        head = items[0]
        name = str(head.get("name", ""))
        # Все артикулы заказа (без дублей, в исходном порядке)
        seen = set()
        offer_ids_list = []
        for it in items:
            oid = str(it.get("offer_id", ""))
            if oid and oid not in seen:
                seen.add(oid)
                offer_ids_list.append(oid)
        offer_ids_joined = ", ".join(offer_ids_list)

        # Себестоимость (по всем товарам, со знаком минус)
        cost_price = 0.0
        for it in items:
            oid = str(it.get("offer_id", "") or "").strip()
            q = int(it.get("quantity", 0) or 0)

            # Совпадение по offer_id (ключ нормализован: 12345 и 12345.0 из Excel → один ключ)
            oid_norm = _normalize_articul_key(oid)
            unit_cost = float(cost_map.get(oid_norm, 0) or 0) if oid_norm else 0.0
            cost_price -= unit_cost * q

        # Суммируем итог каждого начисления один раз; детали не прибавляем повторно.
        amount = 0.0
        sale_commission = 0.0
        price = 0.0

        transactions = transactions_by_posting.get(posting_number, [])
        for trans in transactions or []:
            amount += float(trans.get("amount") or 0)
            sale_commission += float(trans.get("sale_commission") or 0)
            price += float(trans.get("accruals_for_sale") or 0)
        has_commission_data = any(trans.get("has_commission_data") for trans in transactions or [])
        was_returned = any(trans.get("is_return") for trans in transactions or [])
        # Cross-border postings (Ozon settles with the seller on customs/currency-conversion timing,
        # not the usual schedule) can sit with commission_amount/payout at 0 in Ozon's own live
        # snapshot for a long time — this is not a request-window gap like the case below.
        is_pending_buyout = not has_commission_data and any(it.get("is_marketplace_buyout") for it in items)
        PENDING_BUYOUT_LABEL = "Ожидает расчёта Ozon (трансграничный заказ)"
        PENDING_MONTH_GAP_LABEL = "Ожидает расчёта Ozon (начисление ещё не пришло)"

        if price == 0 and not has_commission_data:
            # Accruals are split per service/day by /v1/finance/accrual/by-day; a posting whose
            # sale-with-commission accrual fell in a different month (late-arriving delivery/return
            # fee corrections, common near month boundaries) has none with a seller_price this month.
            # The posting itself always carries the per-unit retail price regardless of accrual timing.
            # Skipped when has_commission_data is True: there price==0 means a sale was fully offset
            # by its own return within the period (see was_returned below) — a genuine zero, not a gap.
            price = sum(float(it.get("price") or 0) * int(it.get("quantity", 0) or 0) for it in items)

        # Формируем значения в зависимости от статуса
        if status == "delivering":
            amount_cell = amount
            sale_commission_cell = "-"
            delivery_cost_cell = "-"
            profit_cell = "-"
            cost_price = 0.0   # себестоимость 0 — заказ ещё в доставке
        elif status == "awaiting_packaging":
            amount_cell = "-"
            sale_commission_cell = "-"
            delivery_cost_cell = "-"
            profit_cell = "-"
            cost_price = 0.0   # себестоимость 0 — заказ ожидает сборки
        elif status == "cancelled":
            amount_cell = amount
            sale_commission_cell = "-"
            delivery_cost_cell = "-"
            profit_cell = amount
            cost_price = 0.0   # ← себестоимость обнуляем при отмене
        elif status == "delivered" and was_returned:
            # Ozon never updates posting status for a post-delivery return — it stays "delivered"
            # forever, with the return visible only in the accruals (see was_returned above). Report
            # it as "returned" so the cost of goods isn't charged for stock that came back, and so it
            # counts correctly in calc_business_indicators()'s per-status totals below.
            status = "returned"
            amount_cell = amount
            sale_commission_cell = sale_commission
            delivery_cost_cell = -amount + price + sale_commission
            cost_price = 0.0
            profit_cell = amount
        elif status == "delivered" and not has_commission_data:
            # No sale-with-commission accrual for this posting fell inside the requested period —
            # either it's simply dated into a different month (common near month boundaries: shipped
            # late in the month, accrual posts a few days into the next one — see the real example
            # 0250581479-0070-1, shipped 2026-08-20, accrual posted 2026-09-03), or it's a cross-border
            # posting (is_marketplace_buyout) that Ozon settles on its own customs/currency timing and
            # which can show commission_amount=payout=0 in Ozon's own live snapshot for a long time.
            # Either way we do not yet know the real revenue, so profit must not be computed from a
            # payout that hasn't happened, and cost of goods must not be charged before revenue is
            # recognized (matching principle) — reported instead as its own status so it's easy to
            # find and re-check once Ozon actually posts the accrual (it will show up as a normal
            # "delivered" row in whatever month that turns out to be).
            status = "ожидает расчёта"
            amount_cell = amount
            sale_commission_cell = PENDING_BUYOUT_LABEL if is_pending_buyout else PENDING_MONTH_GAP_LABEL
            delivery_cost_cell = sale_commission_cell
            cost_price = 0.0
            profit_cell = sale_commission_cell
        elif status == "delivered":
            amount_cell = amount
            sale_commission_cell = sale_commission
            delivery_cost_cell = - amount + price + sale_commission
            profit_cell = amount + cost_price
        elif status == "returned":
            amount_cell = amount
            sale_commission_cell = sale_commission
            delivery_cost_cell = -amount + price + sale_commission
            cost_price = 0.0
            profit_cell = amount
        else:
            amount_cell = "-"
            sale_commission_cell = "-"
            delivery_cost_cell = "-"
            profit_cell = "-"
            cost_price = 0.0  # Себестоимость списываем только для доставленной продажи.

        artikul_val = _artikul_to_number(offer_ids_joined) if len(offer_ids_list) == 1 else offer_ids_joined
        rows.append({
            "Статус": status,
            "Номер заказа": posting_number,
            "Название товара": name,
            "Артикул": artikul_val,
            "Количество шт.": quantity_total,
            "Цена продажи": price,
            "Комиссия за продажу Ozon": sale_commission_cell,
            "Логистика (Включает операционные ошибки продавца)": delivery_cost_cell,
            "Сумма начисления": amount_cell,
            "Себестоимость": cost_price,
            "Прибыль": profit_cell,
            "Дата отгрузки": date,
            "Схема": post.get("__schema", "")
        })

        # выводим прогресс каждые 5 записей и на финише
        if idx % 5 == 0 or idx == total_posts:
            percent = int(idx * 100 / total_posts)
            print(f"\r⚙️ Обработка заказов: {percent}%", end="", flush=True)

    df = pd.DataFrame(rows)
    if "Артикул" in df.columns:
        df["Артикул"] = df["Артикул"].apply(_artikul_to_number)
    output_file = _safe_save_excel(df, output_file, sheet_name="Заказы")
    print("\r✅ Обработка заказов: 100%")
    print(f"✅ Отчёт сохранён: {output_file}")
    return output_file

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

def create_campaigns_sheet(filename: str, session: Optional[requests.Session] = None,
                           date_from: Optional[str] = None, date_to: Optional[str] = None):
    """
    Создаёт лист Excel с данными обо всех рекламных кампаниях за период (активные и неактивные).
    """
    with ExitStack() as books:
        if not session or not date_from or not date_to:
            return
    
        print("📊 Получаем данные о рекламных кампаниях за период...")
    
        campaigns_data = get_campaigns_data_for_excel(session, date_from, date_to)
    
        if campaigns_data is None:
            print("ℹ️ Не настроены переменные для Performance API. Пропускаем создание листа кампаний.")
            return
    
        if not campaigns_data:
            print("ℹ️ Не найдено кампаний за указанный период.")
            return
    
        try:
            # Открываем Excel-файл
            wb = books.enter_context(closing(load_workbook(filename)))
        
            # Удаляем лист "Кампании", если он уже существует
            if "Кампании" in wb.sheetnames:
                wb.remove(wb["Кампании"])
        
            # Создаём новый лист
            ws_campaigns = wb.create_sheet("Кампании")
        
            # Заголовки столбцов
            headers = [
                "ID кампании", "Название кампании", "Состояние", "Тип оплаты", "Тип объекта",
                "Бюджет (руб.)", "Дневной бюджет (руб.)", "Недельный бюджет (руб.)",
                "Расход за период (руб.)", "Показы", "Клики", "CTR (%)",
                "Средняя цена клика (руб.)", "Заказы (шт.)", "Заказы (руб.)", "ДРР (%)"
            ]
        
            # Записываем заголовки
            for col_idx, header in enumerate(headers, start=1):
                cell = ws_campaigns.cell(row=1, column=col_idx)
                cell.value = header
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.font = Font(bold=True, color="FFFFFF")
        
            # Записываем данные
            for row_idx, campaign in enumerate(campaigns_data, start=2):
                for col_idx, header in enumerate(headers, start=1):
                    cell = ws_campaigns.cell(row=row_idx, column=col_idx)
                    value = campaign.get(header, "")
                
                    # Форматируем числовые значения
                    if isinstance(value, (int, float)):
                        cell.value = value
                        if "руб." in header or "ДРР" in header or "CTR" in header:
                            cell.number_format = "#,##0.00"
                        elif "Показы" in header or "Клики" in header or "Заказы (шт.)" in header:
                            cell.number_format = "#,##0"
                    else:
                        cell.value = value
                
                    cell.alignment = Alignment(horizontal="left", vertical="center")
        
            # Автоматически подбираем ширину столбцов
            for col_idx, header in enumerate(headers, start=1):
                max_length = len(str(header))
                for row in ws_campaigns.iter_rows(min_row=2, max_row=ws_campaigns.max_row, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                ws_campaigns.column_dimensions[ws_campaigns.cell(row=1, column=col_idx).column_letter].width = min(max_length + 2, 50)
        
            # Сохраняем изменения
            save_workbook_atomic(wb, filename)
            wb.close()
            print(f"✅ Лист 'Кампании' создан: {len(campaigns_data)} кампаний")
        
        except Exception as e:
            print(f"⚠️ Ошибка при создании листа кампаний: {str(e)}")


def calc_business_indicators(
    filename,
    session: Optional[requests.Session] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    ozon_promotion_cost_override: Optional[float] = None,
    external_marketing_cost_override: Optional[float] = None,
):
    with ExitStack() as books:
        print("💲 Рассчёт бизнес показателей")
    
        # Пытаемся получить затраты на продвижение Ozon из Performance API
        ozon_promotion_cost = 0.0
        if ozon_promotion_cost_override is not None:
            ozon_promotion_cost = abs(float(ozon_promotion_cost_override))
            print(f"💰 Затраты на продвижение Ozon заданы вручную: {ozon_promotion_cost:.2f} ₽")
        elif session and date_from and date_to:
            perf_stats = get_cpc_campaigns_for_month(session, date_from, date_to)
            ozon_promotion_cost = perf_stats.get("total_cost", 0.0)
            if ozon_promotion_cost > 0:
                print(f"💰 Затраты на продвижение Ozon (CPC) из API: {ozon_promotion_cost:.2f} ₽")
    
        # Если не получилось из API или сумма 0 - спрашиваем у пользователя только в интерактивном режиме
        ui_prompts_enabled = os.environ.get("OZONREPORTX_UI_PROMPTS") == "1"
        if ozon_promotion_cost == 0.0 and ozon_promotion_cost_override is None:
            if not sys.stdin.isatty() and not ui_prompts_enabled:
                print("ℹ️ Затраты на продвижение Ozon не переданы и не получены из API. Используем 0.")
                ozon_promotion_cost = 0.0
            else:
                print("Введите сумму затрат на продвижение Ozon за месяц (или Enter для 0):")
                try:
                    user_input = input().strip()
                    if user_input:
                        ozon_promotion_cost = abs(float(user_input.replace(",", ".")))
                    else:
                        ozon_promotion_cost = 0.0
                except ValueError:
                    print("❌ Некорректное число. Используем 0.")
                    ozon_promotion_cost = 0.0

        # Расходы по «Звёздным товарам» из отчёта о балансе (за месяц, с разбивкой по 30 дней)
        # В отчёте показываем со знаком плюс (как Продвижение Ozon)
        star_products_cost = 0.0
        fbo_storage_cost = 0.0
        payout_summary = None
        if date_from and date_to:
            try:
                # date_from вида "2025-02-01T00:00:00Z"
                parts = date_from.split("T")[0].split("-")
                if len(parts) == 3:
                    year = int(parts[0])
                    month = int(parts[1])
                    balance_reports = get_monthly_balance_reports(month, year)
                    raw = get_star_products_for_month(month, year, reports=balance_reports)
                    star_products_cost = abs(float(raw))
                    if star_products_cost > 0:
                        print(f"💰 Звёздные товары (из отчёта о балансе): {star_products_cost:.2f} ₽")
                    raw_storage = get_product_placement_in_ozon_warehouses_for_month(month, year, reports=balance_reports)
                    fbo_storage_cost = abs(float(raw_storage))
                    if fbo_storage_cost > 0:
                        print(f"💰 Расход хранения FBO (из отчёта о балансе): {fbo_storage_cost:.2f} ₽")
                    payout_summary = summarize_balance_month(month, year, reports=balance_reports)
                    if payout_summary:
                        print(f"💰 Ozon выплатил за период: {payout_summary['payments']:.2f} ₽")
            except Exception as e:
                raise RuntimeError(f"Не удалось получить данные из баланса: {e}") from e

        # Продвижение Ozon = CPC + Звёздные товары
        ozon_promotion_total = ozon_promotion_cost + star_products_cost

        # Запрашиваем затраты на внешний маркетинг (кампании не на Ozon)
        external_marketing_cost = 0.0
        if external_marketing_cost_override is not None:
            external_marketing_cost = abs(float(external_marketing_cost_override))
            print(f"💰 Внешний маркетинг задан вручную: {external_marketing_cost:.2f} ₽")
        elif not sys.stdin.isatty() and not ui_prompts_enabled:
            print("ℹ️ Внешний маркетинг не передан. Используем 0.")
            external_marketing_cost = 0.0
        else:
            print("Введите сумму затрат на внешний маркетинг за месяц (кампании не на Ozon, или Enter для 0):")
            try:
                user_input = input().strip()
                if user_input:
                    external_marketing_cost = abs(float(user_input.replace(",", ".")))
                else:
                    external_marketing_cost = 0.0
            except ValueError:
                print("❌ Некорректное число. Используем 0.")
                external_marketing_cost = 0.0
    
    
        # Открываем Excel-файл; лист «Заказы»: A=Статус, F=Цена продажи, G=Комиссия Ozon, H=Логистика, J=Себестоимость, K=Прибыль
        wb = books.enter_context(closing(load_workbook(filename)))
        ws = wb["Заказы"] if "Заказы" in wb.sheetnames else wb.active

        # Считаем Общую выручку, Чистую прибыль, Себестоимость
        sales_revenue = 0
        for cell in ws["F"][1:]:
            if isinstance(cell.value, (int, float)):
                sales_revenue += cell.value

        net_profit = 0
        for cell in ws["K"][1:]:
            if isinstance(cell.value, (int, float)):
                net_profit += cell.value

        cost_price = 0
        for cell in ws["J"][1:]:
            if isinstance(cell.value, (int, float)):
                cost_price += cell.value

        # Новые показатели по строкам заказов: статус, средний чек, отменённые/доставленные, средние доли комиссии и логистики
        total_orders = max(0, ws.max_row - 1)
        delivered_count = 0
        returned_count = 0
        delivering_count = 0
        cancelled_count = 0
        pending_count = 0
        ratios_commission_pct = []   # Комиссия Ozon / Цена продажи, %
        ratios_logistics_pct = []   # Логистика / Цена продажи, %
        commission_total = 0.0
        logistics_total = 0.0
        revenue_for_avg_check = 0.0
        orders_nonzero_price = 0
        our_amount_total = 0.0

        for row in range(2, ws.max_row + 1):
            status_val = ws.cell(row=row, column=1).value
            status = str(status_val).strip().lower() if status_val is not None else ""
            if status == "delivered":
                delivered_count += 1
            if status == "returned":
                returned_count += 1
            if status == "delivering":
                delivering_count += 1
            if status == "cancelled":
                cancelled_count += 1
            if status == "ожидает расчёта":
                pending_count += 1

            price_val = ws.cell(row=row, column=6).value
            comm_val = ws.cell(row=row, column=7).value
            log_val = ws.cell(row=row, column=8).value
            amount_val = ws.cell(row=row, column=9).value
            if isinstance(amount_val, (int, float)):
                our_amount_total += amount_val

            try:
                price = float(price_val) if price_val is not None and str(price_val).strip() not in ("-", "") else None
            except (TypeError, ValueError):
                price = None
            try:
                comm = float(comm_val) if comm_val is not None and str(comm_val).strip() not in ("-", "") else None
            except (TypeError, ValueError):
                comm = None
            try:
                log = float(log_val) if log_val is not None and str(log_val).strip() not in ("-", "") else None
            except (TypeError, ValueError):
                log = None

            if price is not None and price != 0:
                revenue_for_avg_check += price
                orders_nonzero_price += 1
                if comm is not None:
                    commission_total += abs(comm)
                    ratios_commission_pct.append(abs((comm / price) * 100))
                if log is not None:
                    logistics_total += abs(log)
                    ratios_logistics_pct.append((log / price) * 100)

        average_check = (revenue_for_avg_check / orders_nonzero_price) if orders_nonzero_price > 0 else 0
        avg_commission_pct = (sum(ratios_commission_pct) / len(ratios_commission_pct)) if ratios_commission_pct else 0
        avg_logistics_pct = (sum(ratios_logistics_pct) / len(ratios_logistics_pct)) if ratios_logistics_pct else 0

        # Вычитаем затраты на продвижение Ozon (CPC + Звёздные товары) и внешний маркетинг из чистой прибыли
        total_marketing_cost = ozon_promotion_total + external_marketing_cost
        net_profit = net_profit - total_marketing_cost - fbo_storage_cost
        net_profit_margin = (net_profit / sales_revenue) * 100 if sales_revenue > 0 else 0
        cogs = sales_revenue + cost_price
        gross_profit_margin = (cogs / sales_revenue) * 100 if sales_revenue > 0 else 0
        operating_expenses = cogs - net_profit

        # Записываем результат
        ws["P1"] = "Общая выручка"
        ws["Q1"] = sales_revenue
        ws["P2"] = "Чистая прибыль"
        ws["Q2"] = net_profit
        ws["P3"] = "Итоговая себестоимость"
        ws["Q3"] = cost_price
        ws["P4"] = "Рентабельность по чистой прибыли (Net Profit Margin) %"
        ws["Q4"] = net_profit_margin
        ws["P5"] = "COGS (валовая прибыль)"
        ws["Q5"] = cogs
        ws["P6"] = "Gross Profit Margin Рентабельность по валовой прибыли %"
        ws["Q6"] = gross_profit_margin
        ws["P7"] = "Операционные расходы"
        ws["Q7"] = operating_expenses
        ws["P8"] = "Продвижение Ozon"
        ws["Q8"] = ozon_promotion_total
        ws["P9"] = "Звёздные товары"
        ws["Q9"] = star_products_cost
        ws["P10"] = "Внешний маркетинг"
        ws["Q10"] = external_marketing_cost

        ws["P11"] = "Средний чек"
        ws["Q11"] = average_check
        ws["P12"] = "Общее количество заказов"
        ws["Q12"] = total_orders
        ws["P13"] = "Количество отменённых заказов"
        ws["Q13"] = cancelled_count
        ws["P14"] = "Количество доставленных заказов"
        ws["Q14"] = delivered_count
        ws["P15"] = "Комиссии Ozon %"
        ws["Q15"] = avg_commission_pct
        ws["P16"] = "Логистика %"
        ws["Q16"] = avg_logistics_pct
        ws["P17"] = "Расход хранения FBO"
        ws["Q17"] = fbo_storage_cost
        ws["P18"] = "Количество возвращённых заказов"
        ws["Q18"] = returned_count
        ws["P19"] = "Количество заказов в доставке"
        ws["Q19"] = delivering_count
        ws["P20"] = "Комиссии Ozon сумма"
        ws["Q20"] = commission_total
        ws["P21"] = "Логистика сумма"
        ws["Q21"] = logistics_total
        ws["P22"] = "Заказы, ожидающие расчёта Ozon (не учтены в прибыли/себестоимости)"
        ws["Q22"] = pending_count

        row_cursor = 23
        if payout_summary:
            pairs = [
                ("Выплачено Ozon за период (реальный перевод)", payout_summary["payments"]),
                ("Входящий баланс на начало периода", payout_summary["opening_balance"]),
                ("Исходящий баланс на конец периода", payout_summary["closing_balance"]),
                ("Начислено Ozon за период (по балансу, календарные дни)", payout_summary["accrued"]),
                ("Сумма начисления по нашему расчёту (по датам отгрузки)", our_amount_total),
                ("Расхождение: баланс Ozon минус наш расчёт", payout_summary["accrued"] - our_amount_total),
                ("Комиссия за ранний вывод средств", payout_summary["early_payment_fee"]),
            ]
            for label, value in pairs:
                ws.cell(row=row_cursor, column=16, value=label)
                ws.cell(row=row_cursor, column=17, value=value if value is not None else "—")
                row_cursor += 1
            ws.cell(row=row_cursor, column=16, value="Справка")
            ws.cell(row=row_cursor, column=17,
                   value="«Начислено Ozon» считает по календарным дням месяца из отчёта о балансе; наш расчёт — "
                         "по заказам, отгруженным в этом месяце (с доборкой начислений задним числом). "
                         "Небольшое расхождение — это нормально: разная методика группировки, не ошибка. "
                         "Большое расхождение стоит проверить — смотрите строку «Заказы, ожидающие расчёта Ozon» выше.")
            row_cursor += 2
            other_services = {name: amount for name, amount in payout_summary["services"].items()
                              if name not in ("star_products", "product_placement_in_ozon_warehouses") and amount}
            if other_services:
                ws.cell(row=row_cursor, column=16, value="Прочие комиссии и услуги Ozon за период (баланс)")
                row_cursor += 1
                for name, amount in sorted(other_services.items(), key=lambda item: abs(item[1]), reverse=True)[:12]:
                    ws.cell(row=row_cursor, column=16, value=name)
                    ws.cell(row=row_cursor, column=17, value=amount)
                    row_cursor += 1

        # Сохраняем изменения
        save_workbook_atomic(wb, filename)
        wb.close()
        print(f"✅ Бизнес показатели добавлены в отчёт")
    
        # Создаём лист с данными о кампаниях
        create_campaigns_sheet(filename, session=session, date_from=date_from, date_to=date_to)

# 🚀 Точка входа
def date_range_for_month(month: int, year: int):
    """Возвращает (date_from, date_to) в формате API для заданных месяца и года."""
    from calendar import monthrange
    first_day = datetime(year, month, 1)
    last_day = datetime(year, month, monthrange(year, month)[1])
    return first_day.strftime('%Y-%m-%dT00:00:00Z'), last_day.strftime('%Y-%m-%dT23:59:59Z')


def main(argv=None):
    import argparse
    parser = argparse.ArgumentParser(description="Месячный отчёт по продажам Ozon.")
    parser.add_argument("--month", type=int, default=None, help="Номер месяца (1–12), для неинтерактивного запуска")
    parser.add_argument("--year", type=int, default=None, help="Год (например 2025), для неинтерактивного запуска")
    parser.add_argument(
        "--ozon-promotion-cost",
        type=float,
        default=None,
        help="Ручные затраты на продвижение Ozon за месяц. Если не передано, сначала используется Performance API.",
    )
    parser.add_argument(
        "--external-marketing-cost",
        type=float,
        default=None,
        help="Ручные затраты на внешний маркетинг за месяц. Если не передано в неинтерактивном режиме, используется 0.",
    )
    args = parser.parse_args(argv)

    if args.month is not None and args.year is not None:
        if not (1 <= args.month <= 12 and 2000 <= args.year <= 2100):
            raise ValueError("Укажите месяц 1–12 и год 2000–2100")
        month, year = args.month, args.year
        date_from, date_to = date_range_for_month(month, year)
    else:
        date_from, date_to, month, year = get_custom_date_range()

    print("📦 Получаем список заказов за месяц...")
    with create_session() as session:
        fbs_orders = get_orders(date_from, date_to, session=session)
        fbo_orders = get_fbo_orders(date_from, date_to, session=session)

        all_orders = fbs_orders + fbo_orders
        print(f"🔢 Найдено заказов: {len(all_orders)}")

        reports_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'reports'))
        _ensure_reports_dir_and_check_space(reports_dir)
        months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
        output_file = os.path.join(reports_dir, f"{months[month - 1]} {year}.xlsx")
        start_ts = time.time()
        # Publish only after orders, financial metrics and workbook writes succeed.
        with atomic_output_path(output_file) as temporary:
            to_excel(all_orders, date_from, date_to, month, year, output_file=str(temporary), session=session)
            calc_business_indicators(
                str(temporary),
                session=session,
                date_from=date_from,
                date_to=date_to,
                ozon_promotion_cost_override=args.ozon_promotion_cost,
                external_marketing_cost_override=args.external_marketing_cost,
            )
        duration_s = time.time() - start_ts
        print(f"✅ Отчёт сохранён: {output_file}")
    # Краткий итог
    print(f"⏱ Время формирования: {duration_s:.1f} с")


if __name__ == "__main__":
    try:
        main()
    except (RuntimeError, requests.RequestException) as exc:
        print(f"❌ {exc}")
        raise SystemExit(1)
