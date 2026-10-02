"""Read-only WB finance integration. Source: dev.wildberries.ru/docs/openapi/financial-reports-and-accounting."""
import calendar
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import threading
import time

from openpyxl import Workbook, load_workbook
import requests

from scripts.file_io import atomic_output_path, save_workbook_atomic
from scripts.file_lock import exclusive_file
from scripts.marketplace_settings import read_settings

BASE = "https://finance-api.wildberries.ru/api/finance/v1/sales-reports"
_request_lock = threading.Lock()
_next_request = 0.0


class FinanceError(RuntimeError):
    pass


def number(value):
    try:
        result = Decimal(str(value).replace(",", "."))
    except (InvalidOperation, ValueError):
        raise FinanceError("Некорректная сумма в данных WB или справочнике себестоимости.") from None
    if not result.is_finite():
        raise FinanceError("Сумма должна быть конечным числом.")
    return result


def identity(token):
    return hashlib.sha256(token.encode()).hexdigest()


def period_dates(year, month):
    return f"{year:04d}-{month:02d}-01", f"{year:04d}-{month:02d}-{calendar.monthrange(year, month)[1]:02d}"


class FinanceClient:
    def __init__(self, token, cancel=None, progress=lambda message: None, session=None):
        if not token:
            raise FinanceError("Добавьте API-токен WB с доступом к категории «Финансы» в настройках.")
        self.token, self.cancel, self.progress = token, cancel or threading.Event(), progress
        self.session = session or requests.Session()

    def check_cancel(self):
        if self.cancel.is_set():
            raise FinanceError("Загрузка WB отменена. Предыдущий отчёт сохранён.")

    def request(self, suffix, payload):
        global _next_request
        # Finance methods allow one request per minute. Serialize all instances.
        while not _request_lock.acquire(timeout=0.2):
            self.check_cancel()
        try:
            for attempt in range(3):
                self.check_cancel()
                if _next_request > time.monotonic():
                    self.progress("Ожидание лимита WB: не чаще одного финансового запроса в минуту…")
                while time.monotonic() < _next_request:
                    self.cancel.wait(min(0.25, _next_request - time.monotonic()))
                    self.check_cancel()
                _next_request = time.monotonic() + 60
                try:
                    response = self.session.post(BASE + suffix, json=payload,
                                                 headers={"Authorization": self.token}, timeout=(10, 60))
                except requests.RequestException:
                    raise FinanceError("Не удалось связаться с WB. Проверьте соединение и повторите загрузку.") from None
                if response.status_code == 204:
                    return None
                if response.status_code == 429:
                    delays = [60.0]
                    for key in ("X-Ratelimit-Retry", "X-Ratelimit-Reset", "Retry-After"):
                        try:
                            delay = float(response.headers.get(key, 0))
                            if 0 <= delay < float("inf"):
                                delays.append(delay)
                        except (ValueError, TypeError):
                            pass
                    _next_request = time.monotonic() + max(delays)
                    continue
                if response.status_code != 200:
                    message = {401: "Токен WB недействителен или истёк.",
                               403: "У токена WB нет доступа к финансовым отчётам."}.get(
                                   response.status_code, f"WB вернул HTTP {response.status_code}. Отчёт не сохранён.")
                    raise FinanceError(message)
                try:
                    data = response.json()
                except ValueError:
                    raise FinanceError("WB вернул некорректный JSON.") from None
                if not isinstance(data, list) or any(not isinstance(row, dict) for row in data):
                    raise FinanceError("Некорректный формат финансового отчёта WB.")
                return data
            raise FinanceError("WB ограничил запросы. Повторите позже; предыдущий отчёт сохранён.")
        finally:
            _request_lock.release()

    def fetch(self, year, month):
        first, last = period_dates(year, month)
        common = {"dateFrom": first, "dateTo": last + "T23:59:59", "period": "daily"}
        reports, report_ids = [], set()
        while True:
            self.progress(f"Загрузка итогов WB: получено {len(reports)} отчётов")
            page = self.request("/list", {**common, "limit": 1000, "offset": len(reports)})
            if not page:
                break
            for row in page:
                report_id = row.get("reportId")
                if not isinstance(report_id, int) or report_id in report_ids:
                    raise FinanceError("Повторный или некорректный ID отчёта WB.")
                if row.get("currency") != "RUB":
                    raise FinanceError("Сводка WB поддерживает только отчёты в RUB.")
                # Never silently include an overlapping weekly report in a calendar month.
                if not (first <= str(row.get("dateFrom", ""))[:10] <= str(row.get("dateTo", ""))[:10] <= last):
                    raise FinanceError("Период отчёта WB выходит за выбранный месяц. Требуется сверка периода.")
                report_ids.add(report_id)
                reports.append(row)
            if len(page) < 1000:
                break
        details, seen, cursor = [], set(), 0
        if reports:
            while True:
                self.progress(f"Загрузка детализации WB: {len(details)} строк")
                page = self.request("/detailed", {**common, "limit": 100000, "rrdId": cursor})
                if page is None:
                    break
                if not page:
                    raise FinanceError("WB вернул пустую страницу вместо завершения детализации (204).")
                for row in page:
                    row_id = row.get("rrdId")
                    if not isinstance(row_id, int) or row_id <= cursor or row_id in seen:
                        raise FinanceError("Повторная или некорректная страница детализации WB.")
                    if row.get("reportId") not in report_ids or row.get("currency") != "RUB":
                        raise FinanceError("Итоги и детализация WB не совпадают. Повторите загрузку позже.")
                    seen.add(row_id)
                    details.append(row)
                cursor = page[-1]["rrdId"]
            if {row["reportId"] for row in details} != report_ids:
                raise FinanceError("WB вернул неполную детализацию. Предыдущий отчёт сохранён.")
        self.check_cancel()
        return {"version": 1, "year": year, "month": month, "account": identity(self.token),
                "fetched_at": datetime.now().isoformat(timespec="seconds"), "reports": reports, "details": details}


def report_path(root, year, month):
    return Path(root) / "wb reports" / f"{year:04d}-{month:02d}.json"


def save_report(root, data):
    path = report_path(root, data["year"], data["month"])
    path.parent.mkdir(parents=True, exist_ok=True)
    with atomic_output_path(path) as temporary:
        temporary.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def load_report(root, year, month):
    path = report_path(root, year, month)
    if not path.exists():
        return None
    data = json.loads(path.read_text(encoding="utf-8"))
    token = read_settings(root).get("WB_API_TOKEN") or ""
    if data.get("account") != identity(token):
        return None
    if data.get("version") != 1 or data.get("year") != year or data.get("month") != month:
        raise FinanceError("Неподдерживаемый формат сохранённого отчёта WB.")
    return data


def ensure_costs(root, details=()):
    path = Path(root) / "wb_costs.xlsx"
    with exclusive_file(path):
        workbook = load_workbook(path) if path.exists() else Workbook()
        try:
            if not path.exists():
                sheet = workbook.active
                sheet.title = "Себестоимость WB"
                sheet.append(["nmId", "Артикул продавца", "Название", "Себестоимость, руб."])
                sheet.freeze_panes = "A2"
                for col, width in (("A", 18), ("B", 28), ("C", 55), ("D", 25)):
                    sheet.column_dimensions[col].width = width
            sheet = workbook["Себестоимость WB"]
            known = {str(row[0]) for row in sheet.iter_rows(min_row=2, values_only=True) if row[0] is not None}
            for row in details:
                nm_id = str(row.get("nmId") or "")
                if nm_id and nm_id != "0" and nm_id not in known:
                    sheet.append([nm_id, str(row.get("vendorCode") or ""), str(row.get("title") or ""), None])
                    for cell in sheet[sheet.max_row][:3]:
                        cell.data_type = "s"
                    known.add(nm_id)
            save_workbook_atomic(workbook, path)
        finally:
            workbook.close()
    return path


def read_costs(root):
    path = Path(root) / "wb_costs.xlsx"
    if not path.exists():
        return {}
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        result = {}
        for row in workbook["Себестоимость WB"].iter_rows(min_row=2, max_col=4, values_only=True):
            key, value = str(row[0] or "").strip(), row[3]
            if not key:
                continue
            if key in result:
                raise FinanceError(f"Повторный nmId {key} в wb_costs.xlsx.")
            cost = None if value is None or value == "" else number(value)
            if cost is not None and cost < 0:
                raise FinanceError("Себестоимость WB не может быть отрицательной.")
            result[key] = cost
        return result
    finally:
        workbook.close()


def summarize(data, costs):
    reports, rows = data["reports"], data["details"]
    if not reports:
        return {"empty": True, "notes": ["За период WB ещё не вернул финансовых отчётов."]}
    def total(field):
        return sum((number(row[field]) for row in reports), Decimal(0))
    revenue, payout = total("retailAmountSum"), total("bankPaymentSum")
    quantities = defaultdict(lambda: [Decimal(0), Decimal(0)])
    sales, returns = set(), set()
    missing, notes = set(), []
    sale_revenue = Decimal(0)
    for row in rows:
        operation = str(row.get("sellerOperName") or "").strip().lower()
        if operation not in ("продажа", "возврат"):
            continue
        nm_id = str(row.get("nmId") or "")
        order = str(row.get("srid") or "")
        if not order or not nm_id:
            raise FinanceError("В продаже/возврате WB отсутствует srid или nmId.")
        qty = number(row["quantity"])
        if qty < 0 or qty != qty.to_integral_value():
            raise FinanceError("Некорректное количество товара WB.")
        returning = operation == "возврат"
        quantities[(order, nm_id)][int(returning)] += qty
        (returns if returning else sales).add(order)
        if not returning:
            sale_revenue += number(row["retailAmount"])
    cogs = Decimal(0)
    unmatched_returns = False
    for (_order, nm_id), (sold, returned) in quantities.items():
        # A sale and its return within the report period have zero total COGS.
        if returned > sold:
            unmatched_returns = True
        net = max(Decimal(0), sold - returned)
        if net:
            cost = costs.get(nm_id)
            if cost is None:
                missing.add(nm_id)
            else:
                cogs += cost * net
    if missing:
        notes.append("Заполните себестоимость для nmId: " + ", ".join(sorted(missing)))
    if unmatched_returns:
        notes.append("Есть возвраты продаж другого периода. Прибыль не рассчитана: нужно согласовать восстановление себестоимости прошлых месяцев.")
    complete = not missing and not unmatched_returns
    profit = payout - cogs if complete else None
    gross = revenue - cogs if complete else None
    return {"empty": False, "revenue": revenue, "payout": payout,
            "costs": cogs if complete else None, "profit": profit,
            "margin": profit / revenue * 100 if profit is not None and revenue > 0 else None,
            "gross": gross, "operating": gross - profit if complete else None,
            "average": sale_revenue / len(sales) if sales else None,
            "sales": len(sales), "returns": len(returns),
            "logistics": total("deliveryServiceSum"), "storage": total("paidStorageSum"),
            "acceptance": total("paidAcceptanceSum"), "deductions": total("deductionSum"),
            "penalties": total("penaltySum"), "additional": total("additionalPaymentSum"),
            "notes": notes, "missing": sorted(missing)}


def order_rows(data, costs):
    """Per-srid rows for the monthly Excel export.

    Each detail row carries forPay/deliveryService/paidStorage/paidAcceptance/deduction/
    penalty/additionalPayment individually; summed across ALL rows sharing an srid this
    reproduces the report-level bankPaymentSum exactly, so it can be attributed per order.
    """
    orders, order_seq = {}, []
    for row in data.get("details", []):
        srid = str(row.get("srid") or "")
        if not srid:
            continue
        entry = orders.get(srid)
        if entry is None:
            entry = {"titles": {}, "vendors": [], "sold": defaultdict(lambda: Decimal(0)),
                      "returned": defaultdict(lambda: Decimal(0)), "sale_amount": Decimal(0),
                      "net_payout": Decimal(0), "date": None, "ops": set()}
            orders[srid] = entry
            order_seq.append(srid)
        entry["net_payout"] += (number(row.get("forPay", 0)) - number(row.get("deliveryService", 0))
                                 - number(row.get("paidStorage", 0)) - number(row.get("paidAcceptance", 0))
                                 - number(row.get("deduction", 0)) - number(row.get("penalty", 0))
                                 + number(row.get("additionalPayment", 0)))
        operation = str(row.get("sellerOperName") or "").strip().lower()
        if operation in ("продажа", "возврат"):
            entry["ops"].add(operation)
            nm_id = str(row.get("nmId") or "")
            qty = number(row["quantity"])
            title, vendor = str(row.get("title") or ""), str(row.get("vendorCode") or "")
            if title:
                entry["titles"].setdefault(nm_id, title)
            if vendor and vendor not in entry["vendors"]:
                entry["vendors"].append(vendor)
            if operation == "возврат":
                entry["returned"][nm_id] += qty
            else:
                entry["sold"][nm_id] += qty
                entry["sale_amount"] += number(row["retailAmount"])
            date = str(row.get("saleDt") or row.get("rrDate") or "")[:10]
            if date and (entry["date"] is None or date < entry["date"]):
                entry["date"] = date
    rows, residual = [], Decimal(0)
    for srid in order_seq:
        entry = orders[srid]
        if not entry["ops"]:
            # Logistics/storage rows for a sale settled in a different reporting period.
            residual += entry["net_payout"]
            continue
        nm_ids = set(entry["sold"]) | set(entry["returned"])
        quantity, cost_total, cost_known = Decimal(0), Decimal(0), True
        for nm_id in nm_ids:
            net = max(Decimal(0), entry["sold"][nm_id] - entry["returned"][nm_id])
            quantity += net
            if net:
                cost = costs.get(nm_id)
                if cost is None:
                    cost_known = False
                else:
                    cost_total += cost * net
        cost_value = -cost_total if cost_known else None
        profit = entry["net_payout"] + cost_value if cost_known else None
        status = "продажа и возврат" if len(entry["ops"]) > 1 else next(iter(entry["ops"]))
        rows.append({
            "Статус": status, "Номер заказа": srid,
            "Название товара": next(iter(entry["titles"].values()), ""),
            "Артикул": ", ".join(entry["vendors"]),
            "Количество шт.": float(quantity),
            "Цена продажи": float(entry["sale_amount"]),
            "Сумма начисления": float(entry["net_payout"]),
            "Себестоимость": float(cost_value) if cost_value is not None else None,
            "Прибыль": float(profit) if profit is not None else None,
            "Дата": entry["date"] or "",
        })
    return rows, residual
