"""WB Sales Funnel report — organic views/cart/orders/buyouts per item (not advertising).

Accepts any past date range (verified live against dev.wildberries.ru's
/api/analytics/v3/sales-funnel/products), so it can be tied to the selected calendar
month like the financial report.
"""
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
import threading
import time

import requests

BASE = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"
PAGE_LIMIT = 1000
_request_lock = threading.Lock()
_next_request = 0.0


class FunnelError(RuntimeError):
    pass


class FunnelClient:
    def __init__(self, token, cancel=None, progress=lambda message: None, session=None):
        if not token:
            raise FunnelError("Добавьте API-токен WB с доступом к категории «Аналитика» в настройках.")
        self.token, self.cancel, self.progress = token, cancel or threading.Event(), progress
        self.session = session or requests.Session()

    def check_cancel(self):
        if self.cancel.is_set():
            raise FunnelError("Загрузка воронки продаж WB отменена.")

    def request(self, payload):
        global _next_request
        # WB allows 3 requests per 20s for this endpoint; space calls out to stay well within that.
        while not _request_lock.acquire(timeout=0.2):
            self.check_cancel()
        try:
            for _attempt in range(3):
                self.check_cancel()
                if _next_request > time.monotonic():
                    self.progress("Ожидание лимита WB (воронка продаж)…")
                while time.monotonic() < _next_request:
                    self.cancel.wait(min(0.25, _next_request - time.monotonic()))
                    self.check_cancel()
                _next_request = time.monotonic() + 7.0
                try:
                    response = self.session.post(BASE, json=payload,
                                                 headers={"Authorization": self.token}, timeout=(10, 60))
                except requests.RequestException:
                    raise FunnelError("Не удалось связаться с WB. Проверьте соединение и повторите загрузку.") from None
                if response.status_code == 429:
                    delays = [20.0]
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
                               403: "У токена WB нет доступа к категории «Аналитика» (воронка продаж)."}.get(
                                   response.status_code, f"WB вернул HTTP {response.status_code}. Воронка не обновлена.")
                    raise FunnelError(message)
                try:
                    payload_json = response.json()
                except ValueError:
                    raise FunnelError("WB вернул некорректный JSON.") from None
                body = payload_json.get("data") if isinstance(payload_json, dict) else None
                if not isinstance(body, dict) or not isinstance(body.get("products"), list):
                    raise FunnelError("Некорректный формат ответа воронки продаж WB.")
                return body
            raise FunnelError("WB ограничил запросы воронки продаж. Повторите позже.")
        finally:
            _request_lock.release()

    def fetch(self, date_from, date_to):
        period = {"start": date_from.isoformat(), "end": date_to.isoformat()}
        products, offset = [], 0
        while True:
            self.progress(f"Загрузка воронки продаж WB: получено {len(products)} товаров")
            body = self.request({"selectedPeriod": period, "nmIds": [], "brandNames": [],
                                  "subjectIds": [], "tagIds": [], "limit": PAGE_LIMIT, "offset": offset})
            page = body["products"]
            products.extend(page)
            if len(page) < PAGE_LIMIT:
                break
            offset += PAGE_LIMIT
        self.check_cancel()
        return {"version": 1, "fetched_at": datetime.now().isoformat(timespec="seconds"),
                "date_from": date_from.isoformat(), "date_to": date_to.isoformat(), "products": products}


def _row(entry):
    product = entry.get("product") or {}
    selected = ((entry.get("statistic") or {}).get("selected")) or {}
    conversions = selected.get("conversions") or {}
    return {
        "nmId": str(product.get("nmId") or ""),
        "title": str(product.get("title") or ""),
        "vendor_code": str(product.get("vendorCode") or ""),
        "subject": str(product.get("subjectName") or ""),
        "views": int(selected.get("openCount") or 0),
        "cart": int(selected.get("cartCount") or 0),
        "orders": int(selected.get("orderCount") or 0),
        "order_sum": Decimal(str(selected.get("orderSum", 0))),
        "buyouts": int(selected.get("buyoutCount") or 0),
        "buyout_sum": Decimal(str(selected.get("buyoutSum", 0))),
        "cancels": int(selected.get("cancelCount") or 0),
        "add_to_cart_pct": conversions.get("addToCartPercent"),
        "cart_to_order_pct": conversions.get("cartToOrderPercent"),
        "buyout_pct": conversions.get("buyoutPercent"),
    }


def summarize(data):
    rows = [_row(entry) for entry in data.get("products", [])]
    total_views = sum(row["views"] for row in rows)
    total_cart = sum(row["cart"] for row in rows)
    total_orders = sum(row["orders"] for row in rows)
    total_buyouts = sum(row["buyouts"] for row in rows)
    return {
        "products": len(rows), "by_nm": {row["nmId"]: row for row in rows if row["nmId"]},
        "views": total_views, "cart": total_cart, "orders": total_orders, "buyouts": total_buyouts,
        "add_to_cart_pct": round(total_cart / total_views * 100, 1) if total_views else None,
        "cart_to_order_pct": round(total_orders / total_cart * 100, 1) if total_cart else None,
        "buyout_pct": round(total_buyouts / total_orders * 100, 1) if total_orders else None,
    }
