"""WB Order Feed: near-real-time order status (FBS + FBO), replacing the deprecated
Orders/Sales statistics methods. Source: dev.wildberries.ru/en/docs/openapi/analytics#tag/orderFeed.
"""
from collections import Counter
from datetime import datetime, timedelta, timezone
import threading
import time

import requests

BASE = "https://seller-analytics-api.wildberries.ru/api/analytics/v1/order-feed"
MAX_LOOKBACK_DAYS = 31
PAGE_LIMIT = 1000
_request_lock = threading.Lock()
_next_request = 0.0


class OrderFeedError(RuntimeError):
    pass


class OrderFeedClient:
    def __init__(self, token, cancel=None, progress=lambda message: None, session=None):
        if not token:
            raise OrderFeedError("Добавьте API-токен WB с доступом к категории «Аналитика» в настройках.")
        self.token, self.cancel, self.progress = token, cancel or threading.Event(), progress
        self.session = session or requests.Session()

    def check_cancel(self):
        if self.cancel.is_set():
            raise OrderFeedError("Загрузка статусов заказов WB отменена.")

    def request(self, payload):
        global _next_request
        # Order Feed allows one request per minute. Serialize all instances.
        while not _request_lock.acquire(timeout=0.2):
            self.check_cancel()
        try:
            for _attempt in range(3):
                self.check_cancel()
                if _next_request > time.monotonic():
                    self.progress("Ожидание лимита WB: не чаще одного запроса статусов в минуту…")
                while time.monotonic() < _next_request:
                    self.cancel.wait(min(0.25, _next_request - time.monotonic()))
                    self.check_cancel()
                _next_request = time.monotonic() + 60
                try:
                    response = self.session.post(BASE, json=payload,
                                                 headers={"Authorization": self.token}, timeout=(10, 60))
                except requests.RequestException:
                    raise OrderFeedError("Не удалось связаться с WB. Проверьте соединение и повторите загрузку.") from None
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
                               403: "У токена WB нет доступа к категории «Аналитика» (Order Feed)."}.get(
                                   response.status_code, f"WB вернул HTTP {response.status_code}. Статусы не обновлены.")
                    raise OrderFeedError(message)
                try:
                    payload_json = response.json()
                except ValueError:
                    raise OrderFeedError("WB вернул некорректный JSON.") from None
                body = payload_json.get("data") if isinstance(payload_json, dict) else None
                if not isinstance(body, dict) or not isinstance(body.get("orders"), list):
                    raise OrderFeedError("Некорректный формат ответа Order Feed WB.")
                return body
            raise OrderFeedError("WB ограничил запросы статусов заказов. Повторите позже.")
        finally:
            _request_lock.release()

    def fetch(self, days=MAX_LOOKBACK_DAYS):
        """Order Feed only exposes the current status of orders touched in the last 31 days —
        it is a rolling snapshot, not a per-calendar-month report like the financial reports."""
        days = min(days, MAX_LOOKBACK_DAYS)
        end = datetime.now(timezone.utc).astimezone()
        start = end - timedelta(days=days)
        orders, snapshot, offset = [], None, 0
        while True:
            self.progress(f"Загрузка статусов заказов WB: получено {len(orders)}")
            pagination = {"offset": offset, "limit": PAGE_LIMIT}
            if snapshot:
                pagination["snapshotTime"] = snapshot
            body = self.request({
                "selectedPeriod": {"start": start.isoformat(), "end": end.isoformat()},
                "pagination": pagination,
            })
            snapshot = body.get("snapshotTime") or snapshot
            page = body["orders"]
            orders.extend(page)
            if len(page) < PAGE_LIMIT:
                break
            offset += PAGE_LIMIT
        self.check_cancel()
        return {"version": 1, "fetched_at": datetime.now().isoformat(timespec="seconds"),
                "period_days": days, "orders": orders}


def _bucket(orders):
    counts = Counter(str(row.get("status") or "") for row in orders)
    return {"total": len(orders), "in_progress": counts.get("created", 0),
            "delivered": counts.get("buyout", 0), "cancelled": counts.get("cancel", 0)}


def summarize(data):
    """created = в пути/в обработке, buyout = доставлено и выкуплено, cancel = отменено."""
    orders = data.get("orders", [])
    fbs = [row for row in orders if row.get("isMp")]
    fbo = [row for row in orders if not row.get("isMp")]
    overall = _bucket(orders)
    overall["fbs"] = _bucket(fbs)
    overall["fbo"] = _bucket(fbo)
    return overall
