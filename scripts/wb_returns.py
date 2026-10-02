"""WB Returns and Item Movement Report.

This is a logistics report (why an item was returned and where it currently is on
its way back), not a financial one — it carries no monetary amount. Source:
dev.wildberries.ru/en/docs/openapi/reports#tag/mainReports/paths/~1api~1v1~1analytics~1goods-return.
"""
from collections import Counter
from datetime import date, timedelta
import threading
import time

import requests

BASE = "https://seller-analytics-api.wildberries.ru/api/v1/analytics/goods-return"
MAX_LOOKBACK_DAYS = 31
_request_lock = threading.Lock()
_next_request = 0.0


class ReturnsError(RuntimeError):
    pass


class ReturnsClient:
    def __init__(self, token, cancel=None, progress=lambda message: None, session=None):
        if not token:
            raise ReturnsError("Добавьте API-токен WB с доступом к категории «Аналитика» в настройках.")
        self.token, self.cancel, self.progress = token, cancel or threading.Event(), progress
        self.session = session or requests.Session()

    def check_cancel(self):
        if self.cancel.is_set():
            raise ReturnsError("Загрузка отчёта о возвратах WB отменена.")

    def fetch(self, days=MAX_LOOKBACK_DAYS):
        """The report only covers the last 31 days — a rolling window, not a calendar month."""
        global _next_request
        days = min(days, MAX_LOOKBACK_DAYS)
        date_to = date.today()
        date_from = date_to - timedelta(days=days)
        while not _request_lock.acquire(timeout=0.2):
            self.check_cancel()
        try:
            for _attempt in range(3):
                self.check_cancel()
                if _next_request > time.monotonic():
                    self.progress("Ожидание лимита WB: не чаще одного запроса возвратов в минуту…")
                while time.monotonic() < _next_request:
                    self.cancel.wait(min(0.25, _next_request - time.monotonic()))
                    self.check_cancel()
                _next_request = time.monotonic() + 60
                self.progress("Загрузка отчёта о возвратах WB…")
                try:
                    response = self.session.get(
                        BASE, params={"dateFrom": date_from.isoformat(), "dateTo": date_to.isoformat()},
                        headers={"Authorization": self.token}, timeout=(10, 60))
                except requests.RequestException:
                    raise ReturnsError("Не удалось связаться с WB. Проверьте соединение и повторите загрузку.") from None
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
                               403: "У токена WB нет доступа к категории «Аналитика» (отчёт о возвратах)."}.get(
                                   response.status_code, f"WB вернул HTTP {response.status_code}. Возвраты не обновлены.")
                    raise ReturnsError(message)
                try:
                    payload = response.json()
                except ValueError:
                    raise ReturnsError("WB вернул некорректный JSON.") from None
                records = payload.get("report") if isinstance(payload, dict) else None
                if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
                    raise ReturnsError("Некорректный формат отчёта о возвратах WB.")
                self.check_cancel()
                from datetime import datetime
                return {"version": 1, "fetched_at": datetime.now().isoformat(timespec="seconds"),
                        "period_days": days, "records": records}
            raise ReturnsError("WB ограничил запросы отчёта о возвратах. Повторите позже.")
        finally:
            _request_lock.release()


def summarize(data):
    records = data.get("records", [])
    return {
        "total": len(records),
        "by_status": dict(Counter(str(row.get("status") or "не указан") for row in records)),
        "by_reason": dict(Counter(str(row.get("reason") or "не указана") for row in records)),
        "by_type": dict(Counter(str(row.get("returnType") or "не указан") for row in records)),
        "by_srid": {row["srid"]: row for row in records if row.get("srid")},
    }


def top(counter, limit=5):
    return sorted(counter.items(), key=lambda item: item[1], reverse=True)[:limit]
