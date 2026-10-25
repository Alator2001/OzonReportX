"""Pacing and bounded 429 recovery for read-only Seller API report requests."""
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import threading
import time
from urllib.parse import urlsplit
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


def retry_after_seconds(value):
    if value is None:
        return None
    try:
        return max(0.0, float(value))
    except (TypeError, ValueError):
        try:
            date = parsedate_to_datetime(value)
            if date.tzinfo is None:
                date = date.replace(tzinfo=timezone.utc)
            return max(0.0, (date - datetime.now(timezone.utc)).total_seconds())
        except (TypeError, ValueError, OverflowError):
            return None


class ReportSession(requests.Session):
    def __init__(self, min_interval=0.5, max_retries=5, max_wait=300):
        super().__init__()
        self.min_interval = min_interval
        self.max_rate_retries = max_retries
        self.max_rate_wait = max_wait
        self._next_request = 0.0
        self._request_lock = threading.Lock()

    def request(self, method, url, **kwargs):
        if urlsplit(url).hostname != "api-seller.ozon.ru":
            return super().request(method, url, **kwargs)
        with self._request_lock:
            waited = 0.0
            for attempt in range(self.max_rate_retries + 1):
                pause = max(0.0, self._next_request - time.monotonic())
                if pause:
                    time.sleep(pause)
                try:
                    response = super().request(method, url, **kwargs)
                finally:
                    self._next_request = time.monotonic() + self.min_interval
                if response.status_code != 429:
                    return response
                delay = max(5.0 * (2 ** attempt), retry_after_seconds(response.headers.get("Retry-After")) or 0.0)
                try:
                    detail = str(response.json().get("message") or "")
                except (ValueError, AttributeError):
                    detail = ""
                response.close()
                if attempt == self.max_rate_retries or waited + delay > self.max_rate_wait:
                    method_path = urlsplit(url).path
                    detail = f" Ответ Ozon: «{detail}»." if detail else ""
                    raise RuntimeError(
                        f"Ozon продолжает ограничивать запросы к {method_path} (HTTP 429).{detail} "
                        "Если ошибка повторяется, лимит этого метода, скорее всего, расходует другая программа "
                        "или сервис с тем же Client-Id. Повторите формирование отчёта позже; неполный отчёт не сохранён.")
                print(f"⏳ Ozon ограничил частоту запросов (429). Пауза {delay:.0f} с, повтор {attempt + 1}/{self.max_rate_retries}.", flush=True)
                time.sleep(delay)
                waited += delay
                self.min_interval = min(max(self.min_interval * 2, 1.0), 5.0)
                self._next_request = time.monotonic()


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
