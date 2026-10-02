"""Shared client for WB's create-task / poll-status / download report flow.

Acceptance Expenses and Paid Storage both work this way: GET a create endpoint to get a
taskId, poll a status endpoint until it reports "done", then GET the download endpoint
(204 means the report has no rows for the period). Each endpoint has its own WB rate
limit bucket, tracked here by a string key.
"""
import threading
import time

import requests

_lock = threading.Lock()
_next_request = {}


class ReportTaskError(RuntimeError):
    pass


def _get(session, cancel, progress, url, key, *, params=None, headers=None, min_interval=60.0):
    global _next_request
    while not _lock.acquire(timeout=0.2):
        if cancel.is_set():
            raise ReportTaskError("Загрузка отчёта WB отменена.")
    try:
        for _attempt in range(3):
            if cancel.is_set():
                raise ReportTaskError("Загрузка отчёта WB отменена.")
            next_time = _next_request.get(key, 0.0)
            if next_time > time.monotonic():
                progress("Ожидание лимита WB…")
            while time.monotonic() < next_time:
                cancel.wait(min(0.25, next_time - time.monotonic()))
                if cancel.is_set():
                    raise ReportTaskError("Загрузка отчёта WB отменена.")
            _next_request[key] = time.monotonic() + min_interval
            try:
                response = session.get(url, params=params, headers=headers, timeout=(10, 60))
            except requests.RequestException:
                raise ReportTaskError("Не удалось связаться с WB. Проверьте соединение и повторите загрузку.") from None
            if response.status_code == 429:
                delays = [min_interval]
                for header_key in ("X-Ratelimit-Retry", "X-Ratelimit-Reset", "Retry-After"):
                    try:
                        delay = float(response.headers.get(header_key, 0))
                        if 0 <= delay < float("inf"):
                            delays.append(delay)
                    except (ValueError, TypeError):
                        pass
                _next_request[key] = time.monotonic() + max(delays)
                continue
            return response
        raise ReportTaskError("WB ограничил запросы отчёта. Повторите позже.")
    finally:
        _lock.release()


def run_report_task(*, token, cancel, progress, session, task_key,
                    create_url, create_params, status_url, download_url,
                    poll_interval=4.0, poll_timeout=180.0, status_min_interval=5.0):
    """Runs one create->poll->download cycle. Returns the parsed download body (a list),
    or [] when WB answers 204 (report has no rows for the requested period)."""
    if not token:
        raise ReportTaskError("Добавьте API-токен WB с доступом к категории «Аналитика» в настройках.")
    headers = {"Authorization": token}

    def request_json(url, key, params=None, min_interval=60.0):
        response = _get(session, cancel, progress, url, key, params=params, headers=headers, min_interval=min_interval)
        message = {401: "Токен WB недействителен или истёк.",
                   403: "У токена WB нет доступа к категории «Аналитика»."}.get(response.status_code)
        if message:
            raise ReportTaskError(message)
        if response.status_code != 200:
            raise ReportTaskError(f"WB вернул HTTP {response.status_code}. Отчёт не получен.")
        try:
            return response.json()
        except ValueError:
            raise ReportTaskError("WB вернул некорректный JSON.") from None

    progress("Создание задачи на отчёт WB…")
    created = request_json(create_url, f"create:{task_key}", params=create_params)
    task_id = (created.get("data") or {}).get("taskId") if isinstance(created, dict) else None
    if not isinstance(task_id, str) or not task_id:
        raise ReportTaskError("WB не вернул идентификатор задачи отчёта.")

    deadline = time.monotonic() + poll_timeout
    while True:
        if cancel.is_set():
            raise ReportTaskError("Загрузка отчёта WB отменена.")
        progress("Проверка готовности отчёта WB…")
        body = request_json(status_url.format(task_id=task_id), f"status:{task_key}", min_interval=status_min_interval)
        task_status = (body.get("data") or {}).get("status") if isinstance(body, dict) else None
        if task_status == "done":
            break
        if time.monotonic() > deadline:
            raise ReportTaskError("WB слишком долго готовит отчёт. Повторите позже.")
        cancel.wait(poll_interval)

    progress("Скачивание отчёта WB…")
    response = _get(session, cancel, progress, download_url.format(task_id=task_id),
                    f"download:{task_key}", headers=headers, min_interval=60.0)
    if response.status_code == 204:
        return []
    message = {401: "Токен WB недействителен или истёк.",
               403: "У токена WB нет доступа к категории «Аналитика»."}.get(response.status_code)
    if message:
        raise ReportTaskError(message)
    if response.status_code != 200:
        raise ReportTaskError(f"WB вернул HTTP {response.status_code}. Отчёт не получен.")
    try:
        rows = response.json()
    except ValueError:
        raise ReportTaskError("WB вернул некорректный JSON.") from None
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise ReportTaskError("Некорректный формат отчёта WB.")
    return rows
