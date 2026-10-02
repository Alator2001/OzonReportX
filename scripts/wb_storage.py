"""WB Paid Storage report — per-day, per-item warehouse storage cost.

One request covers at most 8 days (WB's limit), so a calendar month is split into
consecutive ≤8-day chunks and each is fetched as its own create/poll/download cycle.
Like Acceptance Expenses, past periods are allowed, so this can be tied to the
selected calendar month.
"""
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal
import threading

import requests

from scripts.wb_report_tasks import ReportTaskError, run_report_task

CREATE_URL = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage"
STATUS_URL = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/{task_id}/status"
DOWNLOAD_URL = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/{task_id}/download"
CHUNK_DAYS = 8


class StorageError(ReportTaskError):
    pass


def _chunks(date_from, date_to):
    cursor = date_from
    while cursor <= date_to:
        end = min(cursor + timedelta(days=CHUNK_DAYS - 1), date_to)
        yield cursor, end
        cursor = end + timedelta(days=1)


class StorageClient:
    def __init__(self, token, cancel=None, progress=lambda message: None, session=None):
        self.token, self.cancel, self.progress = token, cancel or threading.Event(), progress
        self.session = session or requests.Session()

    def fetch(self, date_from, date_to):
        chunks = list(_chunks(date_from, date_to))
        rows = []
        for index, (chunk_from, chunk_to) in enumerate(chunks, start=1):
            self.progress(f"Хранение WB: период {index}/{len(chunks)} ({chunk_from.isoformat()}–{chunk_to.isoformat()})")
            try:
                chunk_rows = run_report_task(
                    token=self.token, cancel=self.cancel, progress=self.progress, session=self.session,
                    task_key="paid_storage",
                    create_url=CREATE_URL,
                    create_params={"dateFrom": chunk_from.isoformat(), "dateTo": chunk_to.isoformat()},
                    status_url=STATUS_URL, download_url=DOWNLOAD_URL,
                )
            except ReportTaskError as exc:
                raise StorageError(str(exc)) from exc
            rows.extend(chunk_rows)
        return {"version": 1, "fetched_at": datetime.now().isoformat(timespec="seconds"),
                "date_from": date_from.isoformat(), "date_to": date_to.isoformat(), "rows": rows}


def summarize(data):
    """warehousePrice is already the line's total storage cost for that day — not a
    per-barcode price, so it is summed as-is; barcodesCount is informational only."""
    rows = data.get("rows", [])
    by_nm = defaultdict(lambda: {"total": Decimal(0), "subject": "", "vendor_code": ""})
    total = Decimal(0)
    for row in rows:
        nm_id = str(row.get("nmId") or "")
        amount = Decimal(str(row.get("warehousePrice", 0)))
        total += amount
        if nm_id:
            entry = by_nm[nm_id]
            entry["total"] += amount
            entry["subject"] = entry["subject"] or str(row.get("subject") or "")
            entry["vendor_code"] = entry["vendor_code"] or str(row.get("vendorCode") or "")
    return {"total": total, "rows": len(rows), "by_nm": dict(by_nm)}
