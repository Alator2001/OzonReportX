"""WB Acceptance Expenses report — per-income (поставка) acceptance cost by item.

Unlike Order Feed/Returns, this report accepts any past date range (verified against
dev.wildberries.ru/en/docs/openapi/reports#tag/mainReports/paths/~1api~1v1~1acceptance_report),
so it can be tied to the selected calendar month like the financial report.
"""
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal
import threading

import requests

from scripts.wb_report_tasks import ReportTaskError, run_report_task

CREATE_URL = "https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report"
STATUS_URL = "https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/{task_id}/status"
DOWNLOAD_URL = "https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/{task_id}/download"


class AcceptanceError(ReportTaskError):
    pass


class AcceptanceClient:
    def __init__(self, token, cancel=None, progress=lambda message: None, session=None):
        self.token, self.cancel, self.progress = token, cancel or threading.Event(), progress
        self.session = session or requests.Session()

    def fetch(self, date_from, date_to):
        """date_from/date_to: date objects, inclusive, span at most 31 days (WB's limit)."""
        try:
            rows = run_report_task(
                token=self.token, cancel=self.cancel, progress=self.progress, session=self.session,
                task_key="acceptance",
                create_url=CREATE_URL, create_params={"dateFrom": date_from.isoformat(), "dateTo": date_to.isoformat()},
                status_url=STATUS_URL, download_url=DOWNLOAD_URL,
            )
        except ReportTaskError as exc:
            raise AcceptanceError(str(exc)) from exc
        return {"version": 1, "fetched_at": datetime.now().isoformat(timespec="seconds"),
                "date_from": date_from.isoformat(), "date_to": date_to.isoformat(), "rows": rows}


def summarize(data):
    rows = data.get("rows", [])
    by_nm = defaultdict(lambda: {"count": 0, "total": Decimal(0), "subject": ""})
    total = Decimal(0)
    for row in rows:
        nm_id = str(row.get("nmID") or row.get("nmId") or "")
        amount = Decimal(str(row.get("total", 0)))
        total += amount
        if nm_id:
            entry = by_nm[nm_id]
            entry["count"] += int(row.get("count") or 0)
            entry["total"] += amount
            entry["subject"] = entry["subject"] or str(row.get("subjectName") or "")
    return {"total": total, "rows": len(rows), "by_nm": dict(by_nm)}
