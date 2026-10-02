from datetime import date
from decimal import Decimal
import unittest
from unittest.mock import patch

from scripts import wb_acceptance as acceptance
import test_technical_fixes as technical


class WBAcceptanceTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def test_fetch_passes_date_range_and_wraps_errors(self):
        client = acceptance.AcceptanceClient("test")
        with patch("scripts.wb_acceptance.run_report_task", return_value=[
            {"nmID": 123, "count": 2, "total": 50.5, "subjectName": "Доски"},
            {"nmID": 123, "count": 1, "total": 25.25, "subjectName": "Доски"},
            {"nmID": 456, "count": 3, "total": 90, "subjectName": "Ножи"},
        ]) as run:
            data = client.fetch(date(2026, 8, 1), date(2026, 8, 31))
        self.assertEqual(run.call_args.kwargs["create_params"], {"dateFrom": "2026-08-01", "dateTo": "2026-08-31"})
        summary = acceptance.summarize(data)
        self.assertEqual(summary["total"], Decimal("165.75"))
        self.assertEqual(summary["by_nm"]["123"]["total"], Decimal("75.75"))
        self.assertEqual(summary["by_nm"]["123"]["count"], 3)
        self.assertEqual(summary["by_nm"]["456"]["subject"], "Ножи")

    def test_empty_report_summarizes_to_zero(self):
        client = acceptance.AcceptanceClient("test")
        with patch("scripts.wb_acceptance.run_report_task", return_value=[]):
            data = client.fetch(date(2026, 8, 1), date(2026, 8, 31))
        summary = acceptance.summarize(data)
        self.assertEqual(summary["total"], Decimal(0))
        self.assertEqual(summary["by_nm"], {})

    def test_task_errors_are_reraised_as_acceptance_error(self):
        from scripts.wb_report_tasks import ReportTaskError
        client = acceptance.AcceptanceClient("test")
        with patch("scripts.wb_acceptance.run_report_task", side_effect=ReportTaskError("boom")):
            with self.assertRaises(acceptance.AcceptanceError):
                client.fetch(date(2026, 8, 1), date(2026, 8, 31))


if __name__ == "__main__":
    unittest.main()
