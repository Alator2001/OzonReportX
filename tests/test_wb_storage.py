from datetime import date
from decimal import Decimal
import unittest
from unittest.mock import patch

from scripts import wb_storage as storage
import test_technical_fixes as technical


def row(nm_id=123, price=1.5, subject="Доски", vendor="1114"):
    return {"nmId": nm_id, "warehousePrice": price, "barcodesCount": 2,
            "subject": subject, "vendorCode": vendor, "date": "2026-08-01"}


class WBStorageTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def test_month_is_split_into_8_day_chunks(self):
        with patch("scripts.wb_storage.run_report_task", return_value=[]) as run:
            storage.StorageClient("test").fetch(date(2026, 8, 1), date(2026, 8, 31))
        ranges = [(call.kwargs["create_params"]["dateFrom"], call.kwargs["create_params"]["dateTo"])
                  for call in run.call_args_list]
        self.assertEqual(ranges, [
            ("2026-08-01", "2026-08-08"), ("2026-08-09", "2026-08-16"),
            ("2026-08-17", "2026-08-24"), ("2026-08-25", "2026-08-31"),
        ])

    def test_rows_from_all_chunks_are_merged(self):
        with patch("scripts.wb_storage.run_report_task", side_effect=[[row(nm_id=1)], [row(nm_id=2)]]):
            data = storage.StorageClient("test").fetch(date(2026, 8, 1), date(2026, 8, 9))
        self.assertEqual(len(data["rows"]), 2)

    def test_summarize_sums_warehouse_price_without_multiplying_by_barcode_count(self):
        data = {"rows": [row(nm_id=123, price=1.5), row(nm_id=123, price=2.5), row(nm_id=456, price=1.0)]}
        summary = storage.summarize(data)
        self.assertEqual(summary["total"], Decimal("5.0"))
        self.assertEqual(summary["by_nm"]["123"]["total"], Decimal("4.0"))
        self.assertEqual(summary["by_nm"]["123"]["vendor_code"], "1114")

    def test_task_errors_are_reraised_as_storage_error(self):
        from scripts.wb_report_tasks import ReportTaskError
        with patch("scripts.wb_storage.run_report_task", side_effect=ReportTaskError("boom")):
            with self.assertRaises(storage.StorageError):
                storage.StorageClient("test").fetch(date(2026, 8, 1), date(2026, 8, 8))


if __name__ == "__main__":
    unittest.main()
