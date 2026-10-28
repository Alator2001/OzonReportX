"""Reports are split into reports/Ozon/<year>/ and reports/WB/<year>/; old flat reports move there by themselves."""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT))
import report_paths as rp  # noqa: E402


class ReportPathsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.reports = self.root / "reports"

    def put(self, *parts):
        path = self.reports.joinpath(*parts)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"fixture")
        return path

    def test_paths_by_store_and_year(self):
        self.assertEqual(rp.ozon_report_path(self.root, 2026, 10), self.reports / "Ozon" / "2026" / "Октябрь 2026.xlsx")
        self.assertEqual(rp.wb_report_path(self.root, 2025, 1), self.reports / "WB" / "2025" / "WB Январь 2025.xlsx")
        self.assertFalse((self.reports / "Ozon" / "2026").exists())
        rp.ozon_report_path(self.root, 2026, 10, create=True)
        self.assertTrue((self.reports / "Ozon" / "2026").is_dir())

    def test_old_flat_reports_move_by_store_and_year(self):
        self.put("Октябрь 2026.xlsx")
        self.put("Декабрь 2025.xlsx")
        self.put("WB Октябрь 2026.xlsx")
        moved = rp.migrate(self.root)
        self.assertEqual(len(moved), 3)
        self.assertEqual(rp.ozon_report_files(self.root), [self.reports / "Ozon" / "2025" / "Декабрь 2025.xlsx",
                                                           self.reports / "Ozon" / "2026" / "Октябрь 2026.xlsx"])
        self.assertEqual(rp.wb_report_files(self.root), [self.reports / "WB" / "2026" / "WB Октябрь 2026.xlsx"])
        self.assertEqual(list(self.reports.glob("*.xlsx")), [])

    def test_store_folder_without_year_is_sorted_too(self):
        self.put("Ozon", "Март 2026.xlsx")
        self.put("WB", "WB Март 2026.xlsx")
        rp.migrate(self.root)
        self.assertTrue((self.reports / "Ozon" / "2026" / "Март 2026.xlsx").is_file())
        self.assertTrue((self.reports / "WB" / "2026" / "WB Март 2026.xlsx").is_file())

    def test_temporary_files_and_files_without_year_stay(self):
        lock = self.put("~$Октябрь 2026.xlsx")
        temp = self.put("~tmp_report.xlsx")
        self.put("Заметки.xlsx")
        rp.migrate(self.root)
        self.assertTrue(lock.exists() and temp.exists())
        self.assertTrue((self.reports / "Ozon" / "Заметки.xlsx").is_file())
        self.assertEqual(rp.ozon_report_files(self.root), [self.reports / "Ozon" / "Заметки.xlsx"])

    def test_existing_report_is_never_overwritten(self):
        old = self.put("Октябрь 2026.xlsx")
        new = self.put("Ozon", "2026", "Октябрь 2026.xlsx")
        new.write_bytes(b"new")
        with mock.patch("builtins.print") as printed:
            rp.migrate(self.root)
        self.assertEqual(new.read_bytes(), b"new")
        self.assertTrue(old.exists())
        self.assertIn("не перенесён", printed.call_args[0][0])

    def test_no_reports_folder(self):
        self.assertEqual(rp.migrate(self.root), [])
        self.assertEqual(rp.ozon_report_files(self.root), [])

    def test_migration_is_idempotent(self):
        self.put("Октябрь 2026.xlsx")
        rp.migrate(self.root)
        self.assertEqual(rp.migrate(self.root), [])


if __name__ == "__main__":
    unittest.main()
