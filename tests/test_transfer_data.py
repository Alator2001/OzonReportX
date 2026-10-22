import unicodedata
import unittest
import zipfile

from scripts import transfer_data as transfer
import test_technical_fixes as technical


NFD_NAME = unicodedata.normalize("NFD", "Май 2026.xlsx")
NFC_NAME = unicodedata.normalize("NFC", "Май 2026.xlsx")


class TransferDataTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def make_zip(self, entries):
        archive = self.directory / "OzonReportX_transfer_test.zip"
        with zipfile.ZipFile(archive, "w") as zf:
            for name, data in entries.items():
                zf.writestr(name, data)
        return archive

    def test_export_writes_nfc_names(self):
        root = self.directory / "project"
        (root / "reports").mkdir(parents=True)
        (root / "reports" / NFD_NAME).write_bytes(b"report")
        (root / ".env").write_text("OZON_API_KEY=x", encoding="utf-8")
        archive = transfer.export_data(root, self.directory / "out.zip")
        with zipfile.ZipFile(archive) as zf:
            names = zf.namelist()
        self.assertIn(f"reports/{NFC_NAME}", names)
        self.assertIn(".env", names)

    def test_import_skips_macos_junk_and_normalizes_names(self):
        root = self.directory / "project"
        root.mkdir()
        archive = self.make_zip({
            ".env": "OZON_API_KEY=x",
            f"reports/{NFD_NAME}": "report",
            "reports/.DS_Store": "junk",
            "reports/._Май 2026.xlsx": "junk",
            "__MACOSX/reports/._Май 2026.xlsx": "junk",
        })
        transfer.import_data(root, archive)
        self.assertEqual((root / ".env").read_text(encoding="utf-8"), "OZON_API_KEY=x")
        self.assertEqual([p.name for p in (root / "reports").iterdir()], [NFC_NAME])
        self.assertFalse((root / "__MACOSX").exists())

    def test_import_strips_folder_added_by_finder(self):
        root = self.directory / "project"
        root.mkdir()
        archive = self.make_zip({
            "OzonReportX_transfer_test/costs.xlsx": "costs",
            "OzonReportX_transfer_test/reports/Июнь 2026.xlsx": "report",
        })
        transfer.import_data(root, archive)
        self.assertEqual((root / "costs.xlsx").read_text(encoding="utf-8"), "costs")
        self.assertTrue((root / "reports" / "Июнь 2026.xlsx").is_file())
        self.assertFalse((root / "OzonReportX_transfer_test").exists())

    def test_import_keeps_known_top_folder(self):
        root = self.directory / "project"
        root.mkdir()
        archive = self.make_zip({"reports/Июнь 2026.xlsx": "report"})
        transfer.import_data(root, archive)
        self.assertTrue((root / "reports" / "Июнь 2026.xlsx").is_file())

    def test_import_from_folder_unpacked_by_safari(self):
        root = self.directory / "project"
        root.mkdir()
        (root / "costs.xlsx").write_text("old", encoding="utf-8")
        unpacked = self.directory / "OzonReportX_transfer_test"
        (unpacked / "reports").mkdir(parents=True)
        (unpacked / ".env").write_text("OZON_API_KEY=x", encoding="utf-8")
        (unpacked / "costs.xlsx").write_text("new", encoding="utf-8")
        (unpacked / "reports" / NFD_NAME).write_text("report", encoding="utf-8")
        (unpacked / ".DS_Store").write_text("junk", encoding="utf-8")
        transfer.import_data(root, unpacked)
        self.assertEqual((root / "costs.xlsx").read_text(encoding="utf-8"), "new")
        self.assertTrue((root / ".env").is_file())
        self.assertFalse((root / ".DS_Store").exists())
        self.assertEqual([p.name for p in (root / "reports").iterdir()], [NFC_NAME])
        backups = list(root.glob("backup_*/costs.xlsx"))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_text(encoding="utf-8"), "old")

    def test_import_rejects_paths_outside_project(self):
        root = self.directory / "project"
        root.mkdir()
        archive = self.make_zip({"../evil.txt": "x", "costs.xlsx": "costs"})
        transfer.import_data(root, archive)
        self.assertFalse((self.directory / "evil.txt").exists())
        self.assertTrue((root / "costs.xlsx").is_file())


if __name__ == "__main__":
    unittest.main()
