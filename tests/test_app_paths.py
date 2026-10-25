"""Where code and user data live: source checkout vs packaged app, and how scripts are launched."""
import importlib
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import app_paths  # noqa: E402

REPO = Path(__file__).resolve().parents[1]


def reload(**env):
    with patch.dict(os.environ, env, clear=False):
        if "OZONREPORTX_HOME" not in env:
            os.environ.pop("OZONREPORTX_HOME", None)
        return importlib.reload(app_paths)


class AppPathsTests(unittest.TestCase):
    def tearDown(self):
        os.environ.pop("OZONREPORTX_HOME", None)
        importlib.reload(app_paths)

    def test_source_checkout_keeps_data_next_to_code(self):
        paths = reload()
        self.assertFalse(paths.FROZEN)
        self.assertEqual((paths.CODE_ROOT, paths.DATA_ROOT), (REPO, REPO))
        self.assertEqual(paths.ENV_FILE, REPO / ".env")

    def test_packaged_app_keeps_data_in_documents_and_runs_scripts_with_itself(self):
        # Временная домашняя папка: на машине разработчика может лежать настоящий выбор папки с данными.
        with tempfile.TemporaryDirectory() as home, patch.object(Path, "home", return_value=Path(home)), \
             patch.object(sys, "frozen", True, create=True), \
             patch.object(sys, "_MEIPASS", "/Applications/OzonReportX.app/Contents/Frameworks", create=True), \
             patch.object(sys, "executable", "/Applications/OzonReportX.app/Contents/MacOS/OzonReportX"):
            paths = reload()
            self.assertEqual((paths.DATA_ROOT, paths.DATA_ROOT_SOURCE), (Path(home) / "Documents" / "OzonReportX", "default"))
            self.assertEqual(paths.script_command("Monthly_sales_report.py", "--month", "9"), [
                "/Applications/OzonReportX.app/Contents/MacOS/OzonReportX",
                "/Applications/OzonReportX.app/Contents/Frameworks/scripts/Monthly_sales_report.py", "--month", "9"])

    def test_folder_chosen_in_settings_is_used_by_packaged_app_until_it_disappears(self):
        with tempfile.TemporaryDirectory() as tmp:
            setting, data = Path(tmp) / "cfg" / "data_folder.txt", Path(tmp) / "repo"
            data.mkdir()
            with patch.object(app_paths, "DATA_FOLDER_SETTING", setting), patch.object(app_paths, "FROZEN", True), \
                 patch.dict(os.environ, {}, clear=False):
                os.environ.pop("OZONREPORTX_HOME", None)
                self.assertEqual(app_paths._data_root(), (app_paths.DEFAULT_DATA_ROOT, "default"))
                app_paths.save_data_folder(data)
                self.assertEqual(app_paths._data_root(), (data.resolve(), "chosen"))
                data.rmdir()
                self.assertEqual(app_paths._data_root()[1], "default")
                app_paths.save_data_folder(None)
                self.assertFalse(setting.exists())

    def test_home_override_wins(self):
        self.assertEqual(reload(OZONREPORTX_HOME="/tmp/ozon-data").DATA_ROOT, Path("/tmp/ozon-data"))


if __name__ == "__main__":
    unittest.main()
