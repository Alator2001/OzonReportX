"""Offline checks: start without Ozon keys, non-interactive update check, offer_id lookup on v3."""
import contextlib
import importlib
import io
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import _auto_update as auto_update  # noqa: E402
import balance_report  # noqa: E402
import recommended_prices as rp  # noqa: E402


class StartWithoutKeysTests(unittest.TestCase):
    def test_balance_module_imports_without_keys_and_fails_only_on_request(self):
        without_keys = {k: v for k, v in os.environ.items() if k not in ("OZON_CLIENT_ID", "OZON_API_KEY")}
        with patch.dict(os.environ, without_keys, clear=True), patch("dotenv.load_dotenv", return_value=False):
            module = importlib.reload(balance_report)  # раньше здесь был sys.exit(1)
            with patch.object(module.requests, "post") as post:
                with self.assertRaisesRegex(RuntimeError, "OZON_CLIENT_ID"):
                    module.get_balance_report("2026-10-01", "2026-10-02")
            post.assert_not_called()

    def test_balance_request_uses_keys_saved_after_import(self):
        with patch.dict(os.environ, {"OZON_CLIENT_ID": "new-id", "OZON_API_KEY": "new-key"}), \
                patch.object(balance_report.requests, "post", return_value=Mock(status_code=200, json=lambda: {})) as post:
            balance_report.get_balance_report("2026-10-01", "2026-10-02")
        self.assertEqual(post.call_args.kwargs["headers"]["Client-Id"], "new-id")


class UpdateCheckTests(unittest.TestCase):
    def setUp(self):
        release = {"tag": "9.9.9", "download_url": "x", "name": "Тест", "notes": ""}
        for name, value in (("get_current_version", lambda: "2.1.2"), ("get_latest_release", lambda: release)):
            patcher = patch.object(auto_update, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        output = contextlib.redirect_stdout(io.StringIO())
        output.__enter__()
        self.addCleanup(output.__exit__, None, None, None)

    def test_without_terminal_only_reports_and_never_waits_for_input(self):
        with patch.object(auto_update.sys, "stdin", io.StringIO("")), \
                patch("builtins.input", side_effect=AssertionError("ждёт ввода")), \
                patch.object(auto_update, "create_backup") as backup:
            self.assertFalse(auto_update.check_and_update())
        backup.assert_not_called()

    def test_install_button_installs_without_question_and_without_restart(self):
        with patch("builtins.input", side_effect=AssertionError("ждёт ввода")), \
                patch.object(auto_update, "create_backup", return_value=Path("backup")), \
                patch.object(auto_update, "download_and_extract", return_value=Path("src")), \
                patch.object(auto_update, "apply_update") as apply, \
                patch.object(auto_update, "update_version_file") as version, \
                patch.object(auto_update.shutil, "rmtree"), \
                patch.object(auto_update.Path, "mkdir"), \
                patch.object(auto_update.os, "execv") as execv:
            self.assertTrue(auto_update.check_and_update(assume_yes=True, restart=False))
        apply.assert_called_once()
        version.assert_called_once_with("9.9.9")
        execv.assert_not_called()


class OfferIdLookupTests(unittest.TestCase):
    def test_offer_id_comes_from_v3_product_info_list(self):
        answer = Mock(status_code=200, json=lambda: {"items": [{"id": 1376477009, "offer_id": "1101"}]})
        answer.raise_for_status = Mock()
        with patch.multiple(rp, OZON_CLIENT_ID="id", OZON_API_KEY="key"), \
                patch.object(rp.requests, "post", return_value=answer) as post:
            self.assertEqual(rp.get_product_info_by_id(1376477009), "1101")
        self.assertEqual(post.call_args.args[0], "https://api-seller.ozon.ru/v3/product/info/list")
        self.assertEqual(post.call_args.kwargs["json"], {"product_id": ["1376477009"]})

    def test_prices_without_keys_return_a_pair(self):
        # Вызывающий код распаковывает «prices_map, marketing_prices_map = …»: одиночный {} ронял расчёт.
        with patch.multiple(rp, OZON_CLIENT_ID=None, OZON_API_KEY=None), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(rp.get_product_prices_from_ozon(["1101"]), ({}, {}))
        with patch.multiple(rp, OZON_CLIENT_ID="id", OZON_API_KEY="key"):
            self.assertEqual(rp.get_product_prices_from_ozon([]), ({}, {}))

    def test_unknown_product_gives_none(self):
        answer = Mock(status_code=200, json=lambda: {"items": []})
        answer.raise_for_status = Mock()
        with patch.multiple(rp, OZON_CLIENT_ID="id", OZON_API_KEY="key"), \
                patch.object(rp.requests, "post", return_value=answer):
            self.assertIsNone(rp.get_product_info_by_id(1))


if __name__ == "__main__":
    unittest.main()
