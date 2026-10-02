import os
import ast
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from scripts import marketplace_settings as settings


class MarketplaceSettingsTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.environment = patch.dict(os.environ)
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def test_migrates_existing_env_without_losing_secrets_or_comments(self):
        path = self.root / '.env'
        path.write_text("# custom config\nOLLAMA_MODEL='custom'\nOZON_CLIENT_ID='123'\nOZON_API_KEY='old-key'\n", encoding='utf-8')
        self.assertTrue(settings.needs_setup(self.root))
        settings.save_settings(self.root, {'WB_API_TOKEN': 'wb-token'})
        values = settings.read_settings(self.root)
        self.assertEqual(values['OZON_API_KEY'], 'old-key')
        self.assertEqual(values['OLLAMA_MODEL'], 'custom')
        self.assertEqual(values['WB_API_TOKEN'], 'wb-token')
        self.assertIn('# custom config', path.read_text(encoding='utf-8'))
        self.assertFalse(settings.needs_setup(self.root))

    def test_incomplete_pair_and_multiline_leave_file_unchanged(self):
        settings.save_settings(self.root, {})
        path = self.root / '.env'
        before = path.read_bytes()
        for updates in ({'OZON_CLIENT_ID': '123'}, {'OZON_PERF_API_KEY': 'secret'},
                        {'WB_API_TOKEN': 'token\nINJECTED=1'}):
            with self.assertRaises(ValueError):
                settings.save_settings(self.root, updates)
            self.assertEqual(path.read_bytes(), before)

    def test_write_failure_preserves_original_and_environment(self):
        settings.save_settings(self.root, {'WB_API_TOKEN': 'original'})
        before = (self.root / '.env').read_bytes()
        with patch.object(settings, 'set_key', side_effect=OSError('disk error')):
            with self.assertRaises(OSError):
                settings.save_settings(self.root, {'WB_API_TOKEN': 'replacement'})
        self.assertEqual((self.root / '.env').read_bytes(), before)
        self.assertEqual(os.environ['WB_API_TOKEN'], 'original')

    def test_single_marketplace_and_quoted_secret_roundtrip(self):
        token = "token'with\\characters#suffix"
        settings.save_settings(self.root, {'WB_API_TOKEN': token})
        self.assertEqual(settings.read_settings(self.root)['WB_API_TOKEN'], token)
        settings.save_settings(self.root, {'WB_API_TOKEN': ''})
        self.assertEqual(os.environ['WB_API_TOKEN'], '')

    def test_console_prompts_on_first_run_only_and_hides_secrets(self):
        with patch('builtins.print'), patch('builtins.input', return_value=''), patch.object(settings, 'getpass', return_value='') as secret:
            settings.configure_console(self.root)
            self.assertEqual(secret.call_count, 3)
            settings.configure_console(self.root)
            self.assertEqual(secret.call_count, 3)
            settings.configure_console(self.root, force=True)
            self.assertEqual(secret.call_count, 6)

    def test_desktop_dialog_builds_and_saves_with_installed_flet(self):
        import flet as ft
        import sys
        import threading
        source = Path(__file__).resolve().parents[1] / 'app_flet.py'
        tree = ast.parse(source.read_text(encoding='utf-8-sig'))
        function = next(node for node in ast.walk(tree)
                        if isinstance(node, ast.FunctionDef) and node.name == 'open_marketplace_settings')
        page = Mock(overlay=[])
        namespace = dict(ft=ft, ROOT=self.root, page=page, FIELDS=settings.FIELDS,
                         HELP=settings.HELP, read_settings=settings.read_settings,
                         save_settings=settings.save_settings, sys=sys,
                         job_lock=threading.Lock(), discount_dialog_state={'busy': False},
                         ai_state={'busy': False}, set_status=Mock(), PRIMARY='green')
        exec(compile(ast.Module(body=[function], type_ignores=[]), str(source), 'exec'), namespace)
        namespace['open_marketplace_settings']()
        dialog = page.overlay[0]
        fields = [control for control in dialog.content.controls if isinstance(control, ft.TextField)]
        self.assertEqual(len(fields), 5)
        self.assertTrue(fields[-1].password)
        fields[-1].value = 'test-wb-token'
        dialog.actions[-1].on_click(None)
        self.assertEqual(settings.read_settings(self.root)['WB_API_TOKEN'], 'test-wb-token')
        self.assertFalse(dialog.open)


if __name__ == '__main__':
    unittest.main()
