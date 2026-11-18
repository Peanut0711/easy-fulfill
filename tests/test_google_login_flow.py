"""Google 로그인(브라우저 인증)이 백그라운드 작업에서 열리거나 화면을 멈추지 않는지 검증한다."""

import json
import os
import subprocess
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import sheets.google_sheets_oauth as oauth  # noqa: E402


class LoginPolicyTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        auth_dir = Path(self.directory.name)
        (auth_dir / "credentials.json").write_text("{}", encoding="utf-8")
        for name, value in (("GOOGLE_AUTH_DIR", auth_dir), ("OAUTH_CREDENTIAL_PATH", auth_dir / "credentials.json"),
                            ("TOKEN_PATH", auth_dir / "token.json")):
            patcher = patch.object(oauth, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.flow = MagicMock()
        from google_auth_oauthlib.flow import InstalledAppFlow  # 모듈은 로그인할 때 이 클래스를 불러온다
        patcher = patch.object(InstalledAppFlow, "from_client_secrets_file", return_value=self.flow)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self.directory.cleanup)

    def test_background_call_never_opens_browser_login(self):
        with self.assertRaises(oauth.GoogleLoginRequiredError):
            oauth.get_authorized_gspread_client(interactive=False)
        self.flow.run_local_server.assert_not_called()

    def test_gui_default_is_non_interactive(self):
        previous = oauth.set_interactive_login_default(False)
        self.addCleanup(oauth.set_interactive_login_default, previous)
        with self.assertRaises(oauth.GoogleLoginRequiredError):
            oauth.get_authorized_gspread_client()
        self.flow.run_local_server.assert_not_called()

    def test_only_one_interactive_login_at_a_time(self):
        self.assertTrue(oauth._login_lock.acquire(blocking=False))
        try:
            with self.assertRaises(oauth.GoogleLoginRequiredError) as raised:
                oauth.get_authorized_gspread_client(interactive=True)
        finally:
            oauth._login_lock.release()
        self.assertIn("진행 중", str(raised.exception))
        self.flow.run_local_server.assert_not_called()

    def test_interactive_login_has_timeout(self):
        self.flow.run_local_server.side_effect = AttributeError("'NoneType' object has no attribute 'replace'")
        with self.assertRaises(oauth.GoogleLoginRequiredError) as raised:
            oauth.get_authorized_gspread_client(interactive=True, timeout_seconds=1)
        self.assertIn("시간", str(raised.exception))
        self.assertEqual(self.flow.run_local_server.call_args.kwargs.get("timeout_seconds"), 1)

    def test_login_required_message_is_recognised_as_oauth_error(self):
        from tests.test_order_load_sheets import load_app_module
        module = load_app_module()
        error = RuntimeError(str(oauth.GoogleLoginRequiredError()))  # 화면에서는 문자열로 다시 감싼다
        self.assertTrue(module._is_likely_google_sheets_oauth_error(error))


def run_scenario(name):
    """재인증 버튼 시나리오(자식 프로세스, os._exit로 종료)."""
    import importlib.util
    import time

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication, QMessageBox

    spec = importlib.util.spec_from_file_location("easy_fulfill_login_scenario", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    calls = []

    def slow_login(*args, **kwargs):
        calls.append(kwargs.get("interactive"))
        time.sleep(1.0)  # 브라우저 로그인 대기를 흉내
        return "CLIENT"

    dialogs = []
    with patch.object(oauth, "get_authorized_gspread_client", slow_login), \
            patch.object(oauth, "delete_oauth_token_file", lambda: True), \
            patch.object(QMessageBox, "information", side_effect=lambda *a, **k: dialogs.append(a[1])), \
            patch.object(QMessageBox, "critical", side_effect=lambda *a, **k: dialogs.append(a[1])):
        window = app_module.MainWindow()
        window._startup_order_index_sync_started = True
        started = time.monotonic()
        handled = window._on_google_reauth_clicked(ask_confirm=False)
        returned_after = time.monotonic() - started
        deadline = time.monotonic() + 5
        while window._gspread_client is None and time.monotonic() < deadline:
            app.processEvents()
            time.sleep(0.02)
    result = {"handled": handled, "returned_after": returned_after, "client": window._gspread_client,
              "interactive": calls, "dialogs": dialogs, "notice": window.notice_bar.text()}
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


class ReauthButtonTests(unittest.TestCase):
    def test_reauth_runs_in_background_without_freezing_window(self):
        with tempfile.TemporaryDirectory() as directory:
            env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
            completed = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), "--scenario", "reauth"],
                cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=120,
            )
        line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
        self.assertIsNotNone(line, completed.stdout[-2000:] + completed.stderr[-2000:])
        result = json.loads(line[len("RESULT "):])
        self.assertTrue(result["handled"])
        self.assertLess(result["returned_after"], 0.5)  # 수정 전에는 로그인이 끝날 때까지(1초) 멈춤
        self.assertEqual(result["client"], "CLIENT")
        self.assertEqual(result["interactive"], [True])
        # 완료는 확인 창 대신 알림 띠로 알린다(알림·상태바 계획 C-3).
        self.assertEqual(result["dialogs"], [])
        self.assertIn("연동이 완료", result["notice"])


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--scenario":
        run_scenario(sys.argv[2])
    else:
        unittest.main()
