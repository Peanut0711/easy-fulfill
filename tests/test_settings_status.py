"""「설정」 탭 상태 표시(창 크기·설정 탭 계획 B-2)를 검증한다.

- Google 연동 상태 종류(토큰 파일 기준)와 실행 중 인증 오류 표시
- 분류 목록 배지(연동 문제가 있으면 「연동  ✗」)
- 슬랙 Webhook·문의 알림·종추적 키 상태 종류
- 슬랙 Webhook 버튼이 「알림」 페이지로 옮겨졌는지
"""

import os
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class OAuthStatusTests(unittest.TestCase):
    def setUp(self):
        import sheets.google_sheets_oauth as oauth
        self.oauth = oauth
        self.dir = Path(tempfile.mkdtemp(prefix="ef-oauth-"))
        self.credential = self.dir / "credentials.json"
        self.token = self.dir / "token.json"
        patcher = patch.multiple(oauth, OAUTH_CREDENTIAL_PATH=self.credential, TOKEN_PATH=self.token)
        patcher.start()
        self.addCleanup(patcher.stop)

    def creds(self, **kwargs):
        values = {"valid": False, "expired": False, "refresh_token": None, **kwargs}
        return patch.object(self.oauth.Credentials, "from_authorized_user_file",
                            return_value=SimpleNamespace(**values))

    def test_kinds(self):
        status = self.oauth.get_oauth_status
        self.assertEqual(status()[0], "error")  # credentials.json 없음
        self.credential.write_text("{}", encoding="utf-8")
        self.assertEqual(status()[0], "warn")  # 이 PC 로그인 없음
        self.token.write_text("{}", encoding="utf-8")
        with patch.object(self.oauth.Credentials, "from_authorized_user_file", side_effect=ValueError("bad")):
            self.assertEqual(status()[0], "error")
        with self.creds(valid=True):
            self.assertEqual(status()[:2], ("ok", "연결됨"))
        # 액세스 토큰은 한 시간마다 만료되고 쓸 때 갱신되므로 정상으로 본다.
        with self.creds(expired=True, refresh_token="r"):
            self.assertEqual(status()[0], "ok")
        with self.creds():
            self.assertEqual(status()[0], "error")

    def test_description_keeps_text_and_detail(self):
        self.credential.write_text("{}", encoding="utf-8")
        text = self.oauth.get_oauth_status_description()
        self.assertIn("로그인이 없습니다", text)
        self.assertIn("재인증", text.splitlines()[1])


class SettingsStatusWindowTests(unittest.TestCase):
    METHODS = (
        "_refresh_google_auth_status_ui", "_mark_google_auth_runtime_error", "_clear_google_auth_runtime_error",
        "_settings_integration_problem", "_update_settings_nav_badges", "_key_status_kind", "_key_status_short",
        "_update_status_displays", "_set_naver_inquiry_status", "_update_naver_inquiry_status_label",
        "_on_naver_inquiry_result", "_apply_order_index_sync_payload",
    )

    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from tests.test_order_load_sheets import load_app_module
        cls.app = load_app_module()

    def window(self, auth=("ok", "연결됨", "")):
        from PySide6.QtWidgets import QLabel, QListWidget, QStackedWidget, QWidget
        cls = type("FakeWindow", (), {name: getattr(self.app.MainWindow, name) for name in self.METHODS})
        window = cls()
        window.ui = SimpleNamespace()
        nav = QListWidget()
        stack = QStackedWidget()
        for key, title in (("integration", "연동"), ("alerts", "알림"), ("program", "프로그램")):
            nav.addItem(title)
            page = QWidget()
            page.setObjectName(f"scrollArea_settings_{key}")
            stack.addWidget(page)
        window.ui.listWidget_settings_nav = nav
        window.ui.stackedWidget_settings = stack
        window.ui.label_google_auth_status = QLabel()
        window.ui.label_slack_webhook_status = QLabel()
        window.ui.label_settings_api_status = QLabel()
        window._dlg_key_status = None
        window._dlg_slack_status = None
        window._dlg_naver_status = QLabel()
        window._key_status_text = "확인 중…"
        window._slack_status_text = "미설정"
        window._naver_inquiry_status_text = "미설정"
        window._naver_inquiry_status_kind = "idle"
        window._google_auth_kind = ""
        window._google_auth_runtime_error = ""
        window._inquiry_work_hours = lambda: (10, 19)
        window.index_label_updates = 0
        window._update_index_sheet_sync_label = lambda: setattr(
            window, "index_label_updates", window.index_label_updates + 1)
        self._auth = auth
        patcher = patch("sheets.google_sheets_oauth.get_oauth_status", side_effect=lambda: self._auth)
        patcher.start()
        self.addCleanup(patcher.stop)
        return window

    def nav_texts(self, window):
        nav = window.ui.listWidget_settings_nav
        return [nav.item(row).text() for row in range(nav.count())]

    def test_google_status_and_nav_badge(self):
        window = self.window(auth=("ok", "연결됨", "저장된 토큰이 유효합니다."))
        window._refresh_google_auth_status_ui()
        label = window.ui.label_google_auth_status
        self.assertEqual((label.text(), label.property("status")), ("✓ 연결됨", "ok"))
        self.assertEqual(self.nav_texts(window), ["연동", "알림", "프로그램"])

        self._auth = ("error", "token.json을 읽을 수 없습니다.", "「연결 해제」 후 「재인증」")
        window._refresh_google_auth_status_ui()
        self.assertEqual(label.property("status"), "error")
        self.assertEqual(self.nav_texts(window), ["연동  ✗", "알림", "프로그램"])

        self._auth = ("warn", "이 PC에 저장된 Google 로그인이 없습니다.", "")
        window._refresh_google_auth_status_ui()
        self.assertEqual(self.nav_texts(window)[0], "연동")  # 로그인 전은 배지 없이 주의 표시만

    def test_runtime_auth_error_until_cleared(self):
        window = self.window()
        window._mark_google_auth_runtime_error("invalid_grant: Token has been expired or revoked.")
        label = window.ui.label_google_auth_status
        self.assertEqual(label.property("status"), "error")
        self.assertIn("invalid_grant", label.toolTip())
        self.assertEqual(self.nav_texts(window)[0], "연동  ✗")
        window._clear_google_auth_runtime_error()
        window._refresh_google_auth_status_ui()
        self.assertEqual(label.property("status"), "ok")
        self.assertEqual(self.nav_texts(window)[0], "연동")

    def test_index_poll_marks_and_clears_auth_error(self):
        window = self.window()
        window._prompt_google_reauth_for_oauth_error = lambda **kw: self.fail("폴링 중에는 창을 띄우지 않는다")
        window._apply_order_index_sync_payload(
            {"ok": False, "error": "invalid_grant: Token has been expired or revoked."}, prompt_reauth=False)
        self.assertEqual(self.nav_texts(window)[0], "연동  ✗")
        window._apply_order_index_sync_payload({"ok": False, "error": "네트워크 오류"}, prompt_reauth=False)
        self.assertEqual(self.nav_texts(window)[0], "연동  ✗")  # 인증과 무관한 실패는 상태를 바꾸지 않는다
        window._apply_order_index_sync_payload({"ok": True, "kind": "unchanged"}, prompt_reauth=False)
        self.assertEqual(self.nav_texts(window)[0], "연동")
        self.assertEqual(window.ui.label_google_auth_status.property("status"), "ok")

    def test_key_check_failure_badges_integration(self):
        window = self.window()
        cases = {
            "유효함 ✓": "ok", "미설정 — 「키 변경」으로 등록하세요.": "warn", "설정됨 (오프라인 — 유효성 미확인)": "warn",
            "확인 중…": "idle", "검증 중…": "idle", "확인 실패 — 「키 변경」으로 다시 등록": "error",
            "검증 실패 — 저장하지 않음": "error",
        }
        for text, kind in cases.items():
            window._key_status_text = text
            self.assertEqual(window._key_status_kind(), kind, text)
        window._key_status_text = "확인 실패 — 「키 변경」으로 다시 등록"
        window._update_status_displays()
        self.assertEqual(self.nav_texts(window)[0], "연동  ✗")
        window._key_status_text = "유효함 ✓"
        window._update_status_displays()
        self.assertEqual(self.nav_texts(window)[0], "연동")

    def test_slack_webhook_status(self):
        window = self.window()
        window._update_status_displays()
        label = window.ui.label_slack_webhook_status
        self.assertEqual(label.property("status"), "warn")
        self.assertTrue(label.text().startswith("! 미설정"))
        window._slack_status_text = "설정됨 ✓"
        window._update_status_displays()
        self.assertEqual((label.text(), label.property("status")), ("✓ 설정됨", "ok"))
        self.assertNotIn("슬랙", window.ui.label_settings_api_status.text())

    def test_inquiry_status_kinds(self):
        window = self.window()
        label = window._dlg_naver_status
        window._on_naver_inquiry_result({"ok": True, "open": 3, "new": 0, "reminded": 0, "sent": 0})
        self.assertEqual(label.property("status"), "warn")
        self.assertTrue(label.text().startswith("! 미답변 3건"))
        window._on_naver_inquiry_result({"ok": True, "open": 0})
        self.assertEqual(label.property("status"), "ok")
        window._on_naver_inquiry_result({"ok": False, "error": "x" * 100})
        self.assertEqual(label.property("status"), "error")
        self.assertEqual(label.toolTip(), "x" * 100)
        window._set_naver_inquiry_status("조회 중…", "idle")
        self.assertEqual((label.text(), label.property("status")), ("조회 중…", "idle"))


class SettingsUiFileTests(unittest.TestCase):
    def test_slack_webhook_button_moved_to_alerts_page(self):
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        alerts = tree.find(".//widget[@name='scrollArea_settings_alerts']")
        self.assertIsNotNone(alerts.find(".//widget[@name='groupBox_slack_webhook']//widget[@name='pushButton_cfg_slack_webhook']"))
        self.assertIsNotNone(alerts.find(".//widget[@name='label_slack_webhook_status']"))
        keys = tree.find(".//widget[@name='groupBox_api_keys']")
        self.assertIsNone(keys.find(".//widget[@name='pushButton_cfg_slack_webhook']"))

    def test_status_labels_have_no_widget_color(self):
        # 위젯 스타일시트가 앱 스타일시트(set_status 색)를 덮지 않게 한다.
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        for name in ("label_google_auth_status", "label_index_sheet_sync", "label_slack_webhook_status"):
            widget = tree.find(f".//widget[@name='{name}']")
            self.assertIsNone(widget.find("./property[@name='styleSheet']"), name)


if __name__ == "__main__":
    unittest.main()
