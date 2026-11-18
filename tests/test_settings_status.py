"""「설정」 탭 상태 표시(창 크기·설정 탭 계획 B-2·B-3)를 검증한다.

- Google 연동 상태 종류(토큰 파일 기준)와 실행 중 인증 오류 표시
- 분류 목록 배지(연동 문제가 있으면 「연동  ✗」)
- 슬랙 Webhook·문의 알림·종추적 키 상태 종류
- 슬랙 Webhook 버튼이 「알림」 페이지로 옮겨졌는지
- API 키 키별 행(등록 여부, 검증 결과, 등록/변경 버튼), 인증 안내는 오류일 때만, 연결 해제 위험 버튼
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
        "API_KEY_ROWS", "_api_key_row_status", "_update_api_key_rows", "_set_shared_key_status",
        "_apply_shared_key_registration", "_on_naver_creds_validated", "_on_coupang_creds_validated",
        "_shared_key_status_store",
        "_on_settings_page_shown",
    )

    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from tests.test_order_load_sheets import load_app_module
        cls.app = load_app_module()

    def window(self, auth=("ok", "연결됨", "")):
        from PySide6.QtWidgets import QLabel, QListWidget, QPushButton, QStackedWidget, QWidget
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
        window.ui.label_google_auth_tip = QLabel()
        for label_name, button_name in self.app.MainWindow.API_KEY_ROWS.values():
            setattr(window.ui, label_name, QLabel())
            setattr(window.ui, button_name, QPushButton("변경"))
        window._shared_key_status = {}
        window.settings = {}
        window.get_app_setting = lambda key, default=None: window.settings.get(key, default)
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
            "유효함 ✓": "ok", "미설정 — 「등록」을 눌러 주세요.": "warn", "설정됨 (오프라인 — 유효성 미확인)": "warn",
            "확인 중…": "idle", "검증 중…": "idle", "확인 실패 — 「변경」으로 다시 등록": "error",
            "검증 실패 — 저장하지 않음": "error",
        }
        for text, kind in cases.items():
            window._key_status_text = text
            self.assertEqual(window._key_status_kind(), kind, text)
        window._key_status_text = "확인 실패 — 「변경」으로 다시 등록"
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

    def row(self, window, key):
        label_name, button_name = self.app.MainWindow.API_KEY_ROWS[key]
        label, button = getattr(window.ui, label_name), getattr(window.ui, button_name)
        return label.text(), label.property("status"), button.text()

    def test_google_tip_only_on_error(self):
        window = self.window()
        window._refresh_google_auth_status_ui()
        self.assertTrue(window.ui.label_google_auth_tip.isHidden())
        self._auth = ("error", "token.json을 읽을 수 없습니다.", "")
        window._refresh_google_auth_status_ui()
        self.assertFalse(window.ui.label_google_auth_tip.isHidden())

    def test_api_key_rows_from_shared_registration(self):
        window = self.window()
        window._update_status_displays()
        self.assertEqual(self.row(window, "naver"), ("확인 중…", "idle", "변경"))
        window._apply_shared_key_registration({"postcode": False, "naver": True, "coupang": False})
        self.assertEqual(self.row(window, "naver"), ("✓ 등록됨", "ok", "변경"))
        self.assertEqual(self.row(window, "coupang"), ("! 미설정", "warn", "등록"))
        self.assertEqual(self.row(window, "postcode"), ("! 미설정", "warn", "등록"))
        # 주소 검색 전용 키가 없으면 종추적 키로 조회한다는 것을 보인다.
        window.settings["kpost_regkey"] = "abc"
        window._apply_shared_key_registration({"postcode": False, "naver": True, "coupang": True})
        text, kind, button = self.row(window, "postcode")
        self.assertEqual((kind, button), ("warn", "등록"))
        self.assertIn("종추적 키로 조회", text)

    def test_api_key_rows_when_sheet_unreadable(self):
        window = self.window()
        window._set_shared_key_status("naver", "ok", "등록됨 (확인함 · 전원 적용)")
        window._apply_shared_key_registration(None)
        self.assertEqual(self.row(window, "naver")[1], "ok")  # 이미 아는 상태는 지우지 않는다
        self.assertEqual(self.row(window, "coupang")[:2], ("확인 못 함 (공유 시트를 읽지 못함)", "idle"))

    def test_integration_page_reads_shared_config_once_when_visible(self):
        from PySide6.QtWidgets import QTabWidget, QWidget
        window = self.window()
        reads = []
        window._begin_tracking_config_read = lambda: reads.append(1)
        tabs = QTabWidget()
        first, settings = QWidget(), QWidget()
        tabs.addTab(first, "주문·발송")
        tabs.addTab(settings, "⚙ 설정")
        window.ui.tabWidget, window.ui.tab_settings = tabs, settings
        window._on_settings_page_shown(0)  # 시작할 때 분류 0을 고르는 것: 설정 탭이 안 보이면 읽지 않는다
        self.assertEqual(reads, [])
        tabs.setCurrentWidget(settings)
        window._on_settings_page_shown(1)  # 알림 페이지
        self.assertEqual(reads, [])
        window._on_settings_page_shown(0)
        self.assertEqual(reads, [1])
        window._apply_shared_key_registration({"postcode": True, "naver": True, "coupang": True})
        window._on_settings_page_shown(0)
        self.assertEqual(reads, [1])  # 한 번 읽은 뒤에는 다시 읽지 않는다

    def test_kpost_row_and_button(self):
        window = self.window()
        window._key_status_text = "미설정 — 「등록」을 눌러 주세요."
        window._update_status_displays()
        self.assertEqual(self.row(window, "kpost")[1:], ("warn", "등록"))
        window._key_status_text = "유효함 ✓ (저장됨 · 전원 적용)"
        window._update_status_displays()
        self.assertEqual(self.row(window, "kpost"), ("✓ 유효함 (저장됨 · 전원 적용)", "ok", "변경"))

    def test_key_validation_goes_to_key_row_not_inquiry_status(self):
        window = self.window()
        window._set_naver_inquiry_status("미답변 0건 · 새 알림 없음 · 10:00", "ok")
        writes = []
        window._tracking_controller = SimpleNamespace(queue_tracking_config_write=writes.append)
        window._naver_pending_creds = ("id", "secret")
        window._coupang_pending_creds = ("v", "a", "s")
        with patch.object(self.app.QMessageBox, "warning"):
            window._on_naver_creds_validated({"ok": True, "valid": False, "error": "401 Unauthorized"})
        self.assertEqual(self.row(window, "naver")[1], "error")
        self.assertEqual(window.ui.label_api_status_naver.toolTip(), "401 Unauthorized")
        window._on_coupang_creds_validated({"ok": True, "valid": True})
        self.assertEqual(self.row(window, "coupang")[1:], ("ok", "변경"))
        self.assertEqual(len(writes), 1)
        self.assertEqual(window._dlg_naver_status.text(), "✓ 미답변 0건 · 새 알림 없음 · 10:00")
        # 입력한 키의 검증 실패는 저장된 키 문제가 아니므로 분류 배지를 붙이지 않는다.
        window._update_settings_nav_badges()
        self.assertEqual(self.nav_texts(window)[0], "연동")


class SettingsUiFileTests(unittest.TestCase):
    def test_slack_webhook_button_moved_to_alerts_page(self):
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        alerts = tree.find(".//widget[@name='scrollArea_settings_alerts']")
        self.assertIsNotNone(alerts.find(".//widget[@name='groupBox_slack_webhook']//widget[@name='pushButton_cfg_slack_webhook']"))
        self.assertIsNotNone(alerts.find(".//widget[@name='label_slack_webhook_status']"))
        keys = tree.find(".//widget[@name='groupBox_api_keys']")
        self.assertIsNone(keys.find(".//widget[@name='pushButton_cfg_slack_webhook']"))

    def test_api_key_rows_in_grid(self):
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        grid = tree.find(".//widget[@name='groupBox_api_keys']/layout[@class='QGridLayout']")
        self.assertIsNotNone(grid)
        for key, button in (("kpost", "pushButton_cfg_kpost_key"), ("postcode", "pushButton_cfg_kpost_postcode_key"),
                            ("naver", "pushButton_cfg_naver_key"), ("coupang", "pushButton_cfg_coupang_key")):
            self.assertIsNotNone(grid.find(f".//widget[@name='label_api_status_{key}']"), key)
            self.assertIsNotNone(grid.find(f".//widget[@name='{button}']"), key)
        self.assertIsNone(tree.find(".//widget[@name='label_settings_api_status']"))

    def test_disconnect_is_last_google_button(self):
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        row = tree.find(".//layout[@name='horizontalLayout_google_buttons']")
        names = [item[0].get("name") for item in row.findall("./item")]
        self.assertEqual(names[-1], "pushButton_google_disconnect")
        self.assertEqual(names[-2], "horizontalSpacer_google_buttons")  # 다른 버튼과 떨어뜨린다
        self.assertIn("pushButton_google_open_sheet", names)

    def test_status_labels_have_no_widget_color(self):
        # 위젯 스타일시트가 앱 스타일시트(set_status 색)를 덮지 않게 한다.
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        for name in ("label_google_auth_status", "label_index_sheet_sync", "label_slack_webhook_status",
                     "label_api_status_kpost", "label_api_status_postcode", "label_api_status_naver",
                     "label_api_status_coupang"):
            widget = tree.find(f".//widget[@name='{name}']")
            self.assertIsNone(widget.find("./property[@name='styleSheet']"), name)


class SharedKeyRegistrationWorkerTests(unittest.TestCase):
    def test_config_read_reports_registration_without_values(self):
        import tracking.workers as workers
        cfg = {"kpost_regkey": "k", "naver_client_id": "id", "naver_client_secret": "sec",
               "coupang_vendor_id": "v", "coupang_access_key": "", "coupang_secret_key": "s"}
        with patch("sheets.google_sheets_oauth.get_authorized_gspread_client", return_value=object()), \
                patch.object(workers, "_standalone_open_config_ws", return_value=object()), \
                patch.object(workers, "_read_config_values_map", return_value=cfg):
            payload = workers.run_tracking_config_read_worker()
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["shared_keys"], {"postcode": False, "naver": True, "coupang": False})
        self.assertNotIn("sec", repr(payload))  # 비밀 값은 화면으로 넘기지 않는다


if __name__ == "__main__":
    unittest.main()
