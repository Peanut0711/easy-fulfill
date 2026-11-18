"""상태바 상시 표시(알림·상태바 계획 C-2)를 검증한다: Google 연결 · 인덱스 동기화 · 앱 버전."""

import os
import sys
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class StatusChipTests(unittest.TestCase):
    METHODS = (
        "_setup_status_chips", "_update_app_version_label", "_refresh_google_auth_status_ui",
        "_mark_google_auth_runtime_error", "_update_settings_nav_badges", "_settings_integration_problem",
        "_key_status_kind", "_update_index_sheet_sync_label", "_update_index_status_chip",
    )

    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from tests.test_order_load_sheets import load_app_module
        cls.module = load_app_module()

    def window(self):
        from PySide6.QtCore import QTimer
        from PySide6.QtWidgets import QLabel, QMainWindow
        attrs = {name: vars(self.module.MainWindow)[name] for name in self.METHODS}
        window = type("FakeWindow", (QMainWindow,), attrs)()
        window.ui = SimpleNamespace(label_app_version=QLabel(), label_index_sheet_sync=QLabel())
        window.pages = []
        window.show_settings_page = window.pages.append
        window._key_status_text = "유효함 ✓"
        window._google_auth_runtime_error = ""
        window._index_sheet_push_timer = QTimer()
        window._index_sheet_push_pending = False
        window._index_sheet_last_sync_display = None
        window._setup_status_chips()
        return window

    def chip(self, window, name):
        chip = getattr(window, name)
        return chip.text(), chip.property("status")

    def test_chips_on_status_bar_and_click_opens_settings(self):
        window = self.window()
        bar = window.statusBar()
        for name in ("statusChip_google", "statusChip_index", "statusChip_version"):
            self.assertIs(getattr(window, name).parent(), bar)
        self.assertEqual(self.chip(window, "statusChip_version"), (f"v{self.module.APP_VERSION}", "idle"))
        for name in ("statusChip_google", "statusChip_index", "statusChip_version"):
            getattr(window, name).clicked.emit()
        self.assertEqual(window.pages, ["integration", "orders", "program"])

    def test_google_chip_follows_auth_status(self):
        window = self.window()
        with patch("sheets.google_sheets_oauth.get_oauth_status", return_value=("ok", "연결됨", "")):
            window._refresh_google_auth_status_ui()
            self.assertEqual(self.chip(window, "statusChip_google"), ("✓ Google", "ok"))
            window._mark_google_auth_runtime_error("invalid_grant")
            self.assertEqual(self.chip(window, "statusChip_google"), ("✗ Google", "error"))
        with patch("sheets.google_sheets_oauth.get_oauth_status", return_value=("warn", "로그인이 없습니다.", "")):
            window._google_auth_runtime_error = ""
            window._refresh_google_auth_status_ui()
            self.assertEqual(self.chip(window, "statusChip_google")[1], "warn")
            self.assertIn("로그인이 없습니다", window.statusChip_google.toolTip())

    def test_index_chip_states(self):
        window = self.window()
        window._update_index_sheet_sync_label()
        self.assertEqual(self.chip(window, "statusChip_index"), ("인덱스 —", "idle"))
        window._index_sheet_last_sync_display = datetime(2026, 11, 18, 10, 21, 33)
        window._update_index_sheet_sync_label()
        self.assertEqual(self.chip(window, "statusChip_index"), ("✓ 인덱스 10:21", "ok"))
        self.assertIn("10:21:33", window.statusChip_index.toolTip())
        window._index_sheet_push_pending = True
        window._update_index_sheet_sync_label()
        self.assertEqual(self.chip(window, "statusChip_index"), ("✗ 인덱스 반영 실패", "error"))
        window._update_index_status_chip(syncing=True)
        self.assertEqual(self.chip(window, "statusChip_index"), ("인덱스 동기화 중…", "idle"))

    def test_version_chip_states(self):
        window = self.window()
        window._update_app_version_label({"ok": True, "version": "0.0.19", "latest": "0.0.20"})
        self.assertEqual(self.chip(window, "statusChip_version"), ("! v0.0.19 → 0.0.20", "warn"))
        window._update_app_version_label({"ok": True, "version": "0.0.19", "blocked": True, "minimum": "0.0.20"})
        self.assertEqual(self.chip(window, "statusChip_version"), ("✗ v0.0.19 · 업데이트 필요", "error"))
        window._update_app_version_label({"ok": True, "version": "0.0.19"})
        self.assertEqual(self.chip(window, "statusChip_version"), ("v0.0.19", "idle"))
        self.assertIn("설정 › 프로그램", window.statusChip_version.toolTip())


if __name__ == "__main__":
    unittest.main()
