"""공유 「설정」 탭 쓰기가 진행 중일 때 들어온 변경도 마지막 값까지 시트에 저장되는지 검증한다."""

import os
import time
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QObject, QThread, Signal
from PySide6.QtWidgets import QApplication

import tracking.controller as controller_module
from tracking.controller import TrackingController

WRITES = []


class FakeConfigWriteThread(QThread):
    result_ready = Signal(dict)

    def __init__(self, updates, parent=None):
        super().__init__(parent)
        self.updates = dict(updates)

    def run(self):
        time.sleep(0.15)  # 시트 쓰기 중에 다음 변경이 들어오도록 지연
        WRITES.append(self.updates)
        self.result_ready.emit({"ok": True})


class Host(QObject):
    def __init__(self):
        super().__init__()
        self.ui = object()
        self._tracking_list_thread = None
        self.settings = {}

    def get_app_setting(self, key, default=None):
        return self.settings.get(key, default)

    def set_app_setting(self, key, value):
        self.settings[key] = value


class ConfigWriteQueueTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def wait_idle(self, host, seconds=5):
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            self.app.processEvents()
            if host._tracking_config_write_thread is None and not getattr(
                    host, "_tracking_config_pending_writes", None):
                return
            time.sleep(0.01)

    def test_rapid_changes_end_with_last_value_in_sheet(self):
        WRITES.clear()
        host = Host()
        controller = TrackingController(host)
        controller.initialize_runtime()
        with patch.object(controller_module, "TrackingConfigWriteThread", FakeConfigWriteThread):
            for value in (10, 11, 12, 13):
                controller._on_stale_threshold_changed("허브정체", value)
            controller._on_remote_bonus_changed(5)
            self.wait_idle(host)
        self.assertGreaterEqual(len(WRITES), 1)
        merged = {}
        for updates in WRITES:
            merged.update(updates)
        self.assertEqual(merged["stale_hub_hours"], "13")
        self.assertEqual(merged["stale_remote_bonus_hours"], "5")
        self.assertLessEqual(len(WRITES), 2)  # 진행 중 들어온 변경은 모아서 한 번 더 쓴다
        self.assertIsNone(host._tracking_config_write_thread)


if __name__ == "__main__":
    unittest.main()
