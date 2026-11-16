"""배송추적 수취인 검색·정체 기준 변경 시 표를 글자·단계마다 다시 그리지 않는지 검증한다."""

import os
import time
import unittest
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QObject
from PySide6.QtWidgets import QApplication, QLineEdit, QTableWidget

from tracking.controller import TrackingController


class Host(QObject):
    def __init__(self):
        super().__init__()
        self._tracking_list_thread = None
        self._tracking_config_write_thread = object()  # 시트 쓰기 생략
        self.ui = SimpleNamespace(tableWidget_tracking=QTableWidget(),
                                  lineEdit_tracking_recipient_search=QLineEdit())
        self.settings = {}

    def get_app_setting(self, key, default=None):
        return self.settings.get(key, default)

    def set_app_setting(self, key, value):
        self.settings[key] = value


class SearchDebounceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def pump(self, seconds):
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(0.01)

    def test_typing_and_spinbox_steps_redraw_table_once(self):
        host = Host()
        controller = TrackingController(host)
        calls = []
        controller._populate_tracking_table = lambda *a: calls.append(
            host.ui.lineEdit_tracking_recipient_search.text())
        controller.bind_ui()
        for text in ("김", "김민", "김민수"):  # 빠르게 입력
            host.ui.lineEdit_tracking_recipient_search.setText(text)
            self.pump(0.05)
        self.pump(0.5)
        self.assertEqual(calls, ["김민수"])
        calls.clear()
        for value in (13, 14, 15):  # 스핀박스를 연속으로 올림
            controller._on_stale_threshold_changed("허브정체", value)
            self.pump(0.05)
        self.pump(0.5)
        self.assertEqual(len(calls), 1)
        self.assertEqual(host.settings["stale_hub_hours"], 15)


if __name__ == "__main__":
    unittest.main()
