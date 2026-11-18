"""느린 네트워크·엑셀 작업 동안 화면 이벤트 루프가 멈추지 않는지(심장박동 타이머) 검증한다."""

import os
import time
import unittest
from types import SimpleNamespace
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import requests
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QLabel

from shared.qt_background import run_in_background
from tests.test_order_load_sheets import load_app_module


class Heartbeat:
    """20ms마다 시각을 기록해 화면 스레드가 가장 오래 멈춘 시간을 잰다."""

    def __init__(self):
        self.stamps = []
        self.timer = QTimer()
        self.timer.setInterval(20)
        self.timer.timeout.connect(lambda: self.stamps.append(time.monotonic()))

    def __enter__(self):
        self.started = time.monotonic()
        self.timer.start()
        return self

    def __exit__(self, *exc):
        self.ended = time.monotonic()
        self.timer.stop()

    def longest_gap(self):
        points = [self.started] + self.stamps + [self.ended]
        return max(b - a for a, b in zip(points, points[1:]))


def slow(seconds, value=None, error=None):
    def inner(*args, **kwargs):
        time.sleep(seconds)
        if error:
            raise error
        return value
    return inner


class BackgroundHelperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_ui_keeps_running_and_results_errors_pass_through(self):
        with Heartbeat() as beat:
            self.assertEqual(run_in_background(slow(0.6, value=42)), 42)
        self.assertLess(beat.longest_gap(), 0.2)
        with self.assertRaises(ValueError):
            run_in_background(slow(0.1, error=ValueError("실패")))


class CallSiteTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app_module = load_app_module()

    def test_postcode_search_does_not_freeze_and_hides_key(self):
        main_window = self.app_module.MainWindow
        host = SimpleNamespace(_kpost_parse_postnew_response=lambda text: (["주소"], None))
        response = SimpleNamespace(raise_for_status=lambda: None, content="<xml/>".encode())
        with patch.object(requests, "get", slow(0.6, value=response)), Heartbeat() as beat:
            items, err = main_window._request_kpost_postnew(host, "SECRETREGKEY", "서울 강남구")
        self.assertEqual(items, ["주소"])
        self.assertLess(beat.longest_gap(), 0.2)  # 수정 전에는 요청 시간(0.6초)만큼 멈춤
        failure = requests.ConnectionError("Max retries exceeded with url: /openapi2?regkey=SECRETREGKEY&query=x")
        with patch.object(requests, "get", side_effect=failure):
            items, err = main_window._request_kpost_postnew(host, "SECRETREGKEY", "서울 강남구")
        self.assertIsNone(items)
        self.assertNotIn("SECRETREGKEY", str(err))

    def test_image_dialog_download_has_timeout_and_does_not_freeze(self):
        seen = {}

        def fake_get(url, **kwargs):
            seen.update(kwargs)
            time.sleep(0.6)
            return SimpleNamespace(content=b"", raise_for_status=lambda: None)

        host = SimpleNamespace(image_label=QLabel())
        with patch.object(requests, "get", fake_get), Heartbeat() as beat:
            self.app_module.ImageDialog.load_image(host, "https://image.test/a.jpg")
        self.assertLess(beat.longest_gap(), 0.2)
        self.assertIn("timeout", seen)


if __name__ == "__main__":
    unittest.main()
