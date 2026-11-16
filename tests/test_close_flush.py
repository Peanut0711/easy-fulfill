"""창을 닫을 때 대기 중인 시트 반영(주문번호 인덱스·송장추적 등록·메모)을 마치고 종료하는지 검증한다.

MainWindow를 만든 프로세스는 정상 인터프리터 종료 시 Qt 정리 순서 때문에 비정상 종료되므로
(앱도 같은 이유로 os._exit로 끝낸다) 각 시나리오를 별도 프로세스에서 실행하고 결과만 받는다.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run_scenario(name):
    """자식 프로세스에서 시나리오를 실행한다. 결과 JSON을 표준 출력으로 받는다."""
    import importlib.util
    import time
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS", "--disable-gpu")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtCore import QThread, Signal
    from PySide6.QtWidgets import QApplication, QMessageBox

    import sheets.google_sheets_oauth as oauth
    import tracking.controller as tracking_controller

    writes = []
    fail_index = name == "persistent_failure"

    class FakeIndexWriteThread(QThread):
        result_ready = Signal(dict)

        def __init__(self, naver, coupang, gmarket, parent=None, reason="변경", expected=None):
            super().__init__(parent)
            self.values = (naver, coupang, gmarket)

        def run(self):
            time.sleep(0.2)
            if fail_index:
                self.result_ready.emit({"ok": False, "error": "네트워크 오류"})
                return
            writes.append(["index", list(self.values)])
            self.result_ready.emit({"ok": True})

    class FakeUpsertThread(QThread):
        result_ready = Signal(dict)

        def __init__(self, records, parent=None):
            super().__init__(parent)
            self.records = records

        def run(self):
            time.sleep(0.1)
            writes.append(["tracking", [record["등기번호"] for record in self.records]])
            self.result_ready.emit({"ok": True, "registered": len(self.records), "updated": 0})

    class FakeNotesThread(QThread):
        result_ready = Signal(dict)

        def __init__(self, notes, parent=None):
            super().__init__(parent)
            self.notes = notes

        def run(self):
            writes.append(["notes", dict(self.notes)])
            self.result_ready.emit({"ok": True})

    spec = importlib.util.spec_from_file_location("easy_fulfill_close_scenario", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    asked = []

    def answer(*args, **kwargs):
        asked.append(args[2] if len(args) > 2 else "")
        return QMessageBox.StandardButton.No

    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("테스트: 네트워크 금지")), \
            patch.object(app_module, "OrderIndexWriteThread", FakeIndexWriteThread), \
            patch.object(tracking_controller, "TrackingUpsertThread", FakeUpsertThread), \
            patch.object(tracking_controller, "TrackingNotesUpdateThread", FakeNotesThread), \
            patch.object(app_module, "CLOSE_FLUSH_TIMEOUT_MS", 1500, create=True), \
            patch.object(QMessageBox, "question", side_effect=answer):
        window = app_module.MainWindow()
        window._startup_order_index_sync_started = True  # 시작 동기화(실제 시트 읽기) 생략
        if name in ("flush_all", "persistent_failure"):
            window.current_idx_naver = 21
            window._schedule_order_index_sheet_push()  # 800ms 디바운스 대기 상태
        if name == "flush_all":
            window._tracking_pending_records = [
                {"등기번호": "6012345678901", "스토어": "", "주문번호": "", "수취인명": "김"}]
            window._tracking_push_timer.start(10_000)
            window._tracking_notes_pending = {"6012345678901": "경비실"}
        window.show()
        started = time.monotonic()
        window.close()
        deadline = time.monotonic() + 8
        while window.isVisible() and time.monotonic() < deadline:
            app.processEvents()
            time.sleep(0.02)
        result = {
            "closed": not window.isVisible(),
            "seconds": time.monotonic() - started,
            "writes": writes,
            "asked": asked,
            "poll_timer_active": window._index_sheet_poll_timer.isActive(),
            "tracking_list_timer_active": window._tracking_list_timer.isActive(),
        }
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


class CloseFlushTests(unittest.TestCase):
    def scenario(self, name):
        with tempfile.TemporaryDirectory() as directory:
            env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
            completed = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), "--scenario", name],
                cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=120,
            )
        line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
        self.assertIsNotNone(line, completed.stdout[-2000:] + completed.stderr[-2000:])
        return json.loads(line[len("RESULT "):])

    def test_close_flushes_pending_index_tracking_and_notes(self):
        result = self.scenario("flush_all")
        self.assertTrue(result["closed"])
        writes = dict((kind, value) for kind, value in result["writes"])
        self.assertEqual(writes["index"][0], 21)
        self.assertEqual(writes["tracking"], ["6012345678901"])
        self.assertEqual(writes["notes"], {"6012345678901": "경비실"})
        self.assertFalse(result["poll_timer_active"])
        self.assertFalse(result["tracking_list_timer_active"])
        self.assertEqual(result["asked"], [])

    def test_close_without_pending_writes_is_immediate(self):
        result = self.scenario("nothing_pending")
        self.assertTrue(result["closed"])
        self.assertLess(result["seconds"], 1.0)
        self.assertEqual(result["writes"], [])

    def test_persistent_failure_asks_before_closing(self):
        result = self.scenario("persistent_failure")
        self.assertTrue(result["closed"])
        self.assertEqual(len(result["asked"]), 1)
        self.assertIn("주문번호 인덱스", result["asked"][0])


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--scenario":
        run_scenario(sys.argv[2])
    else:
        unittest.main()
