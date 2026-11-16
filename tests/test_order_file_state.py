"""주문 파일 처리 실패·선택 취소가 주문 상태(유효 여부·이전 주문·스토어·중복 기록)를 망가뜨리지 않는지 검증한다."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run_scenario(name):
    """MainWindow 시나리오(자식 프로세스, os._exit로 종료)."""
    import importlib.util
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

    import sheets.google_sheets_oauth as oauth

    work = Path.cwd()
    spec = importlib.util.spec_from_file_location("easy_fulfill_order_state", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(QMessageBox, "warning", side_effect=lambda *a, **k: None), \
            patch.object(QMessageBox, "critical", side_effect=lambda *a, **k: None):
        window = app_module.MainWindow()
        window._startup_order_index_sync_started = True
        # 인덱스 읽기는 생략하고 바로 파일 처리로 이어간다.
        window._begin_order_index_read = lambda **kwargs: kwargs["on_applied"] and kwargs["on_applied"]()
        if name == "failed_processing":
            path = work / "발송관리.xlsx"
            path.write_bytes(b"gmarket order file")
            window.orders = {"이전 주문": {"수취인명": "김"}}
            window.is_order_file_valid = True
            with patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(path), "")), \
                    patch.object(app_module.MainWindow, "process_gmarket_excel_file", lambda self: None):
                window.select_excel_file()  # 처리 실패(누락 열 등으로 None 반환)
            loaded = window._todays_loaded_hashes()[1]
            result.update(valid=window.is_order_file_valid, orders=window.orders,
                          recorded=window._file_sha1(str(path)) in loaded)
        elif name == "cancel_keeps_state":
            window.store_type = "naver"
            window.orders = {"주문": {"수취인명": "김"}}
            window.is_order_file_valid = True
            with patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: ("", "")):
                window.select_excel_file()
            result["after_dialog_cancel"] = [window.store_type, window.is_order_file_valid, len(window.orders)]
            duplicate = work / "DeliveryList(2026-11-16)_(0).xlsx"
            duplicate.write_bytes(b"coupang order file")
            window._record_loaded_order(str(duplicate))
            with patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(duplicate), "")), \
                    patch.object(QMessageBox, "question", return_value=QMessageBox.StandardButton.No):
                window.select_excel_file()  # 중복 확인에서 취소
            result["after_duplicate_cancel"] = [window.store_type, window.is_order_file_valid, len(window.orders)]
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


class OrderFileStateTests(unittest.TestCase):
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

    def test_failed_processing_is_not_valid_clears_orders_and_is_not_recorded(self):
        result = self.scenario("failed_processing")
        self.assertFalse(result["valid"])
        self.assertEqual(result["orders"], {})
        self.assertFalse(result["recorded"])

    def test_cancel_keeps_loaded_orders_and_store(self):
        result = self.scenario("cancel_keeps_state")
        self.assertEqual(result["after_dialog_cancel"], ["naver", True, 1])
        self.assertEqual(result["after_duplicate_cancel"], ["naver", True, 1])


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--scenario":
        run_scenario(sys.argv[2])
    else:
        unittest.main()
