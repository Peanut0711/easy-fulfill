"""여러 PC가 공유하는 주문번호 인덱스가 서로의 변경을 덮어쓰지 않는지 검증한다."""

import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))  # 자식 프로세스(시나리오)는 임시 폴더에서 실행된다.

from orders.index_sync import IndexConflict, conflict_message, merge_order_indices  # noqa: E402
TODAY = date.today().strftime("%Y-%m-%d")


def values(n, c, g):
    return {"naver": n, "coupang": c, "gmarket": g}


class MergeRuleTests(unittest.TestCase):
    def test_untouched_store_follows_sheet_and_changed_store_keeps_local(self):
        merged, conflicts = merge_order_indices(values(12, 5, 3), values(10, 5, 3), values(10, 8, 3), local_dirty=True)
        self.assertEqual(merged, values(12, 8, 3))
        self.assertEqual(conflicts, [])

    def test_both_changed_same_store_keeps_larger_and_reports(self):
        merged, conflicts = merge_order_indices(values(12, 5, 3), values(10, 5, 3), values(15, 5, 3), local_dirty=True)
        self.assertEqual(merged["naver"], 15)
        self.assertEqual(conflicts, [IndexConflict("naver", 10, 12, 15, 15)])
        self.assertIn("10~11번", conflict_message(conflicts))

    def test_clean_local_follows_sheet(self):
        merged, conflicts = merge_order_indices(values(9, 9, 9), values(1, 1, 1), values(4, 5, 6), local_dirty=False)
        self.assertEqual(merged, values(4, 5, 6))
        self.assertEqual(conflicts, [])


class FakeIndexHTTP:
    """tests.test_order_load_sheets의 가짜 Sheets API에 쓰기(values_update)를 더한다."""

    @staticmethod
    def make(tabs):
        from tests.test_order_load_sheets import FakeSheetsHTTP, _title

        class WritableHTTP(FakeSheetsHTTP):
            def values_update(self, id, range, params=None, body=None):
                title, a1 = _title(range)
                self.calls.append(f"update:{title}!{a1}")
                tab = self._tab(title)["values"]
                import re
                match = re.match(r"([A-Z])(\d+)", a1)
                col, row = ord(match.group(1)) - 65, int(match.group(2)) - 1
                for r, line in enumerate(body["values"]):
                    while len(tab) <= row + r:
                        tab.append([])
                    target = tab[row + r]
                    for c, value in enumerate(line):
                        while len(target) <= col + c:
                            target.append("")
                        target[col + c] = str(value)
                return {}

        return WritableHTTP(tabs)


def index_tabs(naver, coupang, gmarket, index_position=2):
    tabs = [("네이버DB", [["코드"]]), ("쿠팡DB", [["코드"]]), ("설정", [["키", "값"]])]
    index_tab = ("일별 주문번호", [["날짜", "네이버", "쿠팡", "지마켓"], [TODAY, str(naver), str(coupang), str(gmarket)]])
    tabs.insert(index_position, index_tab)
    tabs.append(("주문번호 변경이력", [["시각", "사용자", "스토어", "이전값", "새값", "사유"]]))
    return tabs


class WriteWorkerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from tests.test_order_load_sheets import load_app_module
        import sheets.google_sheets_oauth as oauth
        cls.app = load_app_module()
        cls.oauth = oauth
        cls.original = oauth.get_authorized_gspread_client

    @classmethod
    def tearDownClass(cls):
        cls.oauth.get_authorized_gspread_client = cls.original

    def write(self, http, local, expected):
        from tests.test_order_load_sheets import FakeClient
        self.oauth.get_authorized_gspread_client = lambda: FakeClient(http)
        with contextlib.redirect_stdout(io.StringIO()):
            return self.app.run_order_index_write_worker(
                local["naver"], local["coupang"], local["gmarket"], reason="주문 로드", expected=expected)

    def sheet_row(self, http):
        row = http._tab("일별 주문번호")["values"][1]
        return values(*(int(cell) for cell in row[1:4]))

    def test_does_not_roll_back_other_pc_change(self):
        http = FakeIndexHTTP.make(index_tabs(10, 8, 3))  # 다른 PC가 쿠팡을 5→8로 바꾼 상태
        result = self.write(http, values(12, 5, 3), values(10, 5, 3))
        self.assertTrue(result["ok"], result)
        self.assertEqual(self.sheet_row(http), values(12, 8, 3))
        self.assertEqual(result["written"], values(12, 8, 3))
        self.assertEqual(result["conflicts"], [])

    def test_same_store_collision_keeps_larger_and_reports(self):
        http = FakeIndexHTTP.make(index_tabs(15, 5, 3))
        result = self.write(http, values(12, 5, 3), values(10, 5, 3))
        self.assertEqual(self.sheet_row(http), values(15, 5, 3))
        self.assertEqual([item["store"] for item in result["conflicts"]], ["naver"])

    def test_index_tab_is_found_by_title_not_position(self):
        http = FakeIndexHTTP.make(index_tabs(10, 5, 3, index_position=3))  # 2번 위치는 「설정」
        result = self.write(http, values(11, 5, 3), values(10, 5, 3))
        self.assertTrue(result["ok"], result)
        self.assertEqual(self.sheet_row(http), values(11, 5, 3))
        self.assertEqual(http._tab("설정")["values"], [["키", "값"]])


def run_scenario(name):
    """MainWindow 시나리오(자식 프로세스). MainWindow 프로세스는 os._exit로 끝낸다."""
    import importlib.util
    import time
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtCore import QThread, Signal
    from PySide6.QtWidgets import QApplication, QMessageBox

    import sheets.google_sheets_oauth as oauth

    sent = []

    class FailingWriteThread(QThread):
        result_ready = Signal(dict)

        def __init__(self, naver, coupang, gmarket, parent=None, reason="변경", expected=None):
            super().__init__(parent)
            sent.append({"values": [naver, coupang, gmarket], "expected": expected})

        def run(self):
            self.result_ready.emit({"ok": False, "error": "네트워크 오류"})

    spec = importlib.util.spec_from_file_location("easy_fulfill_index_scenario", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    warnings = []
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(app_module, "OrderIndexWriteThread", FailingWriteThread), \
            patch.object(QMessageBox, "warning", side_effect=lambda *a, **k: warnings.append(a[2] if len(a) > 2 else "")):
        window = app_module.MainWindow()
        window._startup_order_index_sync_started = True
        payload = {"ok": True, "kind": "data", "row": values(10, 5, 3)}
        if name == "startup_keeps_pending_local":
            window._apply_order_index_sync_payload(payload, prompt_reauth=False)  # 기준값 10/5/3
            window.current_idx_naver = 12
            window.save_index_values(reason="주문 로드")  # 디바운스 대기 중인 로컬 변경
            window._apply_order_index_sync_payload(
                {"ok": True, "kind": "data", "row": values(10, 8, 3)}, prompt_reauth=False)
            result["after"] = [window.current_idx_naver, window.current_idx_coupang, window.current_idx_gmarket]
            result["pending"] = window._index_sheet_push_pending or window._index_sheet_push_timer.isActive()
        elif name == "failed_write_retries":
            window._apply_order_index_sync_payload(payload, prompt_reauth=False)
            window.current_idx_naver = 11
            window.save_index_values(reason="주문 로드")
            window._flush_order_indices_to_sheet()
            deadline = time.monotonic() + 3
            while window._index_sheet_op_thread is not None and time.monotonic() < deadline:
                app.processEvents()
                time.sleep(0.01)
            result["sent"] = sent
            result["pending"] = window._index_sheet_push_pending
            result["retry_scheduled"] = window._index_sheet_push_timer.isActive()
        elif name == "read_keeps_pending_push":
            window._apply_order_index_sync_payload(payload, prompt_reauth=False)
            window.current_idx_naver = 11
            window.save_index_values(reason="주문 로드")
            with patch.object(app_module, "OrderIndexReadSyncThread") as reader:
                reader.return_value.result_ready.connect = lambda *a: None
                reader.return_value.finished.connect = lambda *a: None
                window._begin_order_index_read(interactive=False)
            result["pending"] = window._index_sheet_push_pending
        result["warnings"] = warnings
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


class MainWindowIndexTests(unittest.TestCase):
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

    def test_sync_payload_merges_instead_of_overwriting_pending_local_change(self):
        result = self.scenario("startup_keeps_pending_local")
        self.assertEqual(result["after"], [12, 8, 3])
        self.assertTrue(result["pending"])

    def test_failed_write_is_retried_with_expected_base(self):
        result = self.scenario("failed_write_retries")
        self.assertEqual(result["sent"][0]["expected"], values(10, 5, 3))
        self.assertTrue(result["pending"])
        self.assertTrue(result["retry_scheduled"])

    def test_starting_a_read_keeps_debounced_push_pending(self):
        self.assertTrue(self.scenario("read_keeps_pending_push")["pending"])


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--scenario":
        run_scenario(sys.argv[2])
    else:
        unittest.main()
