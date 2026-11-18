"""주문 불러오기 시 Google Sheets 조회 통합(메타데이터 1회 + 값 일괄 조회 1회)을 가짜 API로 검증한다."""

import contextlib
import importlib.util
import io
import json
import os
import re
import unittest
from datetime import date
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS", "--disable-gpu")

from gspread import Spreadsheet
from gspread.http_client import HTTPClient
from PySide6.QtCore import QCoreApplication
from PySide6.QtWidgets import QApplication

import sheets.google_sheets_oauth as oauth
from sheets import worksheet_cache
from sheets.batch_read import batch_get_all_values, fetch_sheet_properties, worksheet_from_properties


ROOT = Path(__file__).resolve().parents[1]
TODAY = date.today().strftime("%Y-%m-%d")


def load_app_module():
    QApplication.instance() or QApplication([])
    spec = importlib.util.spec_from_file_location("easy_fulfill_for_sheet_tests", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main_window_members(module):
    """MainWindow와 그 믹스인(ui/main/)에 정의된 이름 → 클래스 __dict__ 값(staticmethod 등 설명자 그대로).

    메인 창 코드가 믹스인으로 나뉘어 있어 ``vars(MainWindow)``만으로는 옮긴 메서드를 찾을 수 없다.
    Qt 기반 클래스(QMainWindow 등)의 이름은 넣지 않는다.
    """
    members = {}
    for cls in reversed(module.MainWindow.__mro__):
        if cls is module.MainWindow or cls.__module__.startswith("ui.main"):
            members.update(vars(cls))
    return members


def _title(range_name):
    match = re.match(r"^'((?:[^']|'')*)'(?:!(.*))?$", range_name)
    return (match.group(1).replace("''", "'"), match.group(2)) if match else (range_name, None)


class FakeSheetsHTTP(HTTPClient):
    """Sheets API 요청을 세고 메모리의 탭 값을 돌려주는 HTTPClient."""

    def __init__(self, tabs, failing_titles=()):
        self.tabs = [{"title": title, "values": values, "sheetId": 10 + i} for i, (title, values) in enumerate(tabs)]
        self.failing_titles = set(failing_titles)
        self.calls = []

    def _tab(self, title):
        return next(tab for tab in self.tabs if tab["title"] == title)

    def _values(self, title):
        if title in self.failing_titles:
            raise RuntimeError(f"읽기 실패: {title}")
        return [list(row) for row in self._tab(title)["values"]]

    def fetch_sheet_metadata(self, id, params=None):
        self.calls.append("meta")
        return {"properties": {"title": "fake"}, "sheets": [
            {"properties": {"title": tab["title"], "sheetId": tab["sheetId"], "index": i,
                            "gridProperties": {"rowCount": 1000, "columnCount": 26}}}
            for i, tab in enumerate(self.tabs)]}

    def values_get(self, id, range, params=None):
        title, _ = _title(range)
        self.calls.append(f"get:{title}")
        values = self._values(title)
        return {"range": range, **({"values": values} if values else {})}

    def values_batch_get(self, id, ranges, params=None):
        titles = [_title(range_name)[0] for range_name in ranges]
        self.calls.append("batch:" + "|".join(titles))
        value_ranges = []
        for range_name, title in zip(ranges, titles):
            values = self._values(title)
            value_ranges.append({"range": range_name, **({"values": values} if values else {})})
        return {"valueRanges": value_ranges}

    def values_append(self, id, range, params=None, body=None):
        title, _ = _title(range)
        self.calls.append(f"append:{title}")
        self._tab(title)["values"].extend(body["values"])
        return {}

    def request(self, *args, **kwargs):
        raise AssertionError(f"예상하지 못한 요청: {args}")


class FakeClient:
    def __init__(self, http):
        self.http_client = http

    def open_by_key(self, key):
        return Spreadsheet(self.http_client, {"id": key})


def spreadsheet_tabs(today_row=True):
    index = [["날짜", "네이버", "쿠팡", "지마켓"], ["2026-01-01", "5", "6", "7"]]
    if today_row:
        index.append([TODAY, "12", "8", "3"])
    gift = json.dumps({"enabled": True, "stores": ["naver"], "tiers": [[60000, "볼펜"]]}, ensure_ascii=False)
    return [
        ("네이버DB", [["상품코드", "", "", "", "상품번호"], ["N1", "", "", "", "1001"], ["N2"]]),
        ("쿠팡DB", [["제목"], ["상품코드", "", "", "VP", "옵션ID"], ["C1", "", "", "7001", "9001.0"]]),
        ("일별 주문번호", index),
        ("설정", [["키", "값"], ["order_gift_event", gift]]),
    ]


class BatchReadTests(unittest.TestCase):
    def test_batch_values_match_get_all_values_shape(self):
        http = FakeSheetsHTTP([("A", [["1", "2"], ["3"]]), ("B", [])])
        gc = FakeClient(http)
        props = fetch_sheet_properties(gc, "id")
        self.assertEqual([p["title"] for p in props], ["A", "B"])
        values = batch_get_all_values(gc, "id", ["A", "B"])
        worksheet = gc.open_by_key("id").worksheet("A")
        self.assertEqual(values[0], worksheet.get_all_values())
        self.assertEqual(values[0], [["1", "2"], ["3", ""]])
        self.assertEqual(values[1], [[]])
        self.assertEqual(worksheet_from_properties(gc, "id", props[1]).title, "B")


class OrderLoadSheetsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app_module = load_app_module()
        cls.original_client = oauth.get_authorized_gspread_client

    @classmethod
    def tearDownClass(cls):
        oauth.get_authorized_gspread_client = cls.original_client

    def run_worker(self, http, *, legacy, store):
        # 탭 속성 캐시는 프로세스 전체에서 공유되므로 매 실행을 캐시 없이 시작한다(요청 수 비교 기준).
        worksheet_cache.clear()
        oauth.get_authorized_gspread_client = lambda: FakeClient(http)
        with contextlib.redirect_stdout(io.StringIO()):
            if legacy:
                payload = self.app_module.run_order_index_read_sync_worker(include_gift_event=True)
                payload["mapping"] = self.app_module.run_product_code_map_load_worker(store)
                return payload
            return self.app_module.run_order_load_sheets_worker(store)

    def assert_same_as_legacy(self, store, *, today_row=True, failing_titles=()):
        legacy_http = FakeSheetsHTTP(spreadsheet_tabs(today_row), failing_titles)
        batched_http = FakeSheetsHTTP(spreadsheet_tabs(today_row), failing_titles)
        legacy = self.run_worker(legacy_http, legacy=True, store=store)
        batched = self.run_worker(batched_http, legacy=False, store=store)
        self.assertEqual(batched, legacy)
        self.assertEqual([tab["values"] for tab in batched_http.tabs], [tab["values"] for tab in legacy_http.tabs])
        return legacy_http.calls, batched_http.calls, batched

    def test_naver_and_coupang_use_one_metadata_and_one_batch_read(self):
        for store, mapping_tab in (("naver", "네이버DB"), ("coupang", "쿠팡DB")):
            legacy_calls, calls, payload = self.assert_same_as_legacy(store)
            self.assertEqual(len(legacy_calls), 9)
            self.assertEqual(calls, ["meta", f"batch:설정|일별 주문번호|{mapping_tab}"])
            self.assertEqual(payload["row"], {"naver": 12, "coupang": 8, "gmarket": 3})
            self.assertTrue(payload["mapping"]["ok"])

    def test_today_row_is_created_with_one_extra_request(self):
        _, calls, payload = self.assert_same_as_legacy("naver", today_row=False)
        self.assertEqual(payload["kind"], "created")
        self.assertEqual(calls[-1], "append:일별 주문번호")
        self.assertEqual(len(calls), 3)

    def test_failed_batch_falls_back_to_per_sheet_results(self):
        for failing in ("설정", "쿠팡DB"):
            _, _, payload = self.assert_same_as_legacy("coupang", failing_titles=[failing])
            self.assertTrue(payload["ok"])
            self.assertEqual("gift_event_error" in payload, failing == "설정")
            self.assertEqual(payload["mapping"]["ok"], failing != "쿠팡DB")

    def test_index_is_read_fresh_on_every_order_load(self):
        http = FakeSheetsHTTP(spreadsheet_tabs())
        self.assertEqual(self.run_worker(http, legacy=False, store="naver")["row"]["naver"], 12)
        http._tab("일별 주문번호")["values"][2] = [TODAY, "30", "31", "32"]
        self.assertEqual(self.run_worker(http, legacy=False, store="naver")["row"]["naver"], 30)

    def test_order_processing_uses_prefetched_mapping_without_new_thread(self):
        main_window = self.app_module.MainWindow
        received = []
        host = type("Host", (), {})()
        host._prefetched_order_mapping = {"ok": True, "store_type": "naver", "mapping": {"1001": "N1"}}
        host._on_product_mapping_for_order_file_ready = received.append
        original_thread = self.app_module.ProductMappingLoadThread
        self.app_module.ProductMappingLoadThread = None  # 새 매핑 스레드를 만들면 실패한다.
        try:
            main_window._run_order_file_processing_with_async_mapping(host, "naver")
            QCoreApplication.processEvents()
        finally:
            self.app_module.ProductMappingLoadThread = original_thread
        self.assertEqual(received, [{"ok": True, "store_type": "naver", "mapping": {"1001": "N1"}}])
        self.assertIsNone(host._prefetched_order_mapping)


if __name__ == "__main__":
    unittest.main()
