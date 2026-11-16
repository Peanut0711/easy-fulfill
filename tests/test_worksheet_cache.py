"""송장추적·설정·문의알림 탭을 다시 열 때 스프레드시트 메타데이터를 다시 받지 않는지 검증한다."""

import contextlib
import importlib
import io
import unittest

from tests.test_order_load_sheets import FakeClient, FakeSheetsHTTP, load_app_module
from tracking.repository import open_config_worksheet, open_tracking_worksheet
from tracking.workers import TRACKING_SHEET_HEADERS

TRACKING = "송장추적"
CONFIG = "설정"


class CacheHTTP(FakeSheetsHTTP):
    """Worksheet.get()(row_values 등)이 요구하는 majorDimension을 응답에 넣는다."""

    def values_get(self, id, range, params=None):
        return {"majorDimension": "ROWS", **super().values_get(id, range, params)}


def worksheet_cache():
    try:
        return importlib.import_module("sheets.worksheet_cache")
    except ImportError:  # 수정 전 코드에는 캐시가 없다.
        return None


class WorksheetCacheTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        cache = worksheet_cache()
        if cache is not None:
            cache.clear()
            self.addCleanup(cache.clear)
        self.http = CacheHTTP([
            (TRACKING, [list(TRACKING_SHEET_HEADERS), ["6000000000000", "2026-11-17 09:00:00"]]),
            (CONFIG, [["키", "값"], ["slack_webhook_url", "https://hooks.test/x"]]),
            (self.app.INQUIRY_SHEET_TITLE, [list(self.app.INQUIRY_SHEET_HEADERS)]),
        ])
        self.gc = FakeClient(self.http)

    def open_all(self):
        tracking = open_tracking_worksheet(self.gc, "sid", TRACKING, TRACKING_SHEET_HEADERS)
        config = open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])
        with contextlib.redirect_stdout(io.StringIO()):
            inquiry = self.app._standalone_open_inquiry_ws(FakeClient(self.http))
        return tracking, config, inquiry

    def test_reopening_tabs_sends_no_metadata_requests(self):
        self.open_all()
        first = list(self.http.calls)
        self.assertEqual(first.count("meta"), 6)  # 탭마다 open_by_key·worksheets() 2회
        self.http.calls.clear()

        tracking, config, inquiry = self.open_all()
        self.assertEqual(self.http.calls, [])  # 두 번째부터는 요청 없이 연다
        self.assertEqual((tracking.title, tracking.id), (TRACKING, 10))
        self.assertEqual((config.title, config.id), (CONFIG, 11))
        self.assertEqual(inquiry.title, self.app.INQUIRY_SHEET_TITLE)
        # 캐시로 연 워크시트도 값 읽기는 그대로 동작한다.
        self.assertEqual(tracking.get_all_values()[1][0], "6000000000000")
        self.assertEqual(config.get_all_values()[1], ["slack_webhook_url", "https://hooks.test/x"])

    def test_cache_expires_and_metadata_is_fetched_again(self):
        cache = worksheet_cache()
        self.assertIsNotNone(cache, "sheets.worksheet_cache 모듈이 없습니다.")
        now = [1000.0]
        original = cache._monotonic
        cache._monotonic = lambda: now[0]
        self.addCleanup(setattr, cache, "_monotonic", original)
        open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])
        now[0] += cache.WORKSHEET_CACHE_TTL_SECONDS + 1
        self.http.calls.clear()
        open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])
        self.assertEqual(self.http.calls.count("meta"), 2)

    def test_objects_without_gspread_properties_are_not_cached(self):
        # 테스트용 가짜 워크시트처럼 속성 정보가 없는 객체는 캐시하지 않고 매번 연다.
        opened = []

        class Sheet:
            title = CONFIG

        class Spreadsheet:
            def worksheets(self):
                opened.append(1)
                return [Sheet()]

        gc = type("G", (), {"open_by_key": lambda self, key: Spreadsheet()})()
        open_config_worksheet(gc, "sid", CONFIG, ["키", "값"])
        open_config_worksheet(gc, "sid", CONFIG, ["키", "값"])
        self.assertEqual(len(opened), 2)


if __name__ == "__main__":
    unittest.main()
