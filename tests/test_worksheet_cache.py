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


class FakeResponse:
    def __init__(self, code, message, status):
        self.status_code = code
        self._body = {"error": {"code": code, "message": message, "status": status}}
        self.text = str(self._body)

    def json(self):
        return self._body


class CacheHTTP(FakeSheetsHTTP):
    """Worksheet.get()(row_values 등)이 요구하는 majorDimension을 응답에 넣고, 실제 API처럼
    없는 탭 이름은 400 「Unable to parse range」로 거절하며, 탭 추가·헤더 쓰기를 지원한다."""

    forbidden_titles = ()

    def values_get(self, id, range, params=None):
        from gspread.exceptions import APIError
        from tests.test_order_load_sheets import _title
        title, _ = _title(range)
        if title in self.forbidden_titles:
            raise APIError(FakeResponse(403, "The caller does not have permission", "PERMISSION_DENIED"))
        if not any(tab["title"] == title for tab in self.tabs):
            raise APIError(FakeResponse(400, f"Unable to parse range: {range}", "INVALID_ARGUMENT"))
        return {"majorDimension": "ROWS", **super().values_get(id, range, params)}

    def batch_update(self, id, body):
        request = body["requests"][0]["addSheet"]["properties"]
        sheet_id = 90 + len(self.tabs)
        self.calls.append(f"addSheet:{request['title']}")
        self.tabs.append({"title": request["title"], "values": [], "sheetId": sheet_id})
        properties = {"title": request["title"], "sheetId": sheet_id, "index": len(self.tabs) - 1,
                      "gridProperties": request.get("gridProperties", {"rowCount": 100, "columnCount": 2})}
        return {"replies": [{"addSheet": {"properties": properties}}]}

    def values_update(self, id, range, params=None, body=None):
        from tests.test_order_load_sheets import _title
        title, _ = _title(range)
        self.calls.append(f"update:{title}")
        self._tab(title)["values"][:1] = [list(body["values"][0])]
        return {}


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

    def test_renamed_tab_recovers_on_the_failing_call(self):
        config = open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])
        config = open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])  # 캐시로 연 탭
        next(tab for tab in self.http.tabs if tab["title"] == CONFIG)["title"] = "설정(예전)"  # 누가 탭 이름을 바꿈
        self.http.calls.clear()
        # 캐시 없이 열던 때처럼 「설정」 탭을 다시 찾고(없으면 만들고) 같은 호출을 다시 실행한다.
        self.assertEqual(config.get_all_values(), [["키", "값"]])
        self.assertIn("meta", self.http.calls)
        self.assertIn(f"addSheet:{CONFIG}", self.http.calls)
        new_id = next(tab for tab in self.http.tabs if tab["title"] == CONFIG)["sheetId"]
        self.http.calls.clear()
        reopened = open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])
        self.assertEqual(reopened.id, new_id)  # 캐시가 새 탭으로 바뀜
        self.assertEqual(self.http.calls, [])

    def test_other_errors_are_raised_without_reopening(self):
        from gspread.exceptions import APIError
        open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])
        config = open_config_worksheet(self.gc, "sid", CONFIG, ["키", "값"])
        self.http.forbidden_titles = (CONFIG,)
        self.http.calls.clear()
        with self.assertRaises(APIError):
            config.get_all_values()
        self.assertNotIn("meta", self.http.calls)

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
