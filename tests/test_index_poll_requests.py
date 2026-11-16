"""주문번호 인덱스 주기 폴링이 메타데이터 없이 값 한 번 읽기로 끝나는지 검증한다."""

import contextlib
import io
import unittest

import sheets.google_sheets_oauth as oauth
from tests.test_order_load_sheets import FakeClient, FakeSheetsHTTP, load_app_module, spreadsheet_tabs
from tests.test_order_index_sync import FakeIndexHTTP


class IndexPollRequestTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()
        cls.original = oauth.get_authorized_gspread_client

    @classmethod
    def tearDownClass(cls):
        oauth.get_authorized_gspread_client = cls.original

    def poll(self, http):
        oauth.get_authorized_gspread_client = lambda *a, **k: FakeClient(http)
        with contextlib.redirect_stdout(io.StringIO()):
            return self.app.run_order_index_read_sync_worker()

    def test_poll_reads_values_once_without_metadata(self):
        http = FakeSheetsHTTP(spreadsheet_tabs())
        payload = self.poll(http)
        self.assertEqual(payload, {"ok": True, "kind": "data", "row": {"naver": 12, "coupang": 8, "gmarket": 3}})
        self.assertEqual(http.calls, ["batch:일별 주문번호"])  # 수정 전: 메타데이터 2회 + 값 1회

    def test_missing_today_row_falls_back_and_creates_row(self):
        http = FakeIndexHTTP.make(spreadsheet_tabs(today_row=False))
        payload = self.poll(http)
        self.assertEqual(payload["kind"], "created")
        self.assertIn("append:일별 주문번호", http.calls)


if __name__ == "__main__":
    unittest.main()
