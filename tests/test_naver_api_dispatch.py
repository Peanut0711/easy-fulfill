"""네이버 API로 불러온 주문으로도 일괄발송 파일을 만들고, 이전 엑셀로 생성하지 않는지 검증한다."""

import contextlib
import io
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook
from PySide6.QtWidgets import QMessageBox

from tests.test_excel_text_values import invoice_workbook, load_app_module


API_ROWS = [
    {"주문번호": "2026111300000001", "수취인명": "김", "수취인연락처1": "010-1111-2222", "통합배송지": "서울 1",
     "우편번호": "06035", "상품명": "상품A", "수량": 1, "배송방법(구매자 요청)": "택배,등기,소포",
     "_상품주문번호": "2026111300000001001", "_상태": "PAYED"},
    {"주문번호": "2026111300000002", "수취인명": "이", "수취인연락처1": "010-3333-4444", "통합배송지": "부산 2",
     "우편번호": "48000", "상품명": "상품B", "수량": 1, "배송방법(구매자 요청)": "방문수령",
     "_상품주문번호": "2026111300000002001", "_상태": "PAYED"},
]


class NaverApiDispatchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.cwd = os.getcwd()
        os.chdir(self.directory.name)
        invoice_workbook("1744778498617.xlsx", [
            {"등기번호": "6012345678901", "수취인명": "김", "수취인 이동통신": "010-1111-2222",
             "수취인 주소": "서울 1", "고객 주문번호": "2026111300000001"},
        ])
        self.dialogs = []
        for name in ("critical", "warning"):
            patcher = patch.object(QMessageBox, name, side_effect=lambda *a, **k: self.dialogs.append(a[2]))
            patcher.start()
            self.addCleanup(patcher.stop)

    def tearDown(self):
        os.chdir(self.cwd)
        self.directory.cleanup()

    def host(self, selected_file_path):
        main_window = self.app.MainWindow
        methods = ("generate_invoice_file", "_process_naver_invoice", "_cached_invoice_excel",
                   "_read_invoice_excel", "_invoice_excel_from_bytes", "_naver_api_rows_to_order_frame")
        Host = type("Host", (), {name: main_window.__dict__[name] for name in methods})
        host = Host()
        host.store_type = "naver"
        host.is_order_file_valid = True
        host.is_invoice_file_valid = True
        host.selected_file_path = selected_file_path
        host.invoice_file_path = str(Path("1744778498617.xlsx").resolve())
        host._naver_dispatch_api_rows = [dict(row) for row in API_ROWS]
        host.ui = SimpleNamespace(label_invoice=None)
        host._set_status_label = lambda *a, **k: None
        host._refresh_order_ship_panels = lambda: None
        host._mark_ship_batch_done = lambda *a: None
        host.saved = []
        host.tracking = []
        host._save_invoice_file = lambda frame: host.saved.append(frame.copy())
        host._enqueue_tracking_registration = lambda records: host.tracking.extend(records)
        return host

    def generate(self, host):
        with contextlib.redirect_stdout(io.StringIO()):
            host.generate_invoice_file()

    def assert_built_from_api_rows(self, host):
        self.assertEqual(self.dialogs, [])
        frame = host.saved[0]
        self.assertEqual(frame["상품주문번호"].tolist(), ["2026111300000001001"])  # 방문수령 행 제외
        self.assertEqual(frame["송장번호"].tolist(), ["6012345678901"])
        self.assertEqual(frame["배송방법"].tolist(), ["택배,등기,소포"])
        self.assertNotIn("_상태", frame.columns)
        self.assertEqual(host._naver_dispatch_records,
                         [{"orderId": "2026111300000001", "trackingNumber": "6012345678901", "수취인명": "김"}])

    def test_api_orders_generate_without_order_excel(self):
        host = self.host(selected_file_path=None)
        self.generate(host)
        self.assert_built_from_api_rows(host)

    def test_api_orders_ignore_previously_selected_excel(self):
        workbook = Workbook()
        workbook.active.append(["이전 주문서"])
        workbook.save("old_order.xlsx")
        host = self.host(selected_file_path=str(Path("old_order.xlsx").resolve()))
        self.generate(host)
        self.assert_built_from_api_rows(host)


if __name__ == "__main__":
    unittest.main()
