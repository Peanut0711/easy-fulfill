"""엑셀 빈 칸 때문에 실수형이 된 열의 번호가 '6012345678901.0'처럼 바뀌지 않는지 검증한다."""

import contextlib
import importlib.util
import io
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
from openpyxl import Workbook

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS", "--disable-gpu")

from orders.excel_values import excel_text
from orders.invoice import _zipcode
from orders.matching import normalize_coupang_text, _normalized_text

ROOT = Path(__file__).resolve().parents[1]


def invoice_workbook(path, rows):
    workbook = Workbook()
    sheet = workbook.active
    for _ in range(6):
        sheet.append(["우체국 접수내역"])
    header = ["등기번호", "수취인명", "수취인 이동통신", "수취인 전화번호", "수취인 주소", "수취인상세주소", "고객 주문번호"]
    sheet.append(header)
    for row in rows:
        sheet.append([row.get(name) for name in header])
    workbook.save(path)


class ExcelTextTests(unittest.TestCase):
    def test_float_whole_numbers_lose_trailing_zero_and_blanks_are_empty(self):
        self.assertEqual(excel_text(6012345678901.0), "6012345678901")
        self.assertEqual(excel_text(6012345678901), "6012345678901")
        self.assertEqual(excel_text(" 6012345678901 "), "6012345678901")
        self.assertEqual(excel_text(float("nan")), "")
        self.assertEqual(excel_text(None), "")
        self.assertEqual(excel_text(12.5), "12.5")
        self.assertEqual(excel_text("010-1234-5678"), "010-1234-5678")

    def test_blank_cell_in_column_does_not_change_numbers(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "1744778498617.xlsx"
            invoice_workbook(path, [
                {"등기번호": "6012345678901", "수취인명": "김"},
                {"등기번호": None, "수취인명": "합계"},
                {"등기번호": 6012345678902, "수취인명": "이"},
            ])
            frame = pd.read_excel(path, header=6)
        self.assertEqual(frame["등기번호"].dtype.kind, "f")  # 원인: 빈 칸 때문에 실수형
        self.assertEqual([excel_text(v) for v in frame["등기번호"]], ["6012345678901", "", "6012345678902"])

    def test_zipcode_and_matching_normalizers_accept_float_cells(self):
        self.assertEqual(_zipcode(6035.0), "06035")
        self.assertEqual(_zipcode("06035"), "06035")
        self.assertEqual(normalize_coupang_text(1012345678.0), normalize_coupang_text(1012345678))
        self.assertEqual(_normalized_text(12345.0), "12345")


def load_app_module():
    from PySide6.QtWidgets import QApplication
    QApplication.instance() or QApplication([])
    spec = importlib.util.spec_from_file_location("easy_fulfill_for_excel_tests", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class InvoiceProcessingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.cwd = os.getcwd()
        os.chdir(self.root)
        Path("temp").mkdir()

    def tearDown(self):
        os.chdir(self.cwd)
        self.directory.cleanup()

    def host(self, store):
        saved = []
        host = SimpleNamespace(store_type=store, tracking=[], saved=saved, clear_list=lambda: None)
        host._enqueue_tracking_registration = lambda records: host.tracking.extend(records)
        host._save_invoice_file = lambda frame: saved.append(frame.copy())
        return host

    def test_coupang_blank_rows_keep_numbers_as_text(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["번호", "묶음배송번호", "주문번호", "수취인이름", "수취인전화번호", "수취인 주소"])
        sheet.append([1, 30000000001, 27000000000001, "김", "010-1111-2222", "서울 1"])
        sheet.append([2, None, 27000000000002, "이", "010-3333-4444", "부산 2"])
        workbook.save("temp/order.xlsx")
        invoice_workbook("temp/invoice.xlsx", [
            {"등기번호": "6012345678901", "수취인명": "김", "수취인 전화번호": "010-1111-2222"},
            {"등기번호": None, "수취인명": "이", "수취인 전화번호": "010-3333-4444"},
            {"등기번호": "6012345678903", "수취인명": "박", "수취인 전화번호": "010-5555-6666"},
        ])
        host = self.host("coupang")
        with contextlib.redirect_stdout(io.StringIO()):
            self.app.MainWindow._process_coupang_invoice(host, Path("temp/order.xlsx"), Path("temp/invoice.xlsx"))
        frame = host.saved[0]
        self.assertEqual(frame["운송장번호"].tolist(), ["6012345678901", ""])
        self.assertEqual(frame["묶음배송번호"].tolist(), ["30000000001", ""])
        self.assertEqual(frame["주문번호"].tolist(), ["27000000000001", "27000000000002"])
        self.assertEqual([record["등기번호"] for record in host.tracking], ["6012345678901"])

    def test_naver_blank_rows_keep_tracking_numbers_as_text(self):
        from msoffcrypto.format.ooxml import OOXMLFile
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "발주발송관리"
        sheet.append(["발주발송관리"])
        sheet.append(["상품주문번호", "주문번호", "배송방법(구매자 요청)", "배송방법", "송장번호",
                      "수취인명", "수취인연락처1", "통합배송지", "우편번호"])
        sheet.append(["20261113000000011", "2026111300000001", "택배,등기,소포", "", "", "김", "010-1111-2222", "서울 1", "06035"])
        sheet.append(["20261113000000021", "2026111300000002", "택배,등기,소포", "", "", "이", "010-3333-4444", "부산 2", "48000"])
        plain = io.BytesIO()
        workbook.save(plain)
        with open("temp/order.xlsx", "wb") as output:
            OOXMLFile(io.BytesIO(plain.getvalue())).encrypt("1234", output)
        invoice_workbook("temp/invoice.xlsx", [
            {"등기번호": "6012345678901", "수취인명": "김", "수취인 이동통신": "010-1111-2222",
             "수취인 주소": "서울 1", "고객 주문번호": "2026111300000001"},
            {"등기번호": None, "수취인명": "이", "수취인 이동통신": "010-3333-4444",
             "수취인 주소": "부산 2", "고객 주문번호": "2026111300000002"},
        ])
        host = self.host("naver")
        with contextlib.redirect_stdout(io.StringIO()):
            self.app.MainWindow._process_naver_invoice(host, Path("temp/order.xlsx"), Path("temp/invoice.xlsx"))
        frame = host.saved[0]
        self.assertEqual([excel_text(v) for v in frame["송장번호"]], ["6012345678901", ""])
        self.assertEqual([record["trackingNumber"] for record in host._naver_dispatch_records], ["6012345678901"])
        self.assertEqual([record["등기번호"] for record in host.tracking], ["6012345678901"])


if __name__ == "__main__":
    unittest.main()
