"""복호화한 주문서 임시 파일이 오류에도 남지 않고, 임시 폴더 정리 실패가 생성 실패로 보이지 않는지 검증한다."""

import contextlib
import io
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pandas
from openpyxl import Workbook
from PySide6.QtWidgets import QMessageBox

from tests.test_excel_text_values import invoice_workbook, load_app_module
from tests.test_naver_api_dispatch import API_ROWS


def encrypted_naver_order(path, rows):
    from msoffcrypto.format.ooxml import OOXMLFile
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "발주발송관리"
    sheet.append(["발주발송관리"])
    sheet.append(["상품주문번호", "주문번호", "배송방법(구매자 요청)", "배송방법", "송장번호",
                  "수취인명", "수취인연락처1", "통합배송지", "우편번호"])
    for row in rows:
        sheet.append(row)
    plain = io.BytesIO()
    workbook.save(plain)
    with open(path, "wb") as output:
        OOXMLFile(io.BytesIO(plain.getvalue())).encrypt("1234", output)


class TempFileCleanupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.cwd = os.getcwd()
        os.chdir(self.directory.name)
        self.dialogs = []
        for name in ("critical", "warning"):
            patcher = patch.object(QMessageBox, name, side_effect=lambda *a, **k: self.dialogs.append(a[2]))
            patcher.start()
            self.addCleanup(patcher.stop)

    def tearDown(self):
        os.chdir(self.cwd)
        self.directory.cleanup()

    def host(self, methods, **attributes):
        main_window = self.app.MainWindow
        Host = type("Host", (), {name: main_window.__dict__[name] for name in methods})
        host = Host()
        host.ui = SimpleNamespace(label_invoice=None)
        host._set_status_label = lambda *a, **k: None
        host._refresh_order_ship_panels = lambda: None
        host._mark_ship_batch_done = lambda *a: None
        host.saved = []
        host._save_invoice_file = lambda frame: host.saved.append(frame.copy())
        host._enqueue_tracking_registration = lambda records: None
        host._oauth_error_dialog_hint = lambda: ""
        for key, value in attributes.items():
            setattr(host, key, value)
        return host

    def test_leftover_file_in_temp_folder_does_not_turn_success_into_failure(self):
        invoice_workbook("1744778498617.xlsx", [
            {"등기번호": "6012345678901", "수취인명": "김", "수취인 이동통신": "010-1111-2222",
             "수취인 주소": "서울 1", "고객 주문번호": "2026111300000001"}])
        Path("temp").mkdir()
        Path("temp/남은파일.txt").write_text("이전 실행에서 남은 파일", encoding="utf-8")
        host = self.host(
            ("generate_invoice_file", "_process_naver_invoice", "_cached_invoice_excel", "_read_invoice_excel",
             "_invoice_excel_from_bytes", "_naver_api_rows_to_order_frame"),
            store_type="naver", is_order_file_valid=True, is_invoice_file_valid=True, selected_file_path=None,
            invoice_file_path=str(Path("1744778498617.xlsx").resolve()),
            _naver_dispatch_api_rows=[dict(row) for row in API_ROWS])
        with contextlib.redirect_stdout(io.StringIO()):
            host.generate_invoice_file()
        self.assertEqual(self.dialogs, [])
        self.assertEqual(len(host.saved), 1)
        self.assertFalse(Path("temp/temp_invoice.xlsx").exists())

    def test_decrypted_order_is_removed_when_dispatch_matching_fails(self):
        Path("temp").mkdir()
        encrypted_naver_order("temp/order.xlsx", [
            ["51234567890100001", "5123456789010001", "택배,등기,소포", "", "", "김", "010-1111-2222", "서울 1", "06035"],
            ["51234567890100002", "5123456789010002", "택배,등기,소포", "", "", "김", "010-1111-2222", "부산 2", "48000"],
        ])
        invoice_workbook("temp/invoice.xlsx", [
            {"등기번호": "6012345678901", "수취인명": "김", "수취인 이동통신": "01011112222",
             "수취인 주소": "대전 3호", "고객 주문번호": "5123456789010001"}])  # 배송지 특정 불가 → 오류
        host = self.host(("_process_naver_invoice", "_naver_api_rows_to_order_frame"), store_type="naver")
        with contextlib.redirect_stdout(io.StringIO()), self.assertRaises(Exception):
            host._process_naver_invoice(Path("temp/order.xlsx"), Path("temp/invoice.xlsx"))
        self.assertFalse(Path("temp/decrypted_order.xlsx").exists())

    def test_decrypted_temp_file_is_removed_when_order_load_fails(self):
        encrypted_naver_order("스마트스토어_전체주문발주발송관리_20261116.xlsx", [
            ["51234567890100001", "5123456789010001", "택배,등기,소포", "", "", "김", "010-1111-2222", "서울 1", "06035"]])
        created = []
        real_named_temp = tempfile.NamedTemporaryFile

        def named_temp(*args, **kwargs):
            handle = real_named_temp(*args, dir=self.directory.name, **kwargs)
            created.append(handle.name)
            return handle

        host = self.host(("process_naver_excel_file",),
                         selected_file_path=str(Path("스마트스토어_전체주문발주발송관리_20261116.xlsx").resolve()),
                         _load_product_code_map_from_spreadsheet=lambda store: {},
                         open_file_with_default_app=lambda path: False)
        with contextlib.redirect_stdout(io.StringIO()), \
                patch.object(tempfile, "NamedTemporaryFile", named_temp), \
                patch.object(pandas, "read_excel", side_effect=ValueError("읽기 실패")), \
                patch.object(QMessageBox, "exec", return_value=QMessageBox.StandardButton.No):
            result = host.process_naver_excel_file()
        self.assertIsNot(result, True)
        self.assertEqual(len(created), 1)
        self.assertFalse(Path(created[0]).exists())


if __name__ == "__main__":
    unittest.main()
