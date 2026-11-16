"""송장 미리보기에서 읽은 엑셀을 발송 결과 생성에 재사용하는 조건을 검증한다."""

import ast
import hashlib
import os
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl import Workbook


# MainWindow 전체 초기화 없이 송장 엑셀 캐시 메서드만 불러온다.
source_path = Path(__file__).resolve().parents[1] / "easy-fulfill.py"
tree = ast.parse(source_path.read_text(encoding="utf-8"))
main_window = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "MainWindow")
methods = [
    node for node in main_window.body
    if isinstance(node, ast.FunctionDef)
    and node.name in {"_invoice_excel_from_bytes", "_read_invoice_excel", "_cached_invoice_excel"}
]
host_class = ast.ClassDef(
    name="InvoiceExcelHost", bases=[], keywords=[], body=methods, decorator_list=[], type_params=[],
)
namespace = {"pd": pd, "BytesIO": BytesIO, "Path": Path, "hashlib": hashlib, "os": os}
exec(compile(ast.fix_missing_locations(ast.Module(body=[host_class], type_ignores=[])), str(source_path), "exec"),
     namespace)
InvoiceExcelHost = namespace["InvoiceExcelHost"]


def write_invoice(path, numbers):
    workbook = Workbook()
    sheet = workbook.active
    for _ in range(6):
        sheet.append(["우체국 접수내역"])
    sheet.append(["등기번호", "수취인명"])
    for number in numbers:
        sheet.append([number, "홍길동"])
    workbook.save(path)


class InvoiceExcelReuseTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name) / "1744778498617.xlsx"
        write_invoice(self.path, ["6000000000001", "6000000000002"])
        self.host = InvoiceExcelHost()

    def tearDown(self):
        self.directory.cleanup()

    def test_same_file_reuses_preview_dataframe(self):
        preview = self.host._read_invoice_excel(str(self.path))
        reused = self.host._cached_invoice_excel(str(self.path))
        pd.testing.assert_frame_equal(reused, pd.read_excel(self.path, header=6))
        pd.testing.assert_frame_equal(reused, preview)
        # 재사용본을 바꿔도 보관한 미리보기 데이터는 그대로다.
        reused.loc[0, "수취인명"] = "변경"
        self.assertEqual(self.host._cached_invoice_excel(str(self.path)).loc[0, "수취인명"], "홍길동")

    def test_replaced_or_modified_file_is_read_again(self):
        self.host._read_invoice_excel(str(self.path))
        stat = self.path.stat()
        write_invoice(self.path, ["6000000000001", "6000000000009"])
        os.utime(self.path, ns=(stat.st_atime_ns, stat.st_mtime_ns))  # 수정 시각을 보존한 교체
        self.assertIsNone(self.host._cached_invoice_excel(str(self.path)))

    def test_other_file_or_missing_cache_is_not_reused(self):
        self.assertIsNone(self.host._cached_invoice_excel(str(self.path)))
        self.host._read_invoice_excel(str(self.path))
        other = Path(self.directory.name) / "1744778498618.xlsx"
        write_invoice(other, ["6000000000001", "6000000000002"])
        self.assertIsNone(self.host._cached_invoice_excel(str(other)))
        self.path.unlink()
        self.assertIsNone(self.host._cached_invoice_excel(str(self.path)))


if __name__ == "__main__":
    unittest.main()
