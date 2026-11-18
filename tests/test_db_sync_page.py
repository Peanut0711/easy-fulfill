"""「설정 › DB 동기화」 페이지(창 크기·설정 탭 계획 B-5, 4단계 4-6)를 검증한다.

- 파일 표시: 「네이버 DB · 파일명 · 11-17 (오늘)」, 어제·n일 전, 파일 없음
- 테스트 모드를 켜고 끌 때 주의 띠와 실행 버튼 글자
- 실행 버튼 강조, 두 묶음(파일 · 이 PC / 공유 시트 추가 · 전원 공유)
"""

import os
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from sheets.db_sheet_sync import db_sync_file_date, format_db_sync_label_line  # noqa: E402


class FileLabelTests(unittest.TestCase):
    TODAY = date(2026, 11, 18)

    def test_today_yesterday_and_days_ago_from_filename(self):
        cases = {
            "Product_20261118_093000.csv": "11-18 (오늘)",
            "Product_20261117_093000.csv": "11-17 (어제)",
            "스마트스토어상품_20261111_120000.csv": "11-11 (7일 전)",
            "price_inventory_261101.xlsx": "11-01 (17일 전)",
            "price_inventory_20261119.xlsx": "11-19 (미래 날짜)",
        }
        for name, tail in cases.items():
            # 이름에 날짜가 있으면 파일이 없어도(수정 시각을 몰라도) 이름의 날짜를 쓴다.
            text = format_db_sync_label_line("네이버 DB", Path("C:/db") / name, today=self.TODAY)
            self.assertEqual(text, f"네이버 DB · {name} · {tail}")

    def test_missing_file(self):
        self.assertEqual(format_db_sync_label_line("쿠팡 DB", None, today=self.TODAY), "쿠팡 DB · 파일 없음")

    def test_name_without_date_uses_modified_time(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "내보내기.csv"
            path.write_text("a", encoding="utf-8")
            stamp = datetime(2026, 11, 16, 9, 0).timestamp()
            os.utime(path, (stamp, stamp))
            self.assertEqual(db_sync_file_date(path), date(2026, 11, 16))
            self.assertEqual(format_db_sync_label_line("네이버 DB", path, today=self.TODAY),
                             "네이버 DB · 내보내기.csv · 11-16 (2일 전)")
        missing = Path(directory) / "없는파일.csv"
        self.assertEqual(format_db_sync_label_line("네이버 DB", missing, today=self.TODAY),
                         "네이버 DB · 없는파일.csv · 날짜 모름")


class DbSyncPageTests(unittest.TestCase):
    METHODS = ("_update_db_sync_test_mode_ui", "_refresh_db_sheet_sync_path_labels")

    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from tests.test_order_load_sheets import load_app_module
        cls.app = load_app_module()

    def window(self):
        from PySide6.QtWidgets import QCheckBox, QLabel, QPushButton
        cls = type("FakeWindow", (), {name: getattr(self.app.MainWindow, name) for name in self.METHODS})
        window = cls()
        window.ui = SimpleNamespace(
            checkBox_db_sync_preview=QCheckBox(), label_db_sync_test_mode_band=QLabel(),
            pushButton_db_sync_run=QPushButton("실행"),
            label_db_sync_naver_file=QLabel(), label_db_sync_coupang_file=QLabel(),
        )
        window._db_sync_naver_path_override = None
        window._db_sync_coupang_path_override = None
        return window

    def test_test_mode_band_and_button_text(self):
        window = self.window()
        band, button = window.ui.label_db_sync_test_mode_band, window.ui.pushButton_db_sync_run
        window._update_db_sync_test_mode_ui()
        self.assertTrue(band.isHidden())
        self.assertEqual(button.text(), "실행")
        window.ui.checkBox_db_sync_preview.setChecked(True)
        window._update_db_sync_test_mode_ui()
        self.assertFalse(band.isHidden())
        self.assertEqual(band.property("band"), "warn")
        self.assertEqual(button.text(), "테스트 실행")
        window.ui.checkBox_db_sync_preview.setChecked(False)
        window._update_db_sync_test_mode_ui()
        self.assertTrue(band.isHidden())
        self.assertEqual(button.text(), "실행")

    def test_file_labels(self):
        import sheets.db_sheet_sync as sync
        window = self.window()
        found = Path("C:/db/Product_20261117_093000.csv")
        manual = Path("D:/다른곳/price_inventory_261118.xlsx")
        window._db_sync_coupang_path_override = manual
        with patch.object(sync, "get_latest_file_from_patterns", return_value=found), \
                patch.object(sync, "get_latest_file_from_pattern", return_value=None):
            window._refresh_db_sheet_sync_path_labels()
        naver, coupang = window.ui.label_db_sync_naver_file, window.ui.label_db_sync_coupang_file
        self.assertTrue(naver.text().startswith("네이버 DB · Product_20261117_093000.csv · 11-17 ("))
        self.assertEqual(naver.toolTip(), str(found))  # 전체 경로는 툴팁
        self.assertFalse(naver.property("status"))
        self.assertTrue(coupang.text().startswith("쿠팡 DB · price_inventory_261118.xlsx · 11-18 ("))
        self.assertTrue(coupang.text().endswith(" · 직접 지정"))
        with patch.object(sync, "get_latest_file_from_patterns", return_value=None), \
                patch.object(sync, "get_latest_file_from_pattern", return_value=None):
            window._refresh_db_sheet_sync_path_labels(reset_manual_overrides=True)
        self.assertEqual((coupang.text(), coupang.property("status")), ("쿠팡 DB · 파일 없음", "warn"))
        self.assertEqual(naver.toolTip(), "")


class DbSyncUiFileTests(unittest.TestCase):
    def test_two_groups_keep_widget_names(self):
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        files = tree.find(".//widget[@name='groupBox_db_sync_files']")
        run = tree.find(".//widget[@name='groupBox_db_sync_run']")
        self.assertTrue(files.find("./property[@name='title']/string").text.endswith("· 이 PC"))
        self.assertTrue(run.find("./property[@name='title']/string").text.endswith("· 전원 공유"))
        for name in ("checkBox_db_sync_naver", "label_db_sync_naver_file", "pushButton_db_sync_naver_browse",
                     "checkBox_db_sync_coupang", "label_db_sync_coupang_file", "pushButton_db_sync_coupang_browse",
                     "pushButton_db_sync_refresh_paths"):
            self.assertIsNotNone(files.find(f".//widget[@name='{name}']"), name)
        for name in ("checkBox_db_sync_preview", "checkBox_db_sync_verbose_log", "label_db_sync_test_mode_band",
                     "pushButton_db_sync_run", "plainTextEdit_db_sync_log"):
            self.assertIsNotNone(run.find(f".//widget[@name='{name}']"), name)


if __name__ == "__main__":
    unittest.main()
