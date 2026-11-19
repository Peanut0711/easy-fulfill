"""상품 DB 동기화: API로 상품 목록을 받아 공유 시트에 없는 상품만 추가하는 경로를 검증한다.

- 시트 헤더 행 찾기(쿠팡 탭은 1행 헤더로 바뀜)와 같은 헤더 두 번
- 네이버·쿠팡 API 응답 → 시트 헤더 이름 행
- 쿠팡은 시트에 없는 등록 상품만 상세 조회
- 미리보기는 쓰지 않고, 실반영은 쓰기 직전 다시 읽어 다른 PC가 먼저 추가한 상품을 뺀다
- 한 채널이 실패해도 다른 채널은 계속
- 「설정 › DB 동기화」 가져올 곳 선택(API 기본)
"""

import os
import sys
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import sheets.db_sheet_sync as sync  # noqa: E402

NAVER_HEADERS = ["상품코드", "판매자상품코드", "상품명", "대표이미지 URL", "상품번호(스마트스토어)", "채널",
                 "판매상태", "전시상태", "재고수량", "판매가", "할인가", "대분류", "중분류", "소분류", "세분류", "",
                 "상품등록일"]
COUPANG_HEADERS = ["상품코드", "쿠팡 노출 상품명", "업체상품 ID", "Product ID", "옵션 ID", "상품상태",
                   "업체 등록 상품명", "등록 옵션명", "판매가격", "할인율기준가", "판매상태", "승인상태", "판매가격", "메모"]


class FakeWorksheet:
    def __init__(self, values, after_first_read=None):
        self.values = values
        self.reads = 0
        self.after_first_read = after_first_read  # 두 번째 읽기 전에 다른 PC가 쓴 행
        self.appended = []

    def get_all_values(self):
        self.reads += 1
        if self.reads == 2 and self.after_first_read:
            self.values = self.values + self.after_first_read
        return [list(row) for row in self.values]


class FakeSpreadsheet:
    def __init__(self, tabs, config):
        self.tabs = tabs
        self.config = config

    def get_worksheet(self, index):
        return self.tabs[index]

    def worksheet(self, title):
        assert title == "설정"
        return SimpleNamespace(get_all_values=lambda: [["키", "값"]] + [[k, v] for k, v in self.config.items()])


def naver_product(no, name="상품", status="SALE", channel="STOREFARM"):
    return {
        "channelProductNo": no, "channelServiceType": channel, "name": name, "statusType": status,
        "channelProductDisplayStatusType": "ON", "stockQuantity": 20, "salePrice": 28000,
        "discountedPrice": 25000, "wholeCategoryName": "디지털/가전>PC액세서리>기타PC액세서리",
        "representativeImage": {"url": "https://img/1.jpg"}, "regDate": "2026-11-19T12:32:04.102+09:00",
    }


class HeaderAndAlignTests(unittest.TestCase):
    def test_detect_header_row(self):
        self.assertEqual(sync.detect_header_row([COUPANG_HEADERS, ["S1"]], 2), 1)
        self.assertEqual(sync.detect_header_row([["제목"], COUPANG_HEADERS, ["S1"]], 1), 2)
        self.assertEqual(sync.detect_header_row([["a"], ["b"]], 2), 2)  # 못 찾으면 기본값

    def test_records_keep_first_of_duplicate_header(self):
        values = [["상품코드", "판매가격", "판매가격", ""], ["S1", "100", "200", "x"]]
        records = sync.records_from_values(values, 1)
        self.assertEqual(records, [{"상품코드": "S1", "판매가격": "100", "빈컬럼_3": "x"}])

    def test_align_blank_product_code_and_duplicate_header_suffix(self):
        row = sync.align_to_spreadsheet_headers(
            {"상품코드": "X", "판매가격": 100, "판매가격.1": float("nan"), "옵션 ID": 7, "": "무시"},
            ["상품코드", "판매가격", "옵션 ID", "판매가격", ""])
        self.assertEqual(row, ["", 100, 7, "", ""])


class ApiRowTests(unittest.TestCase):
    def test_naver_rows(self):
        rows = sync.naver_rows_from_api([
            naver_product(13883969437, "보드", "OUTOFSTOCK"),
            naver_product(999, "쇼핑윈도", channel="WINDOW"),
            naver_product(None),
        ])
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["상품번호(스마트스토어)"], 13883969437)
        self.assertEqual((row["판매상태"], row["전시상태"], row["채널"]), ("품절", "전시중", "스마트스토어"))
        self.assertEqual((row["대분류"], row["소분류"], row["세분류"]), ("디지털/가전", "기타PC액세서리", ""))
        self.assertEqual((row["판매가"], row["할인가"], row["재고수량"]), (28000, 25000, 20))
        self.assertEqual(row["상품등록일"], "2026-11-19 12:32:04")
        self.assertNotIn("상품코드", row)

    def test_coupang_rows_one_per_option_and_skip_unapproved(self):
        rows = sync.coupang_rows_from_api([{
            "sellerProductId": 16463844485, "productId": 9830338490, "statusName": "승인완료",
            "sellerProductName": "ESP32 보드", "displayProductName": "Waveshare ESP32 보드",
            "items": [
                {"vendorItemId": 96353340251, "itemName": "1개", "salePrice": 37000, "originalPrice": 40000,
                 "offerCondition": "NEW", "barcode": "", "externalVendorSku": "SKU1"},
                {"vendorItemId": None, "itemName": "승인 전"},
            ],
        }])
        self.assertEqual(rows, [{
            "쿠팡 노출 상품명": "Waveshare ESP32 보드, 1개", "업체상품 ID": 16463844485, "Product ID": 9830338490,
            "옵션 ID": 96353340251, "상품상태": "새상품", "바코드": "", "업체상품코드": "SKU1",
            "업체 등록 상품명": "ESP32 보드", "등록 옵션명": "1개", "판매가격": 37000, "할인율기준가": 40000,
            "승인상태": "승인완료",
        }])

    def test_coupang_details_only_for_products_missing_in_sheet(self):
        from integrations import coupang_commerce
        listed = [{"sellerProductId": 1}, {"sellerProductId": 2}, {"sellerProductId": 2}]
        detail = {"sellerProductId": 2, "productId": 20, "items": [{"vendorItemId": 200, "itemName": "옵션"}]}
        config = {"coupang_vendor_id": "V", "coupang_access_key": "A", "coupang_secret_key": "S"}
        with patch.object(coupang_commerce, "fetch_seller_products", return_value=listed), \
                patch.object(coupang_commerce, "fetch_seller_product", return_value=detail) as fetch_detail:
            rows = sync.load_coupang_products_from_api(config, [{"업체상품 ID": "1"}], lambda _m: None)
        fetch_detail.assert_called_once_with("2", "A", "S")
        self.assertEqual([row["옵션 ID"] for row in rows], [200])

    def test_missing_keys_message(self):
        with self.assertRaisesRegex(RuntimeError, "네이버 API 키가 등록되지 않았습니다"):
            sync.load_naver_products_from_api({"naver_client_id": "id"}, [], lambda _m: None)


class SyncChannelTests(unittest.TestCase):
    def run_naver(self, worksheet, products, *, test_mode):
        spreadsheet = FakeSpreadsheet({0: worksheet}, {"naver_client_id": "id", "naver_client_secret": "secret"})
        gc = SimpleNamespace(open_by_key=lambda _id: spreadsheet)
        appended = []

        def fake_append(ws, rows, header_row, **_kwargs):
            appended.append((rows, header_row))

        from integrations import naver_commerce
        logs = []
        with patch.object(sync, "get_authorized_gspread_client", return_value=gc), \
                patch.object(naver_commerce, "get_access_token", return_value="token"), \
                patch.object(naver_commerce, "fetch_channel_products", return_value=products), \
                patch.object(sync, "append_to_spreadsheet", side_effect=fake_append):
            result = sync.sync_channel("naver", "sheet", source=sync.SOURCE_API, test_mode=test_mode,
                                       log=logs.append, verbose_log=False)
        return result, appended, logs

    def test_preview_does_not_write(self):
        worksheet = FakeWorksheet([NAVER_HEADERS, ["A-1", "", "기존", "", "100"]])
        result, appended, logs = self.run_naver(worksheet, [naver_product(100), naver_product(200, "새 상품")],
                                                test_mode=True)
        self.assertEqual((result["new_products_count"], result["appended"]), (1, 0))
        self.assertEqual(appended, [])
        self.assertTrue(any("새 상품" in line for line in logs))  # 상세 로그를 꺼도 신규 목록은 보인다

    def test_apply_rereads_and_skips_rows_added_by_other_pc(self):
        other_pc_row = ["", "", "다른 PC", "", "300"]
        worksheet = FakeWorksheet([NAVER_HEADERS, ["A-1", "", "기존", "", "100"]], after_first_read=[other_pc_row])
        result, appended, _logs = self.run_naver(
            worksheet, [naver_product(100), naver_product(200, "새 상품"), naver_product(300, "다른 PC")],
            test_mode=False)
        self.assertEqual((result["new_products_count"], result["skipped_concurrent"], result["appended"]), (2, 1, 1))
        rows, header_row = appended[0]
        self.assertEqual(header_row, 1)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], "")  # 상품코드는 작업자가 채운다
        self.assertEqual(rows[0][NAVER_HEADERS.index("상품번호(스마트스토어)")], 200)
        self.assertEqual(rows[0][NAVER_HEADERS.index("상품명")], "새 상품")

    def test_job_continues_after_channel_failure(self):
        def fake_sync(channel, *_args, **_kwargs):
            if channel == "naver":
                raise RuntimeError("네이버 API 키가 등록되지 않았습니다.")
            result = sync._empty_channel_result(channel)
            result["ok"] = True
            return result

        with patch.object(sync, "sync_channel", side_effect=fake_sync) as called:
            out = sync.run_db_sheet_sync_job("sheet", do_naver=True, do_coupang=True, test_mode=True,
                                             test_count=0, verbose_log=False, source=sync.SOURCE_API)
        self.assertEqual(called.call_count, 2)
        self.assertFalse(out["ok"])
        self.assertTrue(out["coupang"]["ok"])
        self.assertIn("네이버 API 키", out["error"])
        self.assertEqual(called.call_args.kwargs["source"], sync.SOURCE_API)
        self.assertIsNone(called.call_args.kwargs["realtime_file_path"])


class SourceSelectionUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from tests.test_order_load_sheets import load_app_module
        cls.app = load_app_module()

    def test_radio_buttons_in_ui_file_api_default(self):
        tree = ET.parse(ROOT / "ui" / "main_window.ui")
        files = tree.find(".//widget[@name='groupBox_db_sync_files']")
        api = files.find(".//widget[@name='radioButton_db_sync_source_api']")
        self.assertIsNotNone(files.find(".//widget[@name='radioButton_db_sync_source_file']"))
        self.assertEqual(api.find("./property[@name='checked']/bool").text, "true")

    def test_file_controls_follow_source(self):
        from PySide6.QtWidgets import QButtonGroup, QLabel, QPushButton, QRadioButton
        names = ("_db_sync_source", "_update_db_sync_source_ui", "_set_db_sync_path_controls_busy")
        cls = type("FakeWindow", (), {name: getattr(self.app.MainWindow, name) for name in names})
        window = cls()
        window._db_sheet_sync_thread = None
        window.ui = SimpleNamespace(
            radioButton_db_sync_source_api=QRadioButton(), radioButton_db_sync_source_file=QRadioButton(),
            label_db_sync_naver_file=QLabel(), label_db_sync_coupang_file=QLabel(),
            pushButton_db_sync_naver_browse=QPushButton(), pushButton_db_sync_coupang_browse=QPushButton(),
            pushButton_db_sync_refresh_paths=QPushButton(), pushButton_db_sync_run=QPushButton(),
        )
        group = QButtonGroup()
        group.addButton(window.ui.radioButton_db_sync_source_api)
        group.addButton(window.ui.radioButton_db_sync_source_file)
        window.ui.radioButton_db_sync_source_api.setChecked(True)
        window._update_db_sync_source_ui()
        self.assertEqual(window._db_sync_source(), sync.SOURCE_API)
        self.assertFalse(window.ui.pushButton_db_sync_naver_browse.isEnabled())
        window.ui.radioButton_db_sync_source_file.setChecked(True)
        window._update_db_sync_source_ui()
        self.assertEqual(window._db_sync_source(), sync.SOURCE_FILE)
        self.assertTrue(window.ui.pushButton_db_sync_naver_browse.isEnabled())
        window._set_db_sync_path_controls_busy(True)
        self.assertFalse(window.ui.pushButton_db_sync_refresh_paths.isEnabled())
        self.assertFalse(window.ui.radioButton_db_sync_source_api.isEnabled())
        window._db_sheet_sync_thread = SimpleNamespace(isRunning=lambda: True)  # 끝 신호 때 아직 실행 중
        window._set_db_sync_path_controls_busy(False)
        self.assertTrue(window.ui.pushButton_db_sync_refresh_paths.isEnabled())
        self.assertTrue(window.ui.pushButton_db_sync_run.isEnabled())


if __name__ == "__main__":
    unittest.main()
