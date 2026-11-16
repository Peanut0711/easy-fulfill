"""우체국 실제 접수가 재시도로 두 번 접수되지 않는지, API 오류에 인증키가 드러나지 않는지 검증한다."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import requests  # noqa: E402

from shipping import post_parcel  # noqa: E402
from shipping.post_parcel_receipt_store import ParcelReceiptStore  # noqa: E402


class ApiErrorRedactionTests(unittest.TestCase):
    def test_connection_error_hides_key_and_request_data(self):
        error = requests.ConnectionError(
            "HTTPConnectionPool(host='ship.epost.go.kr'): Max retries exceeded with url: "
            "/api.InsertOrder.jparcel?key=SECRETKEY&regData=ENCRYPTEDDATA")
        with patch.object(post_parcel.requests, "get", side_effect=error):
            with self.assertRaises(post_parcel.ParcelNetworkError) as raised:
                post_parcel.call_api("InsertOrder", "a=b", "SECRETKEY", "0123456789abcdef")
        message = str(raised.exception)
        self.assertNotIn("SECRETKEY", message)
        self.assertNotIn("ENCRYPTED", message)
        self.assertIn("InsertOrder", message)
        self.assertIsNone(raised.exception.__cause__)
        self.assertTrue(raised.exception.__suppress_context__)

    def test_http_error_and_postcode_error_hide_keys(self):
        response = Mock(status_code=500)
        response.raise_for_status.side_effect = requests.HTTPError(
            "500 Server Error for url: http://ship.epost.go.kr/api.GetResInfo.jparcel?key=SECRETKEY&regData=X")
        with patch.object(post_parcel.requests, "get", return_value=response):
            with self.assertRaises(post_parcel.ParcelNetworkError) as raised:
                post_parcel.call_api("GetResInfo", "a=b", "SECRETKEY", "0123456789abcdef")
        self.assertNotIn("SECRETKEY", str(raised.exception))
        self.assertIn("HTTP 500", str(raised.exception))
        with patch.object(post_parcel.requests, "get", side_effect=requests.Timeout("url ...?regkey=POSTKEY")):
            with self.assertRaises(post_parcel.ParcelNetworkError) as raised:
                post_parcel.postcode_lookup_items("서울", "POSTKEY")
        self.assertNotIn("POSTKEY", str(raised.exception))


class SubmissionStoreTests(unittest.TestCase):
    def test_submission_lifecycle(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ParcelReceiptStore(Path(directory) / "receipts.sqlite3")
            self.assertIsNone(store.find_open_submission(fingerprint="fp1"))
            store.begin_submission("EFREAL-1", "fp1")
            self.assertEqual(store.find_open_submission(fingerprint="fp1").order_no, "EFREAL-1")
            self.assertEqual(store.find_open_submission(order_no="EFREAL-1").fingerprint, "fp1")
            self.assertIsNone(store.find_open_submission(fingerprint="other"))
            store.complete_submission("EFREAL-1")
            self.assertIsNone(store.find_open_submission(fingerprint="fp1"))

    def test_fingerprint_ignores_generated_order_number_and_spacing(self):
        row = {"수취인명": "김 철수", "수취인 주소": "서울  1", "수취인 전화번호": "010-1", "상품명": "A"}
        self.assertEqual(post_parcel.parcel_row_fingerprint(row),
                         post_parcel.parcel_row_fingerprint({**row, "수취인명": " 김 철수 "}))
        self.assertNotEqual(post_parcel.parcel_row_fingerprint(row),
                            post_parcel.parcel_row_fingerprint({**row, "상품명": "B"}))


def run_scenario(name):
    """실제 접수 흐름(자식 프로세스): 재조회 실패 뒤 다시 실행했을 때 접수 요청 횟수."""
    import importlib.util
    import time
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from openpyxl import Workbook
    from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

    import sheets.google_sheets_oauth as oauth
    from orders.invoice import INVOICE_COLUMNS
    from shipping.post_parcel import ParcelApiError, ParcelReceipt, RecipientAddressSuggestion

    work = Path.cwd()
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(list(INVOICE_COLUMNS))
    row = {"수취인명": "김철수", "우편번호": "06035", "수취인 주소": "서울 강남구 테헤란로 1",
           "수취인 전화번호": "010-1111-2222", "수취인 이동통신": "010-1111-2222", "상품명": "상품"}
    sheet.append([row.get(column, "") for column in INVOICE_COLUMNS])
    workbook.save(work / "invoice.xlsx")

    spec = importlib.util.spec_from_file_location("easy_fulfill_parcel_scenario", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    submits, lookups, dialogs = [], [], []

    def submit(row, config):
        submits.append(row["주문번호"])
        if len(submits) == 1:
            # InsertOrder는 성공했지만 GetResInfo 재조회가 실패한 상황
            raise ParcelApiError("", "소포신청 재조회 결과의 등기번호가 최초 응답과 일치하지 않습니다.")
        return ParcelReceipt(row["주문번호"], "REQ2", "RES2", "6099999999999", True)

    def lookup(order_no, settings, req_ymd):
        lookups.append(order_no)
        if name == "retry_found":
            return ParcelReceipt(order_no, "REQ1", "RES1", "6011111111111", True)
        return None

    def record_dialog(kind):
        def inner(*args, **kwargs):
            dialogs.append([kind, args[1] if len(args) > 1 else "", args[2] if len(args) > 2 else ""])
            return QMessageBox.StandardButton.Yes if kind == "question" else None
        return inner

    store_path = work / "receipts.sqlite3"
    suggestion = RecipientAddressSuggestion("06035", "서울 강남구 테헤란로 1", "-")
    with patch.object(oauth, "get_authorized_gspread_client", lambda: object()), \
            patch.object(app_module, "_standalone_open_config_ws", lambda gc: object()), \
            patch.object(app_module, "_read_config_values_map", lambda ws: {"epost_postcode_api_key": "K"}), \
            patch.object(app_module, "resolve_recipient_address", lambda address, key: [suggestion]), \
            patch.object(app_module, "submit_real_order", submit), \
            patch.object(app_module, "lookup_order_receipt", lookup, create=True), \
            patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(work / "invoice.xlsx"), "")), \
            patch.object(QMessageBox, "question", side_effect=record_dialog("question")), \
            patch.object(QMessageBox, "warning", side_effect=record_dialog("warning")), \
            patch.object(QMessageBox, "critical", side_effect=record_dialog("critical")), \
            patch.object(QMessageBox, "information", side_effect=record_dialog("information")):
        from shipping.post_parcel_receipt_store import ParcelReceiptStore as RealStore
        app_module.ParcelReceiptStore = lambda: RealStore(store_path)
        window = app_module.MainWindow()
        window._run_epost_receipt(test_mode=False)  # 1회차: 재조회 실패
        window._run_epost_receipt(test_mode=False)  # 2회차: 같은 엑셀로 다시 실행
        stored = RealStore(store_path).list_pending_prints()
    result = {"submits": submits, "lookups": lookups, "dialogs": dialogs,
              "stored": [[item.order_no, item.regi_no] for item in stored]}
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


class RealReceiptRetryTests(unittest.TestCase):
    def scenario(self, name):
        with tempfile.TemporaryDirectory() as directory:
            env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
            completed = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), "--scenario", name],
                cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=120,
            )
        line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
        self.assertIsNotNone(line, completed.stdout[-2000:] + completed.stderr[-2000:])
        return json.loads(line[len("RESULT "):])

    def test_retry_finds_previous_receipt_without_submitting_again(self):
        result = self.scenario("retry_found")
        self.assertEqual(len(result["submits"]), 1)
        self.assertEqual(result["lookups"], result["submits"])  # 같은 주문번호로 먼저 조회
        self.assertEqual(result["stored"], [[result["submits"][0], "6011111111111"]])

    def test_retry_not_found_asks_before_submitting_again_with_same_order_no(self):
        result = self.scenario("retry_not_found")
        self.assertEqual(len(result["submits"]), 2)
        self.assertEqual(result["submits"][0], result["submits"][1])
        self.assertEqual(result["lookups"], [result["submits"][0]])
        questions = [text for kind, _title, text in result["dialogs"] if kind == "question"]
        self.assertTrue(any("찾지 못했습니다" in text and result["submits"][0] in text for text in questions))


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--scenario":
        run_scenario(sys.argv[2])
    else:
        unittest.main()
