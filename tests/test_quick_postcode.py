"""단건 송장의 주소 검색용 키 등록·공유·오류 처리를 검증한다."""

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

import requests
from PySide6.QtWidgets import QMainWindow, QInputDialog, QMessageBox

import sheets.google_sheets_oauth as oauth
from sheets import worksheet_cache
from tests.test_order_load_sheets import FakeClient, FakeSheetsHTTP, load_app_module


POSTCODE_KEY = "kpost_postcode_regkey"
TRACE_KEY = "kpost_regkey"
POSTCODE_XML = (
    "<response><item><postcd>04524</postcd><address>서울 중구 세종대로 110</address></item></response>"
).encode()
INVALID_KEY_XML = b"<error><error_code>ERR-123</error_code></error>"


def postal_response(content=POSTCODE_XML):
    return SimpleNamespace(raise_for_status=lambda: None, content=content)


class QuickPostcodeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app_module = load_app_module()

    def setUp(self):
        self.rows = [["키", "값"], [TRACE_KEY, "trace-key"]]
        self.http = FakeSheetsHTTP([("설정", self.rows)])
        self.enterContext(patch.object(oauth, "get_authorized_gspread_client", return_value=FakeClient(self.http)))
        worksheet_cache.clear()
        self.addCleanup(worksheet_cache.clear)
        self.folder = Path(self.enterContext(TemporaryDirectory()))
        self.admin = self.make_window("admin.json")
        self.operator = self.make_window("operator.json")

    def make_window(self, filename):
        window = self.app_module.MainWindow.__new__(self.app_module.MainWindow)
        QMainWindow.__init__(window)
        window.app_settings_path = self.folder / filename
        self.addCleanup(window.deleteLater)
        return window

    def test_register_key_and_use_it_from_other_pc_on_next_search(self):
        def search(url, *, params, **kwargs):
            self.assertEqual(url, "https://biz.epost.go.kr/KpostPortal/openapi2")
            self.assertEqual(params["target"], "postNew")
            self.assertFalse(kwargs["allow_redirects"])
            if params["regkey"] == "trace-key":
                return postal_response(INVALID_KEY_XML)
            return postal_response()

        with (
            patch.dict("os.environ", {"KPOST_REGKEY": "trace-key"}),
            patch.object(requests, "get", side_effect=search) as get,
            patch.object(QInputDialog, "getText", return_value=("postcode-key", True)),
            patch.object(QMessageBox, "information") as information,
            patch.object(self.admin, "notify", create=True) as notify,
        ):
            self.assertEqual(self.operator._kpost_postnew_lookup_items("세종대로 110")[1][0], "ERR-123")
            self.admin._on_kpost_postcode_key_edit_clicked()
            # 저장 성공은 확인 창 대신 알림 띠로 알린다(알림·상태바 계획 C-3).
            information.assert_not_called()
            notify.assert_called_once()
            self.assertEqual(notify.call_args.args[0], "ok")
            self.assertEqual(self.rows, [["키", "값"], [TRACE_KEY, "trace-key"], [POSTCODE_KEY, "postcode-key"]])
            self.assertEqual(self.admin._get_kpost_regkey(), "trace-key")
            self.http.calls.clear()
            matches, error = self.operator._kpost_postnew_lookup_items("세종대로 110")
            self.assertIsNone(error)
            self.assertEqual(matches[0]["postcd"], "04524")
            self.assertEqual(get.call_args.kwargs["params"]["regkey"], "postcode-key")
            self.assertFalse(self.operator.app_settings_path.exists())
            self.assertEqual(self.http.calls, ["batch:설정"])

    def test_replacing_key_updates_only_its_cell(self):
        self.rows.append([POSTCODE_KEY, "previous-postcode-key"])

        def update_sheet(_spreadsheet, body):
            self.assertEqual(body["valueInputOption"], "RAW")
            self.assertEqual(body["data"], [{"range": "'설정'!B3", "values": [["replacement-key"]]}])
            self.rows[2][1] = body["data"][0]["values"][0][0]
            return {}

        with (
            patch.object(self.http, "values_batch_update", side_effect=update_sheet),
            patch.object(requests, "get", return_value=postal_response()) as get,
            patch.object(QInputDialog, "getText", return_value=("replacement-key", True)),
            patch.object(QMessageBox, "information") as information,
            patch.object(self.admin, "notify", create=True) as notify,
        ):
            self.admin._on_kpost_postcode_key_edit_clicked()
            information.assert_not_called()
            notify.assert_called_once()
            self.operator._kpost_postnew_lookup_items("세종대로 110")
            self.assertEqual(get.call_args.kwargs["params"]["regkey"], "replacement-key")
        self.assertEqual(self.rows[1:], [[TRACE_KEY, "trace-key"], [POSTCODE_KEY, "replacement-key"]])

    def test_rejected_key_keeps_existing_shared_key(self):
        self.rows.append([POSTCODE_KEY, "previous-postcode-key"])
        with (
            patch.object(requests, "get", return_value=postal_response(INVALID_KEY_XML)),
            patch.object(QInputDialog, "getText", return_value=("invalid-postcode-key", True)),
            patch.object(QMessageBox, "warning") as warning,
            patch.object(QMessageBox, "information") as information,
        ):
            self.admin._on_kpost_postcode_key_edit_clicked()
            warning.assert_called_once()
            self.assertIn("ERR-123", warning.call_args.args[2])
            self.assertNotIn("invalid-postcode-key", warning.call_args.args[2])
            information.assert_not_called()
        self.assertEqual(self.rows[1:], [[TRACE_KEY, "trace-key"], [POSTCODE_KEY, "previous-postcode-key"]])
        self.assertEqual(self.http.calls, ["batch:설정"])

    def test_concurrent_key_change_during_validation_is_preserved(self):
        self.rows.append([POSTCODE_KEY, "previous-postcode-key"])

        def validate(*args, **kwargs):
            self.rows[2][1] = "second-admin-key"
            return postal_response()

        with (
            patch.object(requests, "get", side_effect=validate),
            patch.object(QInputDialog, "getText", return_value=("first-admin-key", True)),
            patch.object(QMessageBox, "warning") as warning,
            patch.object(QMessageBox, "information") as information,
        ):
            self.admin._on_kpost_postcode_key_edit_clicked()
            warning.assert_called_once()
            information.assert_not_called()
        self.assertEqual(self.rows[2][1], "second-admin-key")


if __name__ == "__main__":
    unittest.main()
