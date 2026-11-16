"""공유 시트의 최소 버전보다 낮은 앱은 시트에 쓰지 못하고, 시작 시 PC별 버전이 기록되는지 검증한다."""

import importlib
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import sheets.google_sheets_oauth as oauth


def gate():
    return importlib.import_module("sheets.version_gate")


class RecordingHTTP:
    def __init__(self):
        self.calls = []

    def request(self, method, endpoint, *args, **kwargs):
        self.calls.append((method, endpoint))
        return "response"


class FakeClient:
    def __init__(self):
        self.http_client = RecordingHTTP()


class FakeConfigSheet:
    def __init__(self, values):
        self.rows = [["키", "값"]] + [[k, v] for k, v in values.items()]
        self.writes = 0

    def get_all_values(self):
        return [list(row) for row in self.rows]

    def append_rows(self, rows, value_input_option=None):
        self.writes += 1
        self.rows.extend([list(row) for row in rows])

    def batch_update(self, updates, value_input_option=None):
        self.writes += 1
        for update in updates:
            self.rows[int(update["range"][1:]) - 1][1] = update["values"][0][0]

    def value(self, key):
        return next((row[1] for row in self.rows[1:] if row[0] == key), "")


class FakePcSheet:
    def __init__(self, headers):
        self.rows = [list(headers)]
        self.writes = 0

    def get_all_values(self):
        return [list(row) for row in self.rows]

    def append_rows(self, rows, value_input_option=None):
        self.writes += 1
        self.rows.extend([list(row) for row in rows])

    def update(self, values, range_name=None, value_input_option=None):
        self.writes += 1
        row = int(range_name.split(":")[0][1:]) - 1
        self.rows[row][1:3] = values[0]


class WriteGuardTests(unittest.TestCase):
    def setUp(self):
        self.gate = gate()
        self.gate.reset_state()
        self.addCleanup(self.gate.reset_state)
        self.min_text = [""]
        patcher = patch.object(self.gate, "_read_min_version_text", lambda gc: self.min_text[0])
        patcher.start()
        self.addCleanup(patcher.stop)

    def guarded(self):
        client = FakeClient()
        self.gate.install_write_guard(client)
        return client

    def test_writes_blocked_when_shared_minimum_is_newer(self):
        self.min_text[0] = "2999.1.1.1"
        client = self.guarded()
        self.assertEqual(client.http_client.request("get", "values"), "response")  # 읽기는 허용
        for method in ("put", "post", "PUT"):
            with self.assertRaises(self.gate.AppVersionBlockedError) as raised:
                client.http_client.request(method, "values:batchUpdate")
            self.assertIn("2999.1.1.1", str(raised.exception))
        self.assertEqual(client.http_client.calls, [("get", "values")])  # 쓰기 요청은 나가지 않음

    def test_writes_allowed_when_app_is_current(self):
        self.min_text[0] = self.gate.APP_VERSION
        client = self.guarded()
        self.assertEqual(client.http_client.request("post", "values:append"), "response")

    def test_minimum_is_rechecked_and_other_pc_raise_takes_effect(self):
        now = [100.0]
        with patch.object(self.gate, "_monotonic", lambda: now[0]):
            client = self.guarded()
            client.http_client.request("put", "a")  # 최소 버전 없음 → 허용
            self.min_text[0] = "2999.1.1.1"  # 다른(새) PC가 최소 버전을 올림
            client.http_client.request("put", "b")  # 다시 확인 전에는 허용
            now[0] += self.gate.MIN_VERSION_RECHECK_SECONDS + 1
            with self.assertRaises(self.gate.AppVersionBlockedError):
                client.http_client.request("put", "c")

    def test_read_failure_does_not_block_writes(self):
        def fail(gc):
            raise RuntimeError("네트워크 오류")

        with patch.object(self.gate, "_read_min_version_text", fail):
            client = self.guarded()
            self.assertEqual(client.http_client.request("put", "a"), "response")

    def test_version_compare(self):
        parse = self.gate.parse_version
        self.assertLess(parse("2026.11.17.1"), parse("2026.11.17.2"))
        self.assertLess(parse("2026.9.30.5"), parse("2026.11.1.1"))
        self.assertIsNone(parse("abc"))
        self.assertIsNone(parse(""))


class AuthorizedClientTests(unittest.TestCase):
    def test_authorized_client_has_write_guard(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        token = Path(directory.name) / "token.json"
        token.write_text("{}", encoding="utf-8")
        creds = type("Creds", (), {"valid": True})()
        client = FakeClient()
        original_request = client.http_client.request
        with patch.object(oauth, "TOKEN_PATH", token), \
                patch.object(oauth, "GOOGLE_AUTH_DIR", Path(directory.name)), \
                patch.object(oauth.Credentials, "from_authorized_user_file", lambda *a, **k: creds), \
                patch("gspread.authorize", lambda c: client):
            returned = oauth.get_authorized_gspread_client(interactive=False)
        self.assertIs(returned, client)
        self.assertNotEqual(returned.http_client.request, original_request)  # 쓰기 차단 장치가 감쌈
        self.assertTrue(getattr(returned.http_client, "_easy_fulfill_write_guard", False))


class StartupCheckTests(unittest.TestCase):
    def setUp(self):
        self.gate = gate()
        self.gate.reset_state()
        self.addCleanup(self.gate.reset_state)

    def run_check(self, config, pcs, when):
        with patch.object(self.gate, "_open_config_ws", lambda gc: config), \
                patch.object(self.gate, "_open_pc_version_ws", lambda gc: pcs):
            return self.gate.startup_check(object(), "직원@PC1", now=when)

    def test_raises_minimum_and_records_pc_version(self):
        config = FakeConfigSheet({"slack_webhook_url": "x"})
        pcs = FakePcSheet(self.gate.PC_VERSION_HEADERS)
        first = self.run_check(config, pcs, datetime(2026, 11, 17, 9, 0, 0))
        self.assertFalse(first["blocked"], first)
        self.assertEqual(config.value(self.gate.CONFIG_KEY_MIN_APP_VERSION), self.gate.MIN_COMPATIBLE_VERSION)
        self.assertEqual(pcs.rows[1], ["직원@PC1", self.gate.APP_VERSION, "2026-11-17 09:00:00"])
        self.run_check(config, pcs, datetime(2026, 11, 17, 13, 0, 0))
        self.assertEqual(len(pcs.rows), 2)  # 같은 PC는 행을 갱신
        self.assertEqual(pcs.rows[1][2], "2026-11-17 13:00:00")

    def test_does_not_lower_a_higher_minimum(self):
        config = FakeConfigSheet({self.gate.CONFIG_KEY_MIN_APP_VERSION: self.gate.APP_VERSION})
        with patch.object(self.gate, "MIN_COMPATIBLE_VERSION", "2000.1.1.1"):
            self.run_check(config, FakePcSheet(self.gate.PC_VERSION_HEADERS), datetime(2026, 11, 17))
        self.assertEqual(config.value(self.gate.CONFIG_KEY_MIN_APP_VERSION), self.gate.APP_VERSION)

    def test_outdated_app_reports_blocked_and_writes_nothing(self):
        config = FakeConfigSheet({self.gate.CONFIG_KEY_MIN_APP_VERSION: "2999.1.1.1"})
        pcs = FakePcSheet(self.gate.PC_VERSION_HEADERS)
        result = self.run_check(config, pcs, datetime(2026, 11, 17))
        self.assertTrue(result["blocked"])
        self.assertEqual(result["minimum"], "2999.1.1.1")
        self.assertIn("업데이트", self.gate.blocked_message(result["minimum"]))
        self.assertEqual((config.writes, pcs.writes), (0, 0))


class StartupNoticeTests(unittest.TestCase):
    def test_outdated_result_shows_update_notice(self):
        from tests.test_order_load_sheets import load_app_module
        app = load_app_module()
        shown = []
        with patch.object(app.QMessageBox, "warning", lambda parent, title, text: shown.append((title, text))):
            app.MainWindow._on_app_version_checked(None, {"ok": True, "blocked": True, "minimum": "2999.1.1.1"})
            app.MainWindow._on_app_version_checked(None, {"ok": True, "blocked": False, "minimum": "1",
                                                          "version": "2"})
        self.assertEqual(len(shown), 1)
        self.assertEqual(shown[0][0], "프로그램 업데이트 필요")
        self.assertIn("2999.1.1.1", shown[0][1])


if __name__ == "__main__":
    unittest.main()
