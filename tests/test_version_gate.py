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
        self.enterContext(patch.object(self.gate, "APP_VERSION", "0.0.1"))
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
        self.min_text[0] = "2026.11.17.1"
        client = self.guarded()
        self.assertEqual(client.http_client.request("post", "values:append"), "response")

    def test_encoded_minimum_blocks_outdated_app_and_shows_display_version(self):
        self.min_text[0] = "2026.11.18.1.0.0.2"
        client = self.guarded()
        with self.assertRaises(self.gate.AppVersionBlockedError) as raised:
            client.http_client.request("put", "values")
        message = str(raised.exception)
        self.assertIn("최소 버전 0.0.2", message)
        self.assertLess(message.index("페치(Fetch)"), message.index("풀(Pull)"))
        self.assertNotIn(self.min_text[0], message)
        self.assertEqual(client.http_client.calls, [])

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
        self.enterContext(patch.object(self.gate, "APP_VERSION", "0.0.1"))

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
        key = self.gate.CONFIG_KEY_MIN_APP_VERSION
        config = FakeConfigSheet({key: "2026.11.17.12"})
        self.run_check(config, FakePcSheet(self.gate.PC_VERSION_HEADERS), datetime(2026, 11, 18))
        self.assertEqual(config.value(key), "2026.11.17.12")
        self.assertEqual(config.writes, 0)

    def test_raised_minimum_blocks_both_date_and_older_three_part_releases(self):
        key = self.gate.CONFIG_KEY_MIN_APP_VERSION
        config = FakeConfigSheet({key: "2026.11.17.1"})
        pcs = FakePcSheet(self.gate.PC_VERSION_HEADERS)
        when = datetime(2026, 11, 18)
        with patch.object(self.gate, "APP_VERSION", "0.0.2"), \
                patch.object(self.gate, "MIN_COMPATIBLE_VERSION", "0.0.2"):
            result = self.run_check(config, pcs, when)
        self.assertFalse(result["blocked"])
        self.assertEqual(result["minimum"], "0.0.2")
        self.assertGreater(self.gate.parse_version(config.value(key)), (2026, 11, 17, 12))
        blocked = self.run_check(config, pcs, when)
        self.assertTrue(blocked["blocked"])
        self.assertEqual(blocked["minimum"], "0.0.2")
        self.assertEqual(pcs.writes, 1)

    def test_minimum_changed_during_startup_is_checked_before_pc_write(self):
        key = self.gate.CONFIG_KEY_MIN_APP_VERSION
        config = FakeConfigSheet({})
        pcs = FakePcSheet(self.gate.PC_VERSION_HEADERS)

        def raise_minimum_first(*args, **kwargs):
            from tracking.repository import ConfigValueChangedError
            config.rows.append([key, "2026.11.18.1.0.0.2"])
            raise ConfigValueChangedError(key)

        with patch("tracking.repository.write_config_values", side_effect=raise_minimum_first):
            result = self.run_check(config, pcs, datetime(2026, 11, 18))
        self.assertTrue(result["blocked"])
        self.assertEqual(result["minimum"], "0.0.2")
        self.assertEqual(pcs.writes, 0)

    def test_outdated_app_reports_blocked_and_writes_nothing(self):
        config = FakeConfigSheet({self.gate.CONFIG_KEY_MIN_APP_VERSION: "2999.1.1.1"})
        pcs = FakePcSheet(self.gate.PC_VERSION_HEADERS)
        result = self.run_check(config, pcs, datetime(2026, 11, 17))
        self.assertTrue(result["blocked"])
        self.assertEqual(result["minimum"], "2999.1.1.1")
        self.assertIn("업데이트", self.gate.blocked_message(result["minimum"]))
        self.assertEqual((config.writes, pcs.writes), (0, 0))


    def test_reports_newer_version_used_by_other_pc(self):
        config = FakeConfigSheet({self.gate.CONFIG_KEY_MIN_APP_VERSION: "2026.11.17.1"})
        pcs = FakePcSheet(self.gate.PC_VERSION_HEADERS)
        pcs.rows += [["직원@PC2", "0.0.9", ""],
                     ["직원@PC3", "0.0.10", ""],
                     ["직원@PC4", "2026.11.17.12", ""],
                     ["직원@PC5", "엉뚱한값", ""]]
        when = datetime(2026, 11, 18)
        result = self.run_check(config, pcs, when)
        self.assertEqual(result["latest"], "0.0.10")
        pcs.rows[1][1] = pcs.rows[2][1] = "0.0.1"
        self.assertEqual(self.run_check(config, pcs, when)["latest"], "")

    def test_own_older_row_is_not_reported_as_newer(self):
        pcs = FakePcSheet(self.gate.PC_VERSION_HEADERS)
        pcs.rows.append(["직원@PC1", "0.0.2", ""])  # 이 PC의 예전 기록(다른 폴더에서 실행 등)
        result = self.run_check(FakeConfigSheet({}), pcs, datetime(2026, 11, 21))
        self.assertEqual(result["latest"], "")

    def test_update_notice_due(self):
        due = self.gate.update_notice_due
        self.assertTrue(due("2026.11.20.1", ignored=False, snoozed_day="", today="2026-11-21"))
        self.assertFalse(due("", ignored=False, snoozed_day="", today="2026-11-21"))
        self.assertFalse(due("2026.11.20.1", ignored=True, snoozed_day="", today="2026-11-21"))
        self.assertFalse(due("2026.11.20.1", ignored=False, snoozed_day="2026-11-21", today="2026-11-21"))
        self.assertTrue(due("2026.11.20.1", ignored=False, snoozed_day="2026-11-20", today="2026-11-21"))


class FakeLabel:
    def __init__(self):
        self.text = ""

    def setText(self, text):
        self.text = text


class StartupNoticeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from tests.test_order_load_sheets import load_app_module
        cls.app = load_app_module()

    def window(self, settings=None, answer="close"):
        methods = ("_on_app_version_checked", "_update_app_version_label", "_maybe_show_update_notice",
                   "_on_ignore_version_update_toggled")
        cls = type("FakeWindow", (), {name: getattr(self.app.MainWindow, name) for name in methods})
        window = cls()
        window.settings = dict(settings or {})
        window.get_app_setting = lambda key, default=None: window.settings.get(key, default)
        window.set_app_setting = lambda key, value: window.settings.__setitem__(key, value)
        window.asked = []
        window._ask_update_notice = lambda latest: window.asked.append(latest) or answer
        window.ui = type("Ui", (), {})()
        window.ui.label_app_version = FakeLabel()
        return window

    def newer(self):
        return {"ok": True, "blocked": False, "minimum": "2026.11.17.1", "version": "0.0.1",
                "latest": "0.0.2"}

    def today(self, text):
        return patch("sheets.version_gate.today_text", lambda: text)

    def test_newer_version_notice_once_per_session_when_closed(self):
        window = self.window(answer="close")
        with self.today("2026-11-21"):
            window._on_app_version_checked(self.newer())
            window._on_app_version_checked(self.newer())  # 1시간 뒤 다시 확인
        self.assertEqual(window.asked, ["0.0.2"])
        self.assertIn("0.0.2", window.ui.label_app_version.text)
        with self.today("2026-11-21"):
            restarted = self.window(window.settings)
            restarted._on_app_version_checked(self.newer())  # 다시 실행하면 다시 안내
        self.assertEqual(restarted.asked, ["0.0.2"])

    def test_snooze_today_survives_restart_until_tomorrow(self):
        window = self.window(answer="snooze")
        with self.today("2026-11-21"):
            window._on_app_version_checked(self.newer())
            restarted = self.window(window.settings)
            restarted._on_app_version_checked(self.newer())
        self.assertEqual(window.asked, ["0.0.2"])
        self.assertEqual(restarted.asked, [])  # 오늘은 다시 묻지 않음
        with self.today("2026-11-22"):
            tomorrow = self.window(window.settings)
            tomorrow._on_app_version_checked(self.newer())
        self.assertEqual(tomorrow.asked, ["0.0.2"])

    def test_ignore_option_suppresses_notice(self):
        window = self.window()
        window._on_ignore_version_update_toggled(True)
        with self.today("2026-11-21"):
            window._on_app_version_checked(self.newer())
        self.assertEqual(window.asked, [])
        self.assertIn("0.0.2", window.ui.label_app_version.text)
        window._on_ignore_version_update_toggled(False)
        with self.today("2026-11-21"):
            window._on_app_version_checked(self.newer())
        self.assertEqual(window.asked, ["0.0.2"])

    def test_blocked_warning_shows_once_even_when_ignoring(self):
        window = self.window({"ignore_version_update_notice": True})
        shown = []
        window._show_required_update_notice = shown.append
        blocked = {"ok": True, "blocked": True, "minimum": "0.0.2", "version": "0.0.1"}
        window._on_app_version_checked(blocked)
        window._on_app_version_checked(blocked)
        self.assertEqual(shown, ["0.0.2"])

    def test_settings_tab_has_ignore_checkbox_off_by_default(self):
        import xml.etree.ElementTree as ET
        from pathlib import Path
        tree = ET.parse(Path(__file__).resolve().parents[1] / "ui" / "main_window.ui")
        box = tree.find(".//widget[@class='QCheckBox'][@name='checkBox_ignore_version_update']")
        self.assertIsNotNone(box)
        checked = box.find("./property[@name='checked']/bool")
        self.assertTrue(checked is None or checked.text == "false")
        self.assertIsNotNone(tree.find(".//widget[@name='tab_2']//widget[@name='checkBox_ignore_version_update']"))


if __name__ == "__main__":
    unittest.main()
