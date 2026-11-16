"""문의 알림: Slack 실패 시 새 문의 알림이 사라지지 않고, 여러 PC가 같은 주기에 중복 발송하지 않는지 검증한다."""

import contextlib
import io
import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

import sheets.google_sheets_oauth as oauth
from integrations import naver_commerce, slack_notify

from tests.test_order_load_sheets import load_app_module


class FakeConfigSheet:
    def __init__(self, values):
        self.rows = [["키", "값"]] + [[k, v] for k, v in values.items()]

    def get_all_values(self):
        return [list(row) for row in self.rows]

    def append_rows(self, rows, value_input_option=None):
        self.rows.extend([list(row) for row in rows])

    def batch_update(self, updates, value_input_option=None):
        for update in updates:
            row = int(update["range"][1:]) - 1
            self.rows[row][1] = update["values"][0][0]

    def value(self, key):
        return next((row[1] for row in self.rows[1:] if row[0] == key), "")


class FakeInquirySheet:
    col_count = 9

    def __init__(self, headers):
        self.rows = [list(headers)]

    def get_all_values(self):
        return [list(row) for row in self.rows]

    def update(self, values, range_name=None, value_input_option=None):
        self.rows[0] = list(values[0])

    def append_rows(self, rows, value_input_option=None):
        start = len(self.rows) + 1
        self.rows.extend([list(row) for row in rows])
        return {"updates": {"updatedRange": f"'문의알림'!A{start}:I{len(self.rows)}"}}

    def batch_update(self, updates, value_input_option=None):
        for update in updates:
            row = int(update["range"][1:]) - 1
            self.rows[row][7] = update["values"][0][0]


class InquiryAlertTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        self.config = FakeConfigSheet({
            "naver_client_id": "id", "naver_client_secret": "secret", "slack_webhook_url": "https://hooks.test/x",
        })
        self.inquiries = FakeInquirySheet(self.app.INQUIRY_SHEET_HEADERS)
        self.slack_ok = [False]
        self.slack_calls = []
        self.token_calls = []
        patches = [
            patch.object(oauth, "get_authorized_gspread_client", lambda *a, **k: object()),
            patch.object(self.app, "_standalone_open_config_ws", lambda gc: self.config),
            patch.object(self.app, "_standalone_open_inquiry_ws", lambda gc: self.inquiries),
            patch.object(self.app, "_inquiry_alerts_allowed", lambda *a: True),
            patch.object(naver_commerce, "get_access_token", lambda *a: self.token_calls.append(1) or "T"),
            patch.object(naver_commerce, "fetch_product_qnas", lambda *a, **k: [
                {"questionId": 7, "createDate": "2026-11-16", "productName": "상품", "question": "재고 있나요?"}]),
            patch.object(naver_commerce, "fetch_customer_inquiries", lambda *a, **k: []),
            patch.object(slack_notify, "send_slack", self.fake_slack),
        ]
        for item in patches:
            item.start()
            self.addCleanup(item.stop)

    def fake_slack(self, webhook, text):
        self.slack_calls.append(text)
        return {"ok": True} if self.slack_ok[0] else {"ok": False, "error": "네트워크 오류"}

    def poll(self):
        with contextlib.redirect_stdout(io.StringIO()):
            return self.app.run_naver_inquiry_poll_worker()

    def test_new_inquiry_is_notified_again_after_slack_failure(self):
        first = self.poll()
        self.assertTrue(first["ok"], first)
        self.assertEqual(first["sent"], 0)
        self.assertEqual(self.inquiries.rows[1][7], "")  # 알리지 못했으므로 최근알림시각을 비워 둔다
        self.slack_ok[0] = True
        second = self.poll()
        self.assertEqual(second["sent"], 1)
        self.assertEqual(len(self.slack_calls), 2)
        self.assertTrue(self.inquiries.rows[1][7])  # 발송 성공 후 기록
        self.assertEqual(len(self.inquiries.rows), 2)  # 같은 문의를 다시 추가하지 않음

    def test_recent_claim_by_other_pc_skips_poll(self):
        recent = (datetime.now() - timedelta(seconds=30)).strftime("%Y-%m-%d %H:%M:%S")
        self.config.rows.append(["inquiry_poll_claim", f"{recent}|다른PC@HOST"])
        result = self.poll()
        self.assertTrue(result["ok"], result)
        self.assertTrue(result.get("skipped_other_pc"))
        self.assertEqual(self.token_calls, [])
        self.assertEqual(self.slack_calls, [])
        self.assertEqual(len(self.inquiries.rows), 1)

    def test_poll_claims_the_slot(self):
        self.slack_ok[0] = True
        self.poll()
        claim = self.config.value("inquiry_poll_claim")
        self.assertIn("|", claim)
        self.assertEqual(claim.split("|", 1)[1], self.app._current_actor())


if __name__ == "__main__":
    unittest.main()
