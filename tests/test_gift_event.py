import unittest
from dataclasses import replace
from unittest.mock import Mock

from orders.gift_event import DEFAULT_GIFT_EVENT, GIFT_EVENT_KEY, read_gift_event, save_gift_event


class GiftEventTests(unittest.TestCase):
    def setUp(self):
        self.event = replace(DEFAULT_GIFT_EVENT, enabled=True, stores=("coupang",))

    def test_applies_only_highest_matching_tier_at_raw_won_amount(self):
        self.assertEqual(self.event.markdown_suffix("coupang", 59_999), "")
        for amount, expected in (
            (60_000, "볼펜 또는 박스커터 (랜덤 1개)"),
            (99_999, "볼펜 또는 박스커터 (랜덤 1개)"),
            (100_000, "에코백"),
            (150_000, "볼펜 + 박스커터 + 에코백"),
            (250_000, "볼펜 + 박스커터 + 에코백 + USB-C 30W 충전기"),
        ):
            with self.subTest(amount=amount):
                self.assertEqual(self.event.gift_name("coupang", amount), expected)
                self.assertEqual(self.event.markdown_suffix("coupang", amount), " · **🎁 사은품**")

    def test_disabled_unselected_and_missing_amount(self):
        self.assertEqual(DEFAULT_GIFT_EVENT.markdown_suffix("coupang", 250_000), "")
        self.assertEqual(self.event.markdown_suffix("naver", 250_000), "")
        self.assertEqual(self.event.markdown_suffix("coupang", float("nan")), "")

    def test_invoice_uses_only_the_selected_tier_label(self):
        for amount, expected in (
            (59_999, ""), (60_000, "[펜/커터] "), (100_000, "[에코백] "),
            (150_000, "[3종세트] "), (250_000, "[4종세트] "),
        ):
            with self.subTest(amount=amount):
                self.assertEqual(self.event.invoice_prefix("coupang", amount), expected)
        self.assertEqual(self.event.invoice_prefix("naver", 250_000), "")
        self.assertEqual(replace(self.event, enabled=False).invoice_prefix("coupang", 250_000), "")
        one_label = replace(self.event, invoice_labels=((60_000, "펜/커터"),))
        self.assertEqual(one_label.invoice_prefix("coupang", 100_000), "")

    def test_tiers_can_be_reordered_and_gift_text_is_preserved_for_preview(self):
        event = replace(self.event, tiers=((100_000, "[에코백]*"), (60_000, "볼펜")), invoice_labels=())
        self.assertEqual(event.gift_name("coupang", 100_000), "[에코백]*")
        with self.assertRaises(ValueError):
            replace(event, tiers=((60_000, "볼펜"), (60_000, "에코백")))
        with self.assertRaises(ValueError):
            replace(event, tiers=((60_000, "볼펜\n"),))

    def test_shared_setting_round_trip_and_stale_edit(self):
        worksheet = Mock()
        worksheet.get_all_values.return_value = [["키", "값"]]
        self.assertEqual(read_gift_event(worksheet), (DEFAULT_GIFT_EVENT, ""))

        edited = replace(self.event, tiers=((90_000, "새 사은품"),), invoice_labels=((90_000, "선물"),))
        save_gift_event(worksheet, edited, "")
        rows = worksheet.append_rows.call_args.args[0]
        self.assertEqual(rows[0][0], GIFT_EVENT_KEY)
        worksheet.get_all_values.return_value.extend(rows)
        saved, revision = read_gift_event(worksheet)
        self.assertEqual(saved, edited)

        with self.assertRaisesRegex(ValueError, "다른 작업자"):
            save_gift_event(worksheet, self.event, "")
        worksheet.batch_update.assert_not_called()

        save_gift_event(worksheet, replace(saved, enabled=False), revision)
        change = worksheet.batch_update.call_args.args[0][0]
        self.assertEqual(change["range"], "B2")
        worksheet.get_all_values.return_value[1][1] = change["values"][0][0]
        self.assertEqual(read_gift_event(worksheet)[0], replace(edited, enabled=False))

    def test_reads_shared_setting_created_by_older_app(self):
        worksheet = Mock()
        worksheet.get_all_values.return_value = [
            ["키", "값"], [GIFT_EVENT_KEY, '{"enabled": true, "stores": ["coupang"], "tiers": [[60000, "볼펜"]]}'],
        ]
        event, _ = read_gift_event(worksheet)
        self.assertEqual(event.gift_name("coupang", 60_000), "볼펜")
        self.assertEqual(event.invoice_prefix("coupang", 60_000), "")

    def test_ignores_markdown_labels_saved_by_trial_version(self):
        worksheet = Mock()
        worksheet.get_all_values.return_value = [
            ["키", "값"],
            [GIFT_EVENT_KEY, '{"enabled": true, "stores": ["coupang"], "tiers": [[60000, "볼펜"], [100000, "에코백"]]}'],
            ["order_gift_markdown", '{"60000": "사은품1", "100000": ""}'],
        ]
        event, _ = read_gift_event(worksheet)
        self.assertEqual(event.markdown_suffix("coupang", 60_000), " · **🎁 사은품**")
        self.assertEqual(event.markdown_suffix("coupang", 100_000), " · **🎁 사은품**")


if __name__ == "__main__":
    unittest.main()
