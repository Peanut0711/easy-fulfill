import json
import unittest
from dataclasses import replace
from unittest.mock import Mock

from orders.gift_event import DEFAULT_GIFT_EVENT, GIFT_EVENT_KEY, GIFT_MARKDOWN_KEY, read_gift_event, save_gift_event


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
                self.assertEqual(self.event.markdown_suffix("coupang", amount), f" · **🎁 사은품: {expected}**")

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

    def test_tiers_can_be_reordered_and_gift_text_is_escaped(self):
        event = replace(self.event, tiers=((100_000, "[에코백]*"), (60_000, "볼펜")), invoice_labels=())
        self.assertEqual(event.markdown_suffix("coupang", 100_000), r" · **🎁 사은품: \[에코백\]\***")
        with self.assertRaises(ValueError):
            replace(event, tiers=((60_000, "볼펜"), (60_000, "에코백")))
        with self.assertRaises(ValueError):
            replace(event, tiers=((60_000, "볼펜\n"),))

    def test_shared_setting_round_trip_and_stale_edit(self):
        worksheet = Mock()
        worksheet.get_all_values.return_value = [["키", "값"]]
        event, empty_revision = read_gift_event(worksheet)
        self.assertEqual(event, DEFAULT_GIFT_EVENT)

        edited = replace(self.event, tiers=((90_000, "새 사은품"),), invoice_labels=((90_000, "선물"),))
        save_gift_event(worksheet, edited, empty_revision)
        rows = worksheet.append_rows.call_args.args[0]
        self.assertEqual(rows[0][0], GIFT_EVENT_KEY)
        worksheet.get_all_values.return_value.extend(rows)
        saved, revision = read_gift_event(worksheet)
        self.assertEqual(saved, edited)

        with self.assertRaisesRegex(ValueError, "다른 작업자"):
            save_gift_event(worksheet, self.event, empty_revision)
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
        self.assertIn("볼펜", event.markdown_suffix("coupang", 60_000))
        self.assertEqual(event.invoice_prefix("coupang", 60_000), "")

    def test_markdown_labels_customize_or_hide_name_for_the_matching_tier(self):
        event = replace(self.event, markdown_labels=((60_000, "사은품1"), (100_000, ""), (250_000, "[사은품4]")))
        self.assertEqual(event.markdown_suffix("coupang", 99_999), " · **🎁 사은품1**")
        self.assertEqual(event.markdown_suffix("coupang", 120_000), " · **🎁**")
        self.assertEqual(event.markdown_suffix("coupang", 150_000), " · **🎁 사은품: 볼펜 + 박스커터 + 에코백**")
        self.assertEqual(event.markdown_suffix("coupang", 250_000), r" · **🎁 \[사은품4\]**")
        self.assertEqual(event.gift_name("coupang", 120_000), "에코백")
        self.assertEqual(event.invoice_prefix("coupang", 120_000), "[에코백] ")

    def test_shared_markdown_labels_survive_legacy_gift_updates(self):
        worksheet = Mock()
        worksheet.get_all_values.return_value = [["키", "값"]]
        save_gift_event(worksheet, self.event, read_gift_event(worksheet)[1])
        base_row = worksheet.append_rows.call_args.args[0][0]
        worksheet.get_all_values.return_value.append(base_row)
        event, revision = read_gift_event(worksheet)

        edited = replace(event, markdown_labels=((60_000, "사은품1"), (100_000, "")))
        save_gift_event(worksheet, edited, revision)
        markdown_row = worksheet.append_rows.call_args.args[0][0]
        self.assertEqual(markdown_row, [GIFT_MARKDOWN_KEY, '{"60000": "사은품1", "100000": ""}'])
        worksheet.get_all_values.return_value.append(markdown_row)
        self.assertEqual(read_gift_event(worksheet)[0], edited)
        worksheet.batch_update.assert_not_called()

        with self.assertRaisesRegex(ValueError, "다른 작업자"):
            save_gift_event(worksheet, event, revision)

        base = json.loads(base_row[1])
        base["tiers"] = [[60_000, "변경된 사은품"]]
        base["invoice_labels"] = {"60000": "선물"}
        base_row[1] = json.dumps(base)
        updated, revision = read_gift_event(worksheet)
        self.assertEqual(updated.gift_name("coupang", 60_000), "변경된 사은품")
        self.assertEqual(updated.markdown_labels, ((60_000, "사은품1"),))
        save_gift_event(worksheet, replace(updated, markdown_labels=()), revision)
        change = worksheet.batch_update.call_args.args[0][-1]
        self.assertEqual(change, {"range": "B3", "values": [["{}"]]})
        markdown_row[1] = change["values"][0][0]
        self.assertEqual(read_gift_event(worksheet)[0].markdown_suffix("coupang", 60_000),
                         " · **🎁 사은품: 변경된 사은품**")

    def test_first_markdown_label_and_existing_event_change_share_one_write(self):
        worksheet = Mock(id=19)
        worksheet.get_all_values.return_value = [["키", "값"]]
        save_gift_event(worksheet, self.event, read_gift_event(worksheet)[1])
        worksheet.get_all_values.return_value.extend(worksheet.append_rows.call_args.args[0])
        worksheet.reset_mock()
        event, revision = read_gift_event(worksheet)
        save_gift_event(worksheet, replace(event, enabled=False, markdown_labels=((60_000, ""),)), revision)

        worksheet.append_rows.assert_not_called()
        worksheet.batch_update.assert_not_called()
        update, append = worksheet.spreadsheet.batch_update.call_args.args[0]["requests"]
        self.assertEqual(update["updateCells"]["start"], {"sheetId": 19, "rowIndex": 1, "columnIndex": 1})
        self.assertEqual(append["appendCells"]["rows"][0]["values"], [
            {"userEnteredValue": {"stringValue": GIFT_MARKDOWN_KEY}},
            {"userEnteredValue": {"stringValue": '{"60000": ""}'}},
        ])


if __name__ == "__main__":
    unittest.main()
