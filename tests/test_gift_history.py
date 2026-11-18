import os
import tempfile
import unittest
from collections import Counter
from dataclasses import replace
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import Mock, patch

from gspread.exceptions import WorksheetNotFound

from orders.gift_event import DEFAULT_GIFT_EVENT
from orders.gift_history import build_gift_records, summarize_gift_records
from sheets.gift_history import GiftRecordOutbox, HEADERS, SHEET_TITLE, append_gift_records, read_gift_records


STAMP = datetime(2026, 11, 17, 15, 1, tzinfo=timezone.utc)
EVENT = replace(DEFAULT_GIFT_EVENT, enabled=True)


def _records(store="coupang", order_id="101", amount=100_000, stamp=STAMP):
    field = {"coupang": "결제액", "naver": "주문총액"}[store]
    return build_gift_records(store, {order_id: {field: amount}}, EVENT, stamp)


class GiftRecordTests(unittest.TestCase):
    def test_highest_tier_only_and_korean_date(self):
        self.assertEqual(_records(amount=59_999), ())
        for tier, amount in enumerate((60_000, 100_000, 150_000, 250_000), start=1):
            record, = _records(amount=amount)
            self.assertEqual((record.day, record.tier, record.minimum), (date(2026, 11, 18), tier, amount))
            self.assertEqual(record.gift, EVENT.gift_name("coupang", amount))
        self.assertEqual(build_gift_records("coupang", {"101": {"결제액": 100_000}}, DEFAULT_GIFT_EVENT, STAMP), ())

    def test_key_is_stable_across_order_number_type_and_date(self):
        first, = _records(order_id="123.0")
        repeat, = _records(order_id="123", amount=250_000, stamp=STAMP + timedelta(days=1))
        other_store, = _records(store="naver", order_id="123")
        self.assertEqual(first.key, repeat.key)
        self.assertNotEqual(first.key, other_store.key)

    def test_gmarket_bundle_has_one_gift_and_stable_key_in_either_order(self):
        delivery = {"수령인명": "수취인", "주소": "서울 1", "수령인 휴대폰": "01000000000", "상품목록": []}
        first = dict(delivery, 판매금액=55_000, **{"배송비 금액": 3_000})
        second = dict(delivery, 판매금액=42_000)
        first_order, = build_gift_records("gmarket", {"201": first, "202": second}, EVENT, STAMP)
        reversed_order, = build_gift_records("gmarket", {"202": second, "201": first}, EVENT, STAMP)
        self.assertEqual((first_order.tier, first_order.key), (2, reversed_order.key))
        self.assertNotIn("수취인", str(first_order))
        self.assertNotIn("01000000000", str(first_order))

    def test_naver_same_order_number_for_two_destinations_counts_both(self):
        first = {"주문번호목록": ["5100000000000"], "주문총액": 60_000, "통합배송지": "서울 1",
                 "수취인연락처1": "010-0000-0000", "우편번호": "06035"}
        second = dict(first, 통합배송지="부산 1")
        orders = {"5100000000000": first, "5100000000000|배송지2": second}
        records = build_gift_records("naver", orders, EVENT, STAMP)
        self.assertEqual(len({record.key for record in records}), 2)
        exported = dict(first, 통합배송지="서울  1", 우편번호="6035.0", 수취인연락처1="01000000000")
        self.assertEqual(build_gift_records("naver", {"5100000000000": exported}, EVENT, STAMP)[0], records[0])


class GiftHistoryStorageTests(unittest.TestCase):
    def setUp(self):
        self.rows = []
        self.spreadsheet = Mock()
        self.worksheet = self.spreadsheet.add_worksheet.return_value
        self.worksheet.get_all_values.side_effect = lambda *args, **kwargs: self.rows.copy() or [[]]
        self.worksheet.update.side_effect = lambda rows, **kwargs: self.rows.extend(rows)
        self.worksheet.append_rows.side_effect = lambda rows, **kwargs: self.rows.extend(rows)

    def test_empty_read_first_save_and_repeat_after_reopen(self):
        self.spreadsheet.worksheet.side_effect = WorksheetNotFound(SHEET_TITLE)
        self.assertEqual(read_gift_records(self.spreadsheet), ())
        self.spreadsheet.add_worksheet.assert_not_called()

        records = _records()
        append_gift_records(self.spreadsheet, records)
        self.assertEqual(self.rows[0], list(HEADERS))
        self.assertEqual(self.worksheet.append_rows.call_args.kwargs["value_input_option"], "RAW")
        self.spreadsheet.worksheet.side_effect = None
        self.spreadsheet.worksheet.return_value = self.worksheet
        append_gift_records(self.spreadsheet, _records(stamp=STAMP + timedelta(days=1), amount=150_000))
        self.assertEqual(read_gift_records(self.spreadsheet), records)
        self.worksheet.append_rows.assert_called_once()

    def test_offline_earliest_record_and_duplicate_append_count_once(self):
        early = _records()
        later = _records(stamp=STAMP + timedelta(days=1), amount=150_000)
        self.spreadsheet.worksheet.return_value = self.worksheet
        append_gift_records(self.spreadsheet, later)
        append_gift_records(self.spreadsheet, early)
        self.rows.append(self.rows[-1].copy())  # 두 PC가 같은 기록 키를 읽기 전에 동시에 추가
        naver, = _records("naver", stamp=STAMP + timedelta(days=1))
        append_gift_records(self.spreadsheet, [naver])
        records = read_gift_records(self.spreadsheet)
        self.assertCountEqual(records, [early[0], naver])
        totals, daily = summarize_gift_records(records, early[0].day, naver.day, "coupang")
        self.assertEqual(totals, Counter({(2, 100_000, "에코백"): 1}))
        self.assertEqual(daily, [(early[0].day, totals)])

    def test_outbox_persists_before_upload_and_only_removes_synced_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "gift_history.sqlite3"
            outbox = GiftRecordOutbox(path)
            self.assertEqual(outbox.list_pending(), ())
            self.assertFalse(path.exists())
            first = _records()
            other = _records(order_id="202", stamp=STAMP + timedelta(days=1))
            outbox.add(first)
            outbox.add(_records(stamp=STAMP + timedelta(days=1)))
            restarted = GiftRecordOutbox(path)
            self.assertEqual(restarted.list_pending(), first)
            outbox.add(other)
            restarted.remove(first)
            self.assertEqual(restarted.list_pending(), other)


class GiftHistoryReportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication

        cls.app = QApplication.instance() or QApplication([])

    def test_sync_retries_and_keeps_original_day(self):
        from PySide6.QtCore import QObject
        from ui.gift_history import GiftHistorySync

        with tempfile.TemporaryDirectory() as directory:
            outbox = GiftRecordOutbox(Path(directory) / "gift_history.sqlite3")
            spreadsheet = object()
            opener = Mock(side_effect=[OSError("오프라인"), spreadsheet])
            parent = QObject()
            sync = GiftHistorySync(opener, outbox, parent)
            records = _records()
            with patch("ui.gift_history.append_gift_records") as append:
                with self.assertLogs(level="WARNING"):
                    sync.enqueue(records)
                    sync._worker.join(5)
                self.assertEqual(outbox.list_pending(), records)
                sync.flush()
                sync._worker.join(5)
                append.assert_called_once_with(spreadsheet, records)
                self.assertEqual(outbox.list_pending(), ())

    def test_report_switches_to_all_dates_and_filters_store(self):
        from ui.gift_history import GiftHistoryDialog

        records = (*_records(), *_records("naver", amount=250_000, stamp=STAMP + timedelta(days=2)))
        dialog = GiftHistoryDialog(lambda: (records, 1))
        dialog.show()
        self.app.processEvents()
        dialog._all_dates.setChecked(True)
        self.assertEqual(dialog._table.item(0, 5).text(), "2")
        self.assertEqual(dialog._table.item(1, 0).text(), "2026-11-20")
        self.assertEqual(dialog._table.item(1, 4).text(), "1")
        self.assertIn("반영 대기", dialog._status.text())
        dialog._store.setCurrentIndex(dialog._store.findData("coupang"))
        self.assertEqual(dialog._table.rowCount(), 2)
        self.assertEqual(dialog._table.item(0, 2).text(), "1")
        self.assertIn("에코백", dialog._table.item(0, 2).toolTip())
        dialog.close()
