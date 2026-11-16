"""배송추적 새로고침(전체·단건)이 사용자가 적은 메모(K열)를 지우지 않는지 검증한다."""

import re
import unittest
from datetime import datetime
from unittest.mock import patch

import sheets.google_sheets_oauth as oauth
import tracking.kpost_tracker as kpost_tracker
import tracking.workers as workers
from tracking.repository import open_tracking_worksheet


def _col(letter):
    return ord(letter) - ord("A")


class FakeTrackingSheet:
    """batch_update로 받은 A1 범위를 메모리 표에 반영하는 최소 워크시트."""

    def __init__(self, rows):
        self.rows = [list(row) for row in rows]

    def get_all_values(self):
        width = max(len(row) for row in self.rows)
        return [row + [""] * (width - len(row)) for row in self.rows]

    def batch_update(self, updates, value_input_option=None):
        for update in updates:
            self._write(update["range"], update["values"])

    def update(self, values, range_name=None, value_input_option=None):
        self._write(range_name, values)

    def batch_get(self, ranges):
        out = []
        for range_name in ranges:
            match = re.match(r"([A-Z])(\d+):([A-Z])(\d*)", range_name)
            first, last = _col(match.group(1)), _col(match.group(3))
            start = int(match.group(2)) - 1
            end = int(match.group(4)) if match.group(4) else len(self.rows)
            block = []
            for line in self.rows[start:end]:
                block.append([line[i] if i < len(line) else "" for i in range(first, last + 1)])
            out.append(block)
        return out

    def _write(self, range_name, values):
        match = re.match(r"([A-Z])(\d+)", range_name)
        col, row = _col(match.group(1)), int(match.group(2)) - 1
        for r, line in enumerate(values):
            while len(self.rows) <= row + r:
                self.rows.append([])
            target = self.rows[row + r]
            for c, value in enumerate(line):
                while len(target) <= col + c:
                    target.append("")
                target[col + c] = value

    def cell(self, row, letter):
        line = self.rows[row - 1]
        index = _col(letter)
        return line[index] if index < len(line) else ""


HEADER = list(workers.TRACKING_SHEET_HEADERS)


def tracking_rows():
    today = datetime.now().strftime("%Y-%m-%d")
    rows = [HEADER]
    for index, memo in enumerate(("부재 시 경비실", "고객 재요청", "파손 문의")):
        row = [f"600000000000{index}", f"{today} 08:00:00", "naver", "1", "김", "우체국",
               "배달준비", "N", "서울", "", memo, "2026.01.01 10:00", "추적중"]
        rows.append(row)
    return rows


RESULTS = {
    "6000000000000": {"ok": True, "status": "배달완료", "complete": True, "where": "강남", "time": "2026.11.16 12:00"},
    "6000000000001": {"ok": False, "error_code": "ERR-001", "error": "조회결과 없음"},
    "6000000000002": {"ok": False, "error_code": "ERR-999", "error": "일시 오류"},
}


class TrackingRefreshMemoTests(unittest.TestCase):
    def run_refresh(self, sheet, single=None):
        with patch.object(oauth, "get_authorized_gspread_client", lambda: object()), \
                patch.object(workers, "_standalone_open_tracking_ws", lambda gc: sheet), \
                patch.object(kpost_tracker, "summarize_tracking", lambda key, regino: RESULTS[regino]), \
                patch.object(workers, "KPOST_TRACKING_REQUEST_DELAY_SEC", 0):
            if single:
                return workers.run_tracking_refresh_one_worker("KEY", single)
            return workers.run_tracking_refresh_worker("KEY", scope="all")

    def test_full_refresh_keeps_memos_and_records_error_elsewhere(self):
        sheet = FakeTrackingSheet(tracking_rows())
        result = self.run_refresh(sheet)
        self.assertTrue(result["ok"], result)
        self.assertEqual([sheet.cell(row, "K") for row in (2, 3, 4)],
                         ["부재 시 경비실", "고객 재요청", "파손 문의"])
        self.assertEqual(sheet.cell(2, "G"), "배달완료")
        self.assertEqual(sheet.cell(3, "G"), "추적정보 없음")
        self.assertEqual(sheet.cell(4, "N"), "일시 오류")
        self.assertEqual(sheet.cell(2, "N"), "")

    def test_single_refresh_keeps_memo_for_each_outcome(self):
        for regino, row in (("6000000000000", 2), ("6000000000001", 3), ("6000000000002", 4)):
            sheet = FakeTrackingSheet(tracking_rows())
            memo = sheet.cell(row, "K")
            result = self.run_refresh(sheet, single=regino)
            self.assertTrue(result["ok"], result)
            self.assertEqual(sheet.cell(row, "K"), memo, regino)
        self.assertEqual(sheet.cell(4, "N"), "일시 오류")

    def test_refresh_keeps_manual_stop_set_during_refresh(self):
        sheet = FakeTrackingSheet(tracking_rows())

        def summarize(key, regino):
            if regino == "6000000000000":
                sheet._write("M3", [["수동 중지"]])  # 새로고침 도중 다른 PC가 3행 추적 중지
            return RESULTS[regino]

        with patch.object(oauth, "get_authorized_gspread_client", lambda: object()), \
                patch.object(workers, "_standalone_open_tracking_ws", lambda gc: sheet), \
                patch.object(kpost_tracker, "summarize_tracking", summarize), \
                patch.object(workers, "KPOST_TRACKING_REQUEST_DELAY_SEC", 0):
            result = workers.run_tracking_refresh_worker("KEY", scope="all")
        self.assertTrue(result["ok"], result)
        self.assertEqual(sheet.cell(3, "M"), "수동 중지")
        self.assertEqual(sheet.cell(3, "G"), "추적정보 없음")  # 조회 결과는 그대로 반영
        self.assertEqual(sheet.cell(2, "M"), "추적중")

    def test_single_refresh_keeps_manual_stop_set_during_lookup(self):
        expected = {"6000000000000": ("G", "배달완료"), "6000000000001": ("G", "추적정보 없음"),
                    "6000000000002": ("N", "일시 오류")}
        for regino, row in (("6000000000000", 2), ("6000000000001", 3), ("6000000000002", 4)):
            sheet = FakeTrackingSheet(tracking_rows())

            def summarize(key, number, row=row):
                sheet._write(f"M{row}", [["수동 중지"]])  # 단건 조회 도중 다른 PC가 추적 중지
                return RESULTS[number]

            with patch.object(oauth, "get_authorized_gspread_client", lambda: object()), \
                    patch.object(workers, "_standalone_open_tracking_ws", lambda gc: sheet), \
                    patch.object(kpost_tracker, "summarize_tracking", summarize):
                result = workers.run_tracking_refresh_one_worker("KEY", regino)
            self.assertTrue(result["ok"], result)
            self.assertEqual(sheet.cell(row, "M"), "수동 중지", regino)
            column, value = expected[regino]
            self.assertEqual(sheet.cell(row, column), value, regino)  # 조회 결과는 그대로 반영
            self.assertEqual(sheet.cell(row, "K"), tracking_rows()[row - 1][10], regino)

    def test_success_clears_previous_lookup_error(self):
        rows = tracking_rows()
        rows[1] += ["지난 오류"]
        sheet = FakeTrackingSheet(rows)
        self.run_refresh(sheet, single="6000000000000")
        self.assertEqual(sheet.cell(2, "N"), "")
        self.assertEqual(sheet.cell(2, "K"), "부재 시 경비실")


class FakeWorksheet:
    def __init__(self, title, header, col_count):
        self.title = title
        self.header = list(header)
        self.col_count = col_count
        self.calls = []

    def row_values(self, row):
        return list(self.header)

    def add_cols(self, count):
        self.calls.append(("add_cols", count))
        self.col_count += count

    def update(self, values, range_name=None, value_input_option=None):
        self.calls.append(("update", range_name, values[0]))


class OpenTrackingWorksheetTests(unittest.TestCase):
    def test_existing_13_column_sheet_gets_error_column(self):
        worksheet = FakeWorksheet("송장추적", HEADER[:13], 13)
        spreadsheet = type("S", (), {"worksheets": lambda self: [worksheet]})()
        gc = type("G", (), {"open_by_key": lambda self, key: spreadsheet})()
        self.assertIs(open_tracking_worksheet(gc, "id", "송장추적", HEADER), worksheet)
        self.assertEqual(worksheet.calls[0], ("add_cols", len(HEADER) - 13))
        self.assertEqual(worksheet.calls[1][1], f"A1:{chr(ord('A') + len(HEADER) - 1)}1")
        self.assertEqual(HEADER[13], "최근조회오류")


if __name__ == "__main__":
    unittest.main()
