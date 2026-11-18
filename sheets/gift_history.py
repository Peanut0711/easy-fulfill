"""공유 사은품 기록과 반영 대기 중인 로컬 기록을 관리한다."""

from contextlib import closing
from datetime import date, datetime
import sqlite3

from orders.gift_event import GIFT_EVENT_STORES
from orders.gift_history import GiftRecord


SHEET_TITLE = "사은품 기록"
HEADERS = ("기록일", "스토어", "사은품 구간", "최소 구매 금액", "지급 사은품", "기록 키", "기록 시각")
DATA_RANGE = "A:G"


def _find_worksheet(spreadsheet):
    from gspread.exceptions import WorksheetNotFound

    try:
        return spreadsheet.worksheet(SHEET_TITLE)
    except WorksheetNotFound:
        return None


def _to_row(record):
    return (record.day.isoformat(), record.store, record.tier, record.minimum, record.gift,
            record.key, record.recorded_at)


def _from_row(row):
    day, store, tier, minimum, gift, key, recorded_at = row[:len(HEADERS)]
    return GiftRecord(date.fromisoformat(day), store, int(tier), int(minimum), gift, key, recorded_at)


def _parse_records(rows) -> tuple[GiftRecord, ...]:
    if tuple(rows[0][:len(HEADERS)]) != HEADERS:
        raise ValueError(f"공유 시트 「{SHEET_TITLE}」의 첫 행 제목을 확인해 주세요.")
    records = {}
    for number, row in enumerate(rows[1:], start=2):
        if not any(row[:len(HEADERS)]):
            continue
        try:
            record = _from_row(row)
            datetime.fromisoformat(record.recorded_at)
            if (record.store not in GIFT_EVENT_STORES or record.tier < 1 or record.minimum < 1
                    or not record.gift or not record.key):
                raise ValueError
        except (TypeError, ValueError) as error:
            raise ValueError(f"공유 시트 「{SHEET_TITLE}」의 {number}행을 확인해 주세요.") from error
        # 서로 다른 PC에서 동시에 저장해도 집계에는 가장 먼저 표시된 한 건만 반영한다.
        previous = records.get(record.key)
        if previous is None or record.recorded_at < previous.recorded_at:
            records[record.key] = record
    return tuple(records.values())


def read_gift_records(spreadsheet) -> tuple[GiftRecord, ...]:
    worksheet = _find_worksheet(spreadsheet)
    if worksheet is None:
        return ()
    rows = worksheet.get_all_values(DATA_RANGE, value_render_option="UNFORMATTED_VALUE")
    return _parse_records(rows) if any(rows) else ()


def append_gift_records(spreadsheet, records) -> None:
    from gspread.exceptions import APIError

    worksheet = _find_worksheet(spreadsheet)
    if worksheet is None:
        try:
            worksheet = spreadsheet.add_worksheet(title=SHEET_TITLE, rows=1000, cols=len(HEADERS))
        except APIError:
            # 두 PC가 동시에 최초 기록을 올리면 먼저 생성된 탭을 사용한다.
            worksheet = _find_worksheet(spreadsheet)
            if worksheet is None:
                raise

    existing = worksheet.get_all_values(DATA_RANGE, value_render_option="UNFORMATTED_VALUE")
    if any(existing):
        seen = {record.key: record.recorded_at for record in _parse_records(existing)}
    else:
        worksheet.update([list(HEADERS)], range_name="A1:G1", value_input_option="RAW")
        seen = {}
    rows = []
    for record in records:
        first_seen = seen.get(record.key)
        if first_seen is None or record.recorded_at < first_seen:
            seen[record.key] = record.recorded_at
            rows.append(list(_to_row(record)))
    if rows:
        worksheet.append_rows(rows, table_range=DATA_RANGE, value_input_option="RAW", insert_data_option="INSERT_ROWS")


class GiftRecordOutbox:
    """네트워크 오류나 앱 재시작 때 원래 집계일로 다시 반영한다."""

    def __init__(self, path):
        self.path = path

    def add(self, records) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute("""
                CREATE TABLE IF NOT EXISTS pending (
                    day TEXT, store TEXT, tier INTEGER, minimum INTEGER, gift TEXT,
                    key TEXT PRIMARY KEY, recorded_at TEXT
                )
            """)
            db.executemany("INSERT OR IGNORE INTO pending VALUES (?, ?, ?, ?, ?, ?, ?)", map(_to_row, records))

    def list_pending(self) -> tuple[GiftRecord, ...]:
        if not self.path.exists():
            return ()
        with closing(sqlite3.connect(self.path)) as db:
            rows = db.execute("SELECT day, store, tier, minimum, gift, key, recorded_at FROM pending")
            return tuple(map(_from_row, rows))

    def remove(self, records) -> None:
        with closing(sqlite3.connect(self.path)) as db, db:
            db.executemany("DELETE FROM pending WHERE key = ? AND recorded_at = ?",
                           ((record.key, record.recorded_at) for record in records))
