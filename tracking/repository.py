"""배송추적과 공유 설정 Google Sheets 접근 함수."""

from __future__ import annotations

from datetime import datetime
from typing import Mapping, Sequence

import pandas as pd


TRACKING_DETAIL_READ_BATCH_SIZE = 100


def open_tracking_worksheet(
    gc, spreadsheet_id: str, sheet_title: str, headers: Sequence[str],
):
    """배송추적 시트를 제목으로 열고 없으면 현재 스키마로 생성한다."""
    spreadsheet = gc.open_by_key(spreadsheet_id)
    for worksheet in spreadsheet.worksheets():
        if worksheet.title == sheet_title:
            if len(worksheet.row_values(1)) < len(headers):
                worksheet.update([list(headers)], range_name="A1:M1", value_input_option="RAW")
            return worksheet
    worksheet = spreadsheet.add_worksheet(title=sheet_title, rows=2000, cols=len(headers))
    worksheet.update([list(headers)], range_name="A1", value_input_option="RAW")
    return worksheet


def open_config_worksheet(
    gc, spreadsheet_id: str, sheet_title: str, headers: Sequence[str],
):
    """공유 설정 시트를 제목으로 열고 없으면 현재 스키마로 생성한다."""
    spreadsheet = gc.open_by_key(spreadsheet_id)
    for worksheet in spreadsheet.worksheets():
        if worksheet.title == sheet_title:
            return worksheet
    worksheet = spreadsheet.add_worksheet(title=sheet_title, rows=50, cols=2)
    worksheet.update([list(headers)], range_name="A1", value_input_option="RAW")
    return worksheet


def normalize_tracking_no(value) -> str:
    """등기번호 비교와 Sheet 저장에 쓰는 문자열 정규화."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    return text[:-2] if text.endswith(".0") else text


def read_config_values_map(worksheet) -> dict[str, str]:
    """설정 시트의 키/값 행을 dict로 반환한다."""
    config = {}
    for row in worksheet.get_all_values()[1:]:
        if not row:
            continue
        key = (row[0] or "").strip()
        if key:
            config[key] = (row[1] if len(row) > 1 else "").strip()
    return config


def write_config_values(worksheet, updates: Mapping[str, object]) -> None:
    """설정 시트의 키/값을 현재 행에 갱신하거나 새 행으로 추가한다."""
    key_to_row = {}
    for row_index, row in enumerate(worksheet.get_all_values()[1:], start=2):
        key = (row[0] if row else "").strip()
        if key and key not in key_to_row:
            key_to_row[key] = row_index
    pairs = [(key, str(value)) for key, value in (updates or {}).items() if value is not None]
    new_rows = []
    batch_updates = []
    for key, value in pairs:
        if key in key_to_row:
            batch_updates.append({"range": f"B{key_to_row[key]}", "values": [[value]]})
        else:
            new_rows.append([key, value])
    if new_rows:
        worksheet.append_rows(new_rows, value_input_option="RAW")
    if batch_updates:
        worksheet.batch_update(batch_updates, value_input_option="RAW")


def read_tracking_values(worksheet) -> list[list[str]]:
    """배송추적 시트 전체 값을 읽는다."""
    return worksheet.get_all_values()


def read_tracking_list_metadata(worksheet) -> tuple[list[str], list[list[str]]]:
    """목록 선별에 필요한 좁은 열만 일괄로 읽는다.

    반환 순서는 헤더, 등록일시(B), 완료(H), 최근이벤트(L), 관리상태(M)다.
    """
    values = worksheet.batch_get(["A1:M1", "B2:B", "H2:H", "L2:L", "M2:M"])
    header = values[0][0] if values and values[0] else []
    columns = [
        [row[0] if row else "" for row in (value or [])]
        for value in values[1:]
    ]
    return header, columns


def read_tracking_rows(worksheet, row_numbers: Sequence[int]) -> list[list[str]]:
    """선별된 행의 A:M 상세값만 한 번에 읽는다."""
    if not row_numbers:
        return []
    rows = []
    for start in range(0, len(row_numbers), TRACKING_DETAIL_READ_BATCH_SIZE):
        batch = row_numbers[start:start + TRACKING_DETAIL_READ_BATCH_SIZE]
        values = worksheet.batch_get([f"A{row}:M{row}" for row in batch])
        rows.extend(value[0] if value else [] for value in values)
    return rows


def batch_update_tracking(worksheet, updates) -> None:
    """배송추적 시트에 현재 payload 형식의 일괄 업데이트를 반영한다."""
    if updates:
        worksheet.batch_update(updates, value_input_option="RAW")


def update_tracking_cell(worksheet, row_index: int, column: str, value: str) -> None:
    """배송추적 시트 단일 셀을 현재 RAW 방식으로 갱신한다."""
    worksheet.update([[value]], range_name=f"{column}{row_index}", value_input_option="RAW")


def upsert_tracking_records(
    worksheet, records, headers: Sequence[str], management_active: str,
) -> dict[str, int]:
    """등기번호 기준으로 신규 행을 추가하고 기존 행의 비어 있는 기본정보만 채운다."""
    normalized = []
    seen = set()
    for record in records or []:
        regino = normalize_tracking_no(record.get("등기번호"))
        if not regino or regino in seen:
            continue
        seen.add(regino)
        normalized.append({
            "등기번호": regino,
            "스토어": str(record.get("스토어", "") or "").strip(),
            "주문번호": normalize_tracking_no(record.get("주문번호")),
            "수취인명": str(record.get("수취인명", "") or "").strip(),
        })
    if not normalized:
        return {"registered": 0, "updated": 0}

    values = read_tracking_values(worksheet)
    if not values:
        worksheet.update([list(headers)], range_name="A1", value_input_option="RAW")
        values = [list(headers)]
    key_to_row = {}
    for row_index, row in enumerate(values[1:], start=2):
        regino = (row[0] if row else "").strip()
        if regino and regino not in key_to_row:
            key_to_row[regino] = (row_index, row)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = []
    updates = []
    registered = updated = 0
    for record in normalized:
        regino = record["등기번호"]
        if regino in key_to_row:
            row_index, row = key_to_row[regino]
            if row_index is None:
                continue
            merged = list(row) + [""] * (len(headers) - len(row))
            changed = False
            for column_index, value in ((2, record["스토어"]), (3, record["주문번호"]), (4, record["수취인명"])):
                if value and not (merged[column_index] or "").strip():
                    merged[column_index] = value
                    changed = True
            if changed:
                updates.append({"range": f"C{row_index}:E{row_index}",
                                "values": [[merged[2], merged[3], merged[4]]]})
                updated += 1
        else:
            new_rows.append([
                regino, timestamp, record["스토어"], record["주문번호"], record["수취인명"],
                "우체국", "", "", "", "", "", "", management_active,
            ])
            key_to_row[regino] = (None, None)
            registered += 1
    if new_rows:
        worksheet.append_rows(new_rows, value_input_option="RAW")
    batch_update_tracking(worksheet, updates)
    return {"registered": registered, "updated": updated}


def update_tracking_management(worksheet, reginos, management: str) -> dict[str, int]:
    """선택된 등기번호의 M열 관리상태를 일괄 갱신한다."""
    keys = {normalize_tracking_no(regino) for regino in reginos or []}
    keys.discard("")
    found = set()
    updates = []
    for row_index, row in enumerate(read_tracking_values(worksheet)[1:], start=2):
        regino = normalize_tracking_no(row[0] if row else "")
        if regino in keys:
            found.add(regino)
            updates.append({"range": f"M{row_index}", "values": [[management]]})
    batch_update_tracking(worksheet, updates)
    return {"updated": len(found), "missing": len(keys) - len(found)}


def update_tracking_notes(worksheet, notes: Mapping[str, object]) -> int:
    """등기번호별 메모를 K열에 일괄 갱신한다."""
    normalized = {
        normalize_tracking_no(regino): str(note or "")
        for regino, note in (notes or {}).items()
        if normalize_tracking_no(regino)
    }
    updates = []
    for row_index, row in enumerate(read_tracking_values(worksheet)[1:], start=2):
        regino = normalize_tracking_no(row[0] if row else "")
        if regino in normalized:
            updates.append({"range": f"K{row_index}", "values": [[normalized[regino]]]})
    batch_update_tracking(worksheet, updates)
    return len(updates)
