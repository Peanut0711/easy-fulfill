"""스프레드시트 메타데이터 1회와 값 일괄 조회 1회로 여러 워크시트를 읽는 도우미.

gspread의 open_by_key·worksheets()·get_all_values()를 탭마다 따로 부르면 탭 하나에
메타데이터 2회와 값 1회 요청이 든다. 같은 시점에 여러 탭이 필요할 때 이 모듈로
메타데이터를 한 번 받아 워크시트를 찾고, 필요한 탭 전체 값을 한 요청으로 받는다.
"""

from __future__ import annotations

from collections.abc import Sequence
from http import HTTPStatus

import gspread
from gspread.exceptions import APIError, SpreadsheetNotFound
from gspread.utils import absolute_range_name, fill_gaps


def fetch_sheet_properties(gc, spreadsheet_id: str) -> list[dict]:
    """워크시트 속성(제목·sheetId·순서)을 탭 순서대로 한 번의 요청으로 받는다.

    스프레드시트를 열 수 없을 때는 gc.open_by_key와 같은 예외를 낸다.
    """
    try:
        metadata = gc.http_client.fetch_sheet_metadata(spreadsheet_id)
    except APIError as error:
        if error.response.status_code == HTTPStatus.NOT_FOUND:
            raise SpreadsheetNotFound(error.response) from error
        if error.response.status_code == HTTPStatus.FORBIDDEN:
            raise PermissionError from error
        raise
    return [sheet["properties"] for sheet in metadata.get("sheets", [])]


def worksheet_from_properties(gc, spreadsheet_id: str, properties: dict) -> gspread.Worksheet:
    """이미 받은 속성으로 쓰기용 Worksheet를 만든다(추가 요청 없음)."""
    return gspread.Worksheet(None, properties, spreadsheet_id, gc.http_client)


def batch_get_all_values(gc, spreadsheet_id: str, titles: Sequence[str]) -> list[list[list[str]]]:
    """여러 탭의 전체 값을 한 요청으로 읽어 get_all_values()와 같은 형태로 반환한다."""
    if not titles:
        return []
    response = gc.http_client.values_batch_get(
        spreadsheet_id, [absolute_range_name(title) for title in titles])
    value_ranges = response.get("valueRanges", [])
    if len(value_ranges) != len(titles):
        raise ValueError(f"시트 값 일괄 조회 결과가 {len(titles)}개가 아닙니다: {len(value_ranges)}개")
    results = []
    for value_range in value_ranges:
        # Worksheet.get_all_values()와 같은 빈 값 처리와 행 길이 맞춤.
        values = value_range.get("values", [[]])
        try:
            values = fill_gaps(values)
        except KeyError:
            values = [[]]
        results.append(values)
    return results
