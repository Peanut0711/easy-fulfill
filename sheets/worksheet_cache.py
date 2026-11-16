"""자주 여는 탭(송장추적·설정·문의알림)의 워크시트 속성을 잠시 기억해 메타데이터 요청을 줄인다.

gspread로 탭을 열면 open_by_key·worksheets()가 각각 스프레드시트 메타데이터를 받아,
5분마다 도는 문의 폴링만 해도 매번 4번의 메타데이터 요청이 든다. 탭 속성(제목·sheetId·
크기)은 거의 바뀌지 않으므로 처음 연 뒤 WORKSHEET_CACHE_TTL_SECONDS 동안은 기억한
속성으로 요청 없이 Worksheet를 만든다. 값 읽기·쓰기는 탭 제목으로 범위를 지정하므로 탭을
지우고 같은 이름으로 다시 만들어도 그대로 동작한다. 탭 이름을 바꾸거나 지운 경우에는 캐시가
만료될 때까지(최대 TTL) 그 탭 작업이 오류를 낸 뒤 다시 메타데이터를 받아 회복한다.
"""

from __future__ import annotations

import threading
import time
from typing import Callable

WORKSHEET_CACHE_TTL_SECONDS = 30 * 60

_monotonic = time.monotonic
_lock = threading.Lock()
_entries: dict[tuple[str, str], tuple[float, dict]] = {}


def clear() -> None:
    """기억한 탭 속성을 모두 지운다."""
    with _lock:
        _entries.clear()


def forget(spreadsheet_id: str, title: str) -> None:
    """한 탭의 기억한 속성을 지운다."""
    with _lock:
        _entries.pop((spreadsheet_id, title), None)


def open_cached(gc, spreadsheet_id: str, title: str, open_worksheet: Callable[[], object]):
    """기억한 속성이 있으면 요청 없이 Worksheet를 만들고, 없으면 open_worksheet()로 연 뒤 기억한다.

    open_worksheet는 탭을 찾거나 만들고 헤더를 맞추는 원래 열기 함수다(첫 열기에서만 실행).
    gspread Worksheet가 아닌 객체(속성 정보가 없는 객체)는 기억하지 않는다.
    """
    key = (spreadsheet_id, title)
    with _lock:
        entry = _entries.get(key)
        if entry is not None and entry[0] <= _monotonic():
            del _entries[key]
            entry = None
    if entry is not None:
        from sheets.batch_read import worksheet_from_properties
        return worksheet_from_properties(gc, spreadsheet_id, dict(entry[1]))

    worksheet = open_worksheet()
    properties = getattr(worksheet, "_properties", None)
    if isinstance(properties, dict) and properties.get("title") == title and "sheetId" in properties:
        with _lock:
            _entries[key] = (_monotonic() + WORKSHEET_CACHE_TTL_SECONDS, dict(properties))
    return worksheet
