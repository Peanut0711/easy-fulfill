"""자주 여는 탭(송장추적·설정·문의알림)의 워크시트 속성을 잠시 기억해 메타데이터 요청을 줄인다.

gspread로 탭을 열면 open_by_key·worksheets()가 각각 스프레드시트 메타데이터를 받아,
5분마다 도는 문의 폴링만 해도 매번 4번의 메타데이터 요청이 든다. 탭 속성(제목·sheetId·
크기)은 거의 바뀌지 않으므로 처음 연 뒤 WORKSHEET_CACHE_TTL_SECONDS 동안은 기억한
속성으로 요청 없이 Worksheet를 만든다. 값 읽기·쓰기는 탭 제목으로 범위를 지정하므로 탭을
지우고 같은 이름으로 다시 만들어도 그대로 동작한다. 탭 이름을 바꾸거나 지워 호출이 실패하면
(SelfHealingWorksheet) 캐시를 지우고 탭을 다시 열어 그 호출을 바로 다시 실행한다.
"""

from __future__ import annotations

import threading
import time
from typing import Callable

WORKSHEET_CACHE_TTL_SECONDS = 30 * 60

_monotonic = time.monotonic
_lock = threading.Lock()
_entries: dict[tuple[str, str], tuple[float, dict]] = {}


# 탭이 없어졌을 때 Sheets API가 내는 400 오류 문구(제목 범위 / sheetId 요청)
_MISSING_TAB_MESSAGES = ("Unable to parse range", "No grid with id")


def is_missing_tab_error(error: BaseException) -> bool:
    """캐시한 탭이 이름 변경·삭제로 없어져 생긴 오류인지(400 + 해당 문구)."""
    response = getattr(error, "response", None)
    code = getattr(error, "code", None) or getattr(response, "status_code", None)
    return code == 400 and any(text in str(error) for text in _MISSING_TAB_MESSAGES)


class SelfHealingWorksheet:
    """캐시한 속성으로 만든 Worksheet를 감싼다.

    탭이 이름 변경·삭제로 없어져 호출이 실패하면, 캐시를 지우고 원래 열기 함수로 탭을 다시 찾거나
    만든 뒤(캐시 없이 열던 때와 같은 동작) 같은 호출을 한 번 다시 실행한다. 실패한 호출은 시트에
    반영되지 않은 상태라 다시 실행해도 중복 쓰기가 생기지 않는다. 다른 오류는 그대로 낸다.
    """

    __slots__ = ("_worksheet", "_reopen")

    def __init__(self, worksheet, reopen: Callable[[], object]):
        object.__setattr__(self, "_worksheet", worksheet)
        object.__setattr__(self, "_reopen", reopen)

    def __getattr__(self, name):
        value = getattr(self._worksheet, name)
        if not callable(value):
            return value

        def call(*args, **kwargs):
            try:
                return getattr(self._worksheet, name)(*args, **kwargs)
            except Exception as error:
                if not is_missing_tab_error(error):
                    raise
                fresh = self._reopen()
                object.__setattr__(self, "_worksheet", fresh)
                return getattr(fresh, name)(*args, **kwargs)

        return call

    def __setattr__(self, name, value):
        setattr(self._worksheet, name, value)

    def __repr__(self) -> str:
        return f"<SelfHealingWorksheet {self._worksheet!r}>"


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
        worksheet = worksheet_from_properties(gc, spreadsheet_id, dict(entry[1]))

        def reopen():
            forget(spreadsheet_id, title)
            return open_cached(gc, spreadsheet_id, title, open_worksheet)

        return SelfHealingWorksheet(worksheet, reopen)

    worksheet = open_worksheet()
    properties = getattr(worksheet, "_properties", None)
    if isinstance(properties, dict) and properties.get("title") == title and "sheetId" in properties:
        with _lock:
            _entries[key] = (_monotonic() + WORKSHEET_CACHE_TTL_SECONDS, dict(properties))
    return worksheet
