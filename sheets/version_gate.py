"""앱 버전 기록과 구버전 쓰기 차단.

여러 PC가 같은 공유 스프레드시트에 쓰므로, 쓰기 방식이 바뀐 뒤에도 업데이트하지 않은 PC가
예전 방식으로 메모·인덱스 등을 덮어쓸 수 있다. 이를 막기 위해:

- 공유 「설정」 탭의 ``min_app_version`` 값보다 이 앱의 APP_VERSION이 낮으면, Google 클라이언트의
  모든 쓰기 요청(GET이 아닌 요청)을 보내기 전에 AppVersionBlockedError로 막는다(읽기는 허용).
  최소 버전은 MIN_VERSION_RECHECK_SECONDS마다 다시 읽어, 실행 중에 다른 PC가 올린 값도 반영한다.
  최소 버전을 읽지 못하면(네트워크 오류 등) 쓰기를 막지 않는다.
- 시작할 때와 실행 중 1시간마다(startup_check) 공유 최소 버전이 이 앱의 MIN_COMPATIBLE_VERSION보다
  낮으면 그 값으로 올리고(낮추지는 않는다), 「PC버전」 탭에 PC·버전·마지막 확인 시각을 기록한다.
  다른 PC가 더 새 버전을 쓰고 있으면 그 버전을 알려 화면이 「새 버전 안내」를 띄우게 한다
  (닫기 / 오늘 더 이상 묻지 않음, 환경설정의 「새 버전 안내 무시」로 끌 수 있음).

릴리스 규칙: 배포할 때마다 APP_VERSION을 올린다. 공유 시트 쓰기 방식이 바뀌어 이전 버전과 함께
쓰면 데이터가 덮어써지는 변경이면 MIN_COMPATIBLE_VERSION도 그 버전으로 올린다. 그러면 새 버전
PC가 한 대라도 시작하는 순간 그보다 낮은 버전의 PC는 쓰기를 멈추고 업데이트 안내를 띄운다.
이 장치가 없는 이전 버전(이 파일이 생기기 전 버전)은 막을 수 없으므로 한 번은 직접 업데이트해야 한다.
"""

from __future__ import annotations

import threading
import time
from datetime import datetime

APP_VERSION = "2026.11.17.8"
MIN_COMPATIBLE_VERSION = "2026.11.17.1"

SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"
CONFIG_SHEET_TITLE = "설정"
CONFIG_SHEET_HEADERS = ["키", "값"]
CONFIG_KEY_MIN_APP_VERSION = "min_app_version"
PC_VERSION_SHEET_TITLE = "PC버전"
PC_VERSION_HEADERS = ["PC", "버전", "마지막 확인"]
VERSION_RECHECK_INTERVAL_MS = 60 * 60 * 1000  # 실행 중 새 버전·최소 버전 다시 확인(1시간)
# 이 PC의 database/app_settings.json에 저장하는 새 버전 안내 설정
SETTING_IGNORE_UPDATE_NOTICE = "ignore_version_update_notice"
SETTING_UPDATE_NOTICE_SNOOZED_DAY = "version_update_notice_snoozed_day"
MIN_VERSION_RECHECK_SECONDS = 10 * 60

_monotonic = time.monotonic
_lock = threading.Lock()
_local = threading.local()
_state = {"minimum": "", "checked_at": None}


class AppVersionBlockedError(RuntimeError):
    """공유 시트가 요구하는 최소 버전보다 낮은 앱이라 시트에 쓰지 않았다."""


def parse_version(text) -> tuple[int, ...] | None:
    parts = str(text or "").strip().split(".")
    if not parts or not all(part.isdigit() for part in parts):
        return None
    return tuple(int(part) for part in parts)


def is_outdated(minimum: str) -> bool:
    required = parse_version(minimum)
    return required is not None and parse_version(APP_VERSION) < required


def blocked_message(minimum: str) -> str:
    return (f"이 PC의 프로그램(버전 {APP_VERSION})이 공유 시트가 요구하는 최소 버전({minimum})보다 낮아 "
            "Google 시트에 쓰지 않았습니다. 다른 PC의 데이터를 예전 방식으로 덮어쓰지 않도록 막은 것입니다. "
            "프로그램을 최신 버전으로 업데이트한 뒤 다시 실행해 주세요.")


def reset_state() -> None:
    """기억한 최소 버전을 지운다(테스트용)."""
    with _lock:
        _state.update(minimum="", checked_at=None)


def _open_config_ws(gc):
    from tracking.repository import open_config_worksheet
    return open_config_worksheet(gc, SPREADSHEET_ID, CONFIG_SHEET_TITLE, CONFIG_SHEET_HEADERS)


def _read_min_version_text(gc) -> str:
    from tracking.repository import read_config_values_map
    return read_config_values_map(_open_config_ws(gc)).get(CONFIG_KEY_MIN_APP_VERSION, "")


def _current_minimum(gc) -> str:
    """기억한 최소 버전. 오래됐으면 다시 읽는다(실패하면 마지막 값 유지)."""
    with _lock:
        checked_at = _state["checked_at"]
        if checked_at is not None and _monotonic() - checked_at < MIN_VERSION_RECHECK_SECONDS:
            return _state["minimum"]
        _local.bypass = True  # 최소 버전 읽기 중의 요청은 다시 검사하지 않는다.
        try:
            _state["minimum"] = _read_min_version_text(gc)
        except Exception as error:
            print(f"! 공유 최소 버전 확인 실패(쓰기는 막지 않음): {error}")
        finally:
            _local.bypass = False
        _state["checked_at"] = _monotonic()
        return _state["minimum"]


def install_write_guard(gc):
    """gc의 모든 쓰기 요청 전에 최소 버전을 확인하도록 HTTP 요청 함수를 감싼다."""
    http = gc.http_client
    if getattr(http, "_easy_fulfill_write_guard", False):
        return gc
    original = http.request

    def guarded_request(method, endpoint, *args, **kwargs):
        if str(method).lower() != "get" and not getattr(_local, "bypass", False):
            minimum = _current_minimum(gc)
            if is_outdated(minimum):
                raise AppVersionBlockedError(blocked_message(minimum))
        return original(method, endpoint, *args, **kwargs)

    http.request = guarded_request
    http._easy_fulfill_write_guard = True
    return gc


def _open_pc_version_ws(gc):
    from sheets.worksheet_cache import open_cached

    def open_worksheet():
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        for worksheet in spreadsheet.worksheets():
            if worksheet.title == PC_VERSION_SHEET_TITLE:
                return worksheet
        worksheet = spreadsheet.add_worksheet(title=PC_VERSION_SHEET_TITLE, rows=50, cols=len(PC_VERSION_HEADERS))
        worksheet.update([PC_VERSION_HEADERS], range_name="A1", value_input_option="RAW")
        return worksheet

    return open_cached(gc, SPREADSHEET_ID, PC_VERSION_SHEET_TITLE, open_worksheet)


def startup_check(gc, actor: str, now: datetime | None = None) -> dict:
    """시작 시 최소 버전을 확인·상향하고 이 PC의 버전을 기록한다.

    반환: {"blocked": bool, "minimum": 공유 최소 버전, "version": APP_VERSION}
    """
    from tracking.repository import ConfigValueChangedError, config_values_map_from_rows, write_config_values

    config_ws = _open_config_ws(gc)
    minimum = config_values_map_from_rows(config_ws.get_all_values()).get(CONFIG_KEY_MIN_APP_VERSION, "")
    if is_outdated(minimum):
        with _lock:
            _state.update(minimum=minimum, checked_at=_monotonic())
        return {"blocked": True, "minimum": minimum, "version": APP_VERSION}

    current = parse_version(minimum)
    if current is None or current < parse_version(MIN_COMPATIBLE_VERSION):
        try:
            write_config_values(config_ws, {CONFIG_KEY_MIN_APP_VERSION: MIN_COMPATIBLE_VERSION},
                                expected={CONFIG_KEY_MIN_APP_VERSION: minimum})
            minimum = MIN_COMPATIBLE_VERSION
        except ConfigValueChangedError:
            pass  # 다른 PC가 먼저 바꿨다. 다음 확인 때 그 값을 읽는다.
    with _lock:
        _state.update(minimum=minimum, checked_at=_monotonic())

    stamp = (now or datetime.now()).strftime("%Y-%m-%d %H:%M:%S")
    pc_ws = _open_pc_version_ws(gc)
    rows = pc_ws.get_all_values()[1:]
    for row_number, row in enumerate(rows, start=2):
        if row and row[0].strip() == actor:
            pc_ws.update([[APP_VERSION, stamp]], range_name=f"B{row_number}:C{row_number}",
                         value_input_option="RAW")
            break
    else:
        pc_ws.append_rows([[actor, APP_VERSION, stamp]], value_input_option="RAW")
    return {"blocked": False, "minimum": minimum, "version": APP_VERSION,
            "latest": newest_other_version(rows, actor)}


def newest_other_version(rows, actor: str) -> str:
    """「PC버전」 탭 행 중 다른 PC가 쓰는, 이 앱보다 높은 가장 새 버전(없으면 빈 문자열)."""
    best_text, best = "", parse_version(APP_VERSION)
    for row in rows:
        if not row or row[0].strip() == actor or len(row) < 2:
            continue
        parsed = parse_version(row[1])
        if parsed is not None and parsed > best:
            best_text, best = row[1].strip(), parsed
    return best_text


def today_text() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def update_notice_due(latest: str, *, ignored: bool, snoozed_day: str, today: str) -> bool:
    """새 버전 안내를 띄울지: 새 버전이 있고, 안내 끄기가 아니며, 오늘 그만 묻기를 고르지 않았을 때."""
    return bool(latest) and not ignored and (snoozed_day or "") != today
