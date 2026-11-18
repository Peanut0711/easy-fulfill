"""앱 버전 기록과 구버전 쓰기 차단.

여러 PC가 같은 공유 스프레드시트에 쓰므로, 쓰기 방식이 바뀐 뒤에도 업데이트하지 않은 PC가
예전 방식으로 메모·인덱스 등을 덮어쓸 수 있다. 이를 막기 위해:

- 공유 「설정」 탭의 ``min_app_version`` 값보다 이 앱의 APP_VERSION이 낮으면, Google 클라이언트의
  모든 쓰기 요청(GET이 아닌 요청)을 보내기 전에 AppVersionBlockedError로 막는다(읽기는 허용).
  최소 버전은 MIN_VERSION_RECHECK_SECONDS마다 다시 읽어, 실행 중에 다른 PC가 올린 값도 반영한다.
  최소 버전을 다시 읽지 못하면 마지막 확인값을 유지한다. 확인한 값이 없으면 쓰기를 허용한다.
- 시작할 때와 실행 중 1시간마다(startup_check) 공유 최소 버전이 이 앱의 MIN_COMPATIBLE_VERSION보다
  낮으면 그 값으로 올리고(낮추지는 않는다), 「PC버전」 탭에 PC·버전·마지막 확인 시각을 기록한다.
  다른 PC가 더 새 버전을 쓰고 있으면 그 버전을 알려 화면이 「새 버전 안내」를 띄우게 한다
  (닫기 / 오늘 더 이상 묻지 않음, 설정 › 프로그램의 「새 버전 안내 무시」로 끌 수 있음).

릴리스 규칙: 배포할 때마다 APP_VERSION을 올린다. 공유 시트 쓰기 방식이 바뀌어 이전 버전과 함께
쓰면 데이터가 덮어써지는 변경이면 MIN_COMPATIBLE_VERSION도 그 버전으로 올린다. 그러면 새 버전
PC가 한 대라도 시작하는 순간 그보다 낮은 버전의 PC는 쓰기를 멈추고 업데이트 안내를 띄운다.
이 장치가 없는 이전 버전(이 파일이 생기기 전 버전)은 막을 수 없으므로 한 번은 직접 업데이트해야 한다.
"""

from __future__ import annotations

import threading
import time
from datetime import datetime

APP_VERSION = "0.0.25"
MIN_COMPATIBLE_VERSION = "2026.11.17.1"
# 기존 앱도 min_app_version의 숫자 비교만으로 향후 호환성 변경을 감지할 수 있게 한다.
_VERSION_RESET_PREFIX = (2026, 11, 18, 1)

SOURCETREE_UPDATE_STEPS = (
    ("프로그램 종료", "작업을 마치고 Easy Fulfill을 닫은 뒤 창이 사라질 때까지 기다리세요."),
    ("소스트리(SourceTree)", "easy-fulfill 저장소를 열고 왼쪽 브랜치에서 main을 선택하세요."),
    ("페치(Fetch)", "상단에서 클릭하고 원격 origin을 선택한 뒤 완료를 기다리세요."),
    ("풀(Pull)", "상단에서 클릭하고 원격 origin과 브랜치 main을 선택한 뒤 확인을 누르세요."),
    ("run.bat 실행", "같은 저장소 폴더에서 실행하고 타이틀바의 버전을 확인하세요."),
)
SOURCETREE_UPDATE_ERROR = "로컬 변경이나 충돌로 멈추면 오류 화면을 개발 담당자에게 보내 주세요."

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
    if not all(part.isascii() and part.isdecimal() for part in parts):
        return None
    return tuple(int(part) for part in parts)


def _version_order(text: str) -> tuple[int, ...]:
    parts = parse_version(text) or ()
    return (*_VERSION_RESET_PREFIX, *parts) if len(parts) == 3 else parts


def _minimum_display(minimum: str) -> str:
    parts = parse_version(minimum)
    prefix_length = len(_VERSION_RESET_PREFIX)
    if parts and len(parts) == prefix_length + 3 and parts[:prefix_length] == _VERSION_RESET_PREFIX:
        return ".".join(map(str, parts[prefix_length:]))
    return minimum


def is_outdated(minimum: str) -> bool:
    return _version_order(APP_VERSION) < _version_order(minimum)


def blocked_message(minimum: str) -> str:
    steps = "\n".join(
        f"{number}. {action} — {detail}"
        for number, (action, detail) in enumerate(SOURCETREE_UPDATE_STEPS, start=1)
    )
    return (
        f"이 PC의 버전 {APP_VERSION}은 공유 시트가 요구하는 최소 버전 {_minimum_display(minimum)}보다 낮습니다.\n"
        "Google 시트는 읽을 수 있습니다. 공유 데이터를 보호하기 위해 시트 쓰기는 업데이트 후 다시 시작됩니다.\n\n"
        f"{steps}\n\n{SOURCETREE_UPDATE_ERROR}"
    )


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
            print(f"! 공유 최소 버전 확인 실패(마지막 확인값 유지): {error}")
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
    with _lock:
        _state.update(minimum=minimum, checked_at=_monotonic())
    required = _version_order(MIN_COMPATIBLE_VERSION)
    if not is_outdated(minimum) and _version_order(minimum) < required:
        shared_minimum = ".".join(map(str, required))
        try:
            write_config_values(config_ws, {CONFIG_KEY_MIN_APP_VERSION: shared_minimum},
                                expected={CONFIG_KEY_MIN_APP_VERSION: minimum})
            minimum = shared_minimum
        except ConfigValueChangedError:
            minimum = config_values_map_from_rows(config_ws.get_all_values()).get(CONFIG_KEY_MIN_APP_VERSION, "")
        with _lock:
            _state.update(minimum=minimum, checked_at=_monotonic())
    if is_outdated(minimum):
        return {"blocked": True, "minimum": _minimum_display(minimum), "version": APP_VERSION}

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
    return {"blocked": False, "minimum": _minimum_display(minimum), "version": APP_VERSION,
            "latest": newest_other_version(rows, actor)}


def newest_other_version(rows, actor: str) -> str:
    """「PC버전」 탭 행 중 다른 PC가 쓰는, 이 앱보다 높은 가장 새 버전(없으면 빈 문자열)."""
    best_text, best = "", parse_version(APP_VERSION)
    for row in rows:
        if not row or row[0].strip() == actor or len(row) < 2:
            continue
        parsed = parse_version(row[1])
        # 날짜형 기록은 전환 전 PC의 실행 기록이다.
        if parsed is not None and len(parsed) == len(best) and parsed > best:
            best_text, best = row[1].strip(), parsed
    return best_text


def today_text() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def update_notice_due(latest: str, *, ignored: bool, snoozed_day: str, today: str) -> bool:
    """새 버전 안내를 띄울지: 새 버전이 있고, 안내 끄기가 아니며, 오늘 그만 묻기를 고르지 않았을 때."""
    return bool(latest) and not ignored and (snoozed_day or "") != today
