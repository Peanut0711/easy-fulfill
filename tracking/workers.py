"""배송추적 Google Sheets/KPOST 워커와 기존 비동기 스레드 계약."""

from __future__ import annotations

import time
from datetime import datetime

from PySide6.QtCore import QThread, Signal

from .repository import (
    batch_update_tracking,
    normalize_tracking_no as _normalize_tracking_no,
    open_config_worksheet,
    open_tracking_worksheet,
    read_config_values_map,
    read_tracking_values,
    update_tracking_cell,
    update_tracking_management,
    update_tracking_notes,
    upsert_tracking_records,
    write_config_values,
)
from .service import (
    CONFIG_KEY_INQUIRY_WORK_END,
    CONFIG_KEY_INQUIRY_WORK_START,
    CONFIG_KEY_STALE_HUB,
    CONFIG_KEY_STALE_PICKUP,
    CONFIG_KEY_STALE_REMOTE_BONUS,
    CONFIG_KEY_STALE_TRANSIT,
    TRACKING_MANAGEMENT_ACTIVE,
    TRACKING_MANAGEMENT_COL,
    TRACKING_MANAGEMENT_EXCLUDED,
    TRACKING_MANAGEMENT_MANUAL_STOP,
    parse_timestamp as _parse_ts,
    tracking_management_state as _tracking_management_state,
)

try:
    import gspread
except ImportError:
    gspread = None

SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"
TRACKING_SHEET_TITLE = "송장추적"
TRACKING_SHEET_HEADERS = [
    "등기번호", "등록일시", "스토어", "주문번호", "수취인명",
    "택배사코드", "배송상태", "완료여부", "마지막위치", "최근조회시각", "비고",
    "최근이벤트시각", "관리상태",
]
CONFIG_SHEET_TITLE = "설정"
CONFIG_SHEET_HEADERS = ["키", "값"]
CONFIG_KEY_KPOST_REGKEY = "kpost_regkey"
CONFIG_KEY_SLACK_WEBHOOK = "slack_webhook_url"
CONFIG_KEY_AUTO_REFRESH_MIN = "tracking_auto_refresh_min"
CONFIG_KEY_LAST_AUTO_REFRESH = "tracking_last_auto_refresh"
KPOST_PICKUP_HOUR = 18
KPOST_TRACKING_REQUEST_DELAY_SEC = 0.1


def _standalone_open_tracking_ws(gc):
    return open_tracking_worksheet(gc, SPREADSHEET_ID, TRACKING_SHEET_TITLE, TRACKING_SHEET_HEADERS)


def _standalone_open_config_ws(gc):
    return open_config_worksheet(gc, SPREADSHEET_ID, CONFIG_SHEET_TITLE, CONFIG_SHEET_HEADERS)


def _read_config_values_map(ws):
    return read_config_values_map(ws)


def _write_config_values(ws, updates):
    return write_config_values(ws, updates)


def run_tracking_upsert_worker(records):
    """배송추적 Sheet upsert를 repository에 위임하고 기존 payload를 반환한다."""
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as error:
        return {"ok": False, "error": str(error)}
    try:
        worksheet = _standalone_open_tracking_ws(get_authorized_gspread_client())
        result = upsert_tracking_records(
            worksheet, records, TRACKING_SHEET_HEADERS, TRACKING_MANAGEMENT_ACTIVE)
        return {"ok": True, **result}
    except Exception as error:
        return {"ok": False, "error": str(error)}


def _cell_date_tuple(s):
    """'2026.06.26 02:45' / '2026-06-26 15:48:07' 등에서 (연,월,일) 튜플을 추출."""
    s = (s or "").strip()
    if not s:
        return None
    datepart = s.split(" ")[0].replace(".", "-")
    parts = datepart.split("-")
    if len(parts) >= 3:
        try:
            return (int(parts[0]), int(parts[1]), int(parts[2]))
        except ValueError:
            return None
    return None


def run_tracking_refresh_worker(regkey, progress_cb=None, auto=False, interval_min=0,
                                scope="all"):
    """「송장추적」 시트의 미완료 행만 골라 우체국 종추적조회로 상태를 갱신합니다.
    progress_cb(done, total)이 주어지면 진행 상황을 보고합니다.
    auto=True(백그라운드 자동 새로고침)면 공유 「설정」의 마지막 자동조회 시각을 보고,
    간격(interval_min, 설정 탭 값이 있으면 그쪽 우선) 안이면 우체국을 부르지 않고 건너뜁니다
    → 여러 대가 켜져 있어도 먼저 도는 1대만 실제 조회(다중 PC 중복 조회 방지).
    반환 dict: ok, total, complete, progress, failed, checked, aborted — 또는 ok False, error.
    자동 스킵 시: ok True, skipped_recent True (disabled True 면 설정상 꺼짐).
    """
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as e:
        return {"ok": False, "error": str(e)}
    try:
        import kpost_tracker
    except ImportError as e:
        return {"ok": False, "error": f"kpost_tracker 모듈을 불러올 수 없습니다: {e}"}
    try:
        gc = get_authorized_gspread_client()
        # 로컬에 키가 없거나 자동 새로고침이면 공유 「설정」 탭을 읽는다.
        # (자동일 땐 regkey 폴백 외에 '마지막 자동조회 시각'으로 다중 PC 중복 조회를 막는다.)
        cfg = None
        cfg_ws = None
        if not regkey or auto:
            try:
                cfg_ws = _standalone_open_config_ws(gc)
                cfg = _read_config_values_map(cfg_ws)
            except Exception:
                cfg = None
        if not regkey and cfg is not None:
            regkey = cfg.get(CONFIG_KEY_KPOST_REGKEY, "") or regkey
        if not regkey:
            return {"ok": False, "error": "우체국 OpenAPI 인증키(regkey)가 없습니다. 「배송추적」 탭에서 키를 등록하거나 관리자에게 문의하세요."}
        if auto:
            # 설정 탭 값이 있으면 간격을 그걸로(운영 중 조정 즉시 반영).
            eff_interval = interval_min
            if cfg is not None:
                cv = (cfg.get(CONFIG_KEY_AUTO_REFRESH_MIN, "") or "").strip()
                if cv:
                    try:
                        eff_interval = int(cv)
                    except ValueError:
                        pass
            if eff_interval <= 0:
                return {"ok": True, "skipped_recent": True, "disabled": True}
            last = _parse_ts(cfg.get(CONFIG_KEY_LAST_AUTO_REFRESH, "")) if cfg is not None else None
            if last is not None and (datetime.now() - last).total_seconds() < eff_interval * 60:
                return {"ok": True, "skipped_recent": True}
            # 슬롯 선점: 지금 시각을 먼저 기록해 다른 PC가 같은 창에서 중복 조회하지 않게 한다.
            # (조회가 실패/중단돼도 갱신되므로, 부하 상황에서 여러 대가 재차 몰리지 않는다.)
            if cfg_ws is not None:
                try:
                    _write_config_values(cfg_ws, {
                        CONFIG_KEY_LAST_AUTO_REFRESH: datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                except Exception:
                    pass
        ws = _standalone_open_tracking_ws(gc)
        values = read_tracking_values(ws)
        if len(values) <= 1:
            return {"ok": True, "total": 0, "complete": 0, "progress": 0,
                    "failed": 0, "checked": 0, "aborted": False}
        # scope: active=화면의 추적중, exceptions=폐기 후보/확정 재확인, all=자동 갱신.
        # (어제 이전의 운송장출력은 수거 누락 후보이므로 제외하지 않고 조회한다.)
        now_dt = datetime.now()
        today_tuple = (now_dt.year, now_dt.month, now_dt.day)
        before_pickup = now_dt.hour < KPOST_PICKUP_HOUR
        active = []
        skipped_pre_pickup = 0
        for ridx, row in enumerate(values[1:], start=2):
            tno = (row[0] if len(row) > 0 else "").strip()
            if not tno:
                continue
            done = (row[7] if len(row) > 7 else "").strip().upper()
            if done == "Y":
                continue
            management = (row[TRACKING_MANAGEMENT_COL]
                          if len(row) > TRACKING_MANAGEMENT_COL else "").strip()
            excluded = management in TRACKING_MANAGEMENT_EXCLUDED
            if (scope == "active" and excluded) or (scope == "exceptions" and not excluded):
                continue
            status = (row[6] if len(row) > 6 else "").strip()
            if before_pickup and status == "운송장출력":
                item_d = (_cell_date_tuple(row[11] if len(row) > 11 else "")
                          or _cell_date_tuple(row[1] if len(row) > 1 else ""))
                if item_d == today_tuple:
                    skipped_pre_pickup += 1
                    continue
            active.append((ridx, tno, row))
        if not active:
            return {"ok": True, "total": 0, "complete": 0, "progress": 0,
                    "failed": 0, "checked": 0, "aborted": False,
                    "skipped": skipped_pre_pickup}
        now = now_dt.strftime("%Y-%m-%d %H:%M:%S")
        complete = progress = failed = 0
        batch_updates = []
        total_active = len(active)
        aborted = False
        checked = 0
        for done_i, (ridx, tno, row) in enumerate(active, start=1):
            if progress_cb is not None:
                try:
                    progress_cb(done_i, total_active)
                except Exception:
                    pass
            s = kpost_tracker.summarize_tracking(regkey, tno)
            checked += 1
            code = (s.get("error_code") or "").upper()
            if not s.get("ok"):
                # ERR-131(시스템 부하 차단): 더 호출하면 차단되므로 중단
                if code == "ERR-131":
                    aborted = True
                    batch_updates.append({
                        "range": f"J{ridx}:K{ridx}",
                        "values": [[now, "우체국 시스템 부하로 조회 중단(ERR-131)"]],
                    })
                    break
                # ERR-001(조회결과 없음): 아직 추적정보 없음 → 실패가 아님
                if code == "ERR-001":
                    progress += 1
                    management = _tracking_management_state(
                        row, "추적정보 없음", error_code=code, now_dt=now_dt)
                    batch_updates.append({
                        "range": f"G{ridx}:M{ridx}",
                        "values": [["추적정보 없음", "N", "", now, "", "", management]],
                    })
                    time.sleep(KPOST_TRACKING_REQUEST_DELAY_SEC)
                    continue
                management = _tracking_management_state(
                    row, (row[6] if len(row) > 6 else ""),
                    error_code=code, now_dt=now_dt)
                failed += 1
                batch_updates.append({
                    "range": f"J{ridx}:K{ridx}",
                    "values": [[now, s.get("error", "조회 실패")]],
                })
                if management != (row[TRACKING_MANAGEMENT_COL]
                                  if len(row) > TRACKING_MANAGEMENT_COL else ""):
                    batch_updates.append({"range": f"M{ridx}", "values": [[management]]})
                time.sleep(KPOST_TRACKING_REQUEST_DELAY_SEC)
                continue
            done_yn = "Y" if s.get("complete") else "N"
            if s.get("complete"):
                complete += 1
            else:
                progress += 1
            management = _tracking_management_state(
                row, s.get("status", ""), s.get("where", ""), now_dt=now_dt)
            # G:M = 배송상태, 완료여부, 마지막위치, 최근조회시각, 비고, 최근이벤트시각, 관리상태
            batch_updates.append({
                "range": f"G{ridx}:M{ridx}",
                "values": [[s.get("status", ""), done_yn, s.get("where", ""),
                            now, "", s.get("time", ""), management]],
            })
            time.sleep(KPOST_TRACKING_REQUEST_DELAY_SEC)
        if batch_updates:
            batch_update_tracking(ws, batch_updates)
        return {
            "ok": True,
            "total": total_active,
            "complete": complete,
            "progress": progress,
            "failed": failed,
            "checked": checked,
            "aborted": aborted,
            "skipped": skipped_pre_pickup,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_tracking_refresh_one_worker(regkey, regino):
    """단일 등기번호만 우체국으로 조회해 「송장추적」 시트의 해당 행만 갱신합니다.
    반환 dict: ok, regino, status, complete — 또는 ok False, error.
    """
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as e:
        return {"ok": False, "error": str(e)}
    try:
        import kpost_tracker
    except ImportError as e:
        return {"ok": False, "error": f"kpost_tracker 모듈을 불러올 수 없습니다: {e}"}
    regino = str(regino).strip()
    if not regino:
        return {"ok": False, "error": "등기번호가 비어 있습니다."}
    try:
        gc = get_authorized_gspread_client()
        if not regkey:
            try:
                cfg = _read_config_values_map(_standalone_open_config_ws(gc))
                regkey = cfg.get(CONFIG_KEY_KPOST_REGKEY, "") or regkey
            except Exception:
                pass
        if not regkey:
            return {"ok": False, "error": "우체국 OpenAPI 인증키가 없습니다."}
        ws = _standalone_open_tracking_ws(gc)
        values = read_tracking_values(ws)
        ridx = None
        tracking_row = []
        for i, row in enumerate(values[1:], start=2):
            if (row[0] if row else "").strip() == regino:
                ridx = i
                tracking_row = row
                break
        if ridx is None:
            return {"ok": False, "error": "시트에서 해당 등기번호를 찾지 못했습니다."}
        s = kpost_tracker.summarize_tracking(regkey, regino)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        code = (s.get("error_code") or "").upper()
        if not s.get("ok"):
            if code == "ERR-001":
                management = _tracking_management_state(
                    tracking_row, "추적정보 없음", error_code=code)
                batch_update_tracking(ws, [{"range": f"G{ridx}:M{ridx}",
                                             "values": [["추적정보 없음", "N", "", now, "", "", management]]}])
                return {"ok": True, "regino": regino, "status": "추적정보 없음", "complete": False}
            management = _tracking_management_state(
                tracking_row, (tracking_row[6] if len(tracking_row) > 6 else ""),
                error_code=code)
            batch_update_tracking(ws, [{"range": f"J{ridx}:K{ridx}",
                                        "values": [[now, s.get("error", "조회 실패")]]}])
            if management != (tracking_row[TRACKING_MANAGEMENT_COL]
                              if len(tracking_row) > TRACKING_MANAGEMENT_COL else ""):
                update_tracking_cell(ws, ridx, "M", management)
            return {"ok": True, "regino": regino, "status": "(조회 실패)",
                    "complete": False, "note": s.get("error", "")}
        done_yn = "Y" if s.get("complete") else "N"
        management = _tracking_management_state(
            tracking_row, s.get("status", ""), s.get("where", ""))
        batch_update_tracking(ws, [{"range": f"G{ridx}:M{ridx}",
                                    "values": [[s.get("status", ""), done_yn, s.get("where", ""),
                                                now, "", s.get("time", ""), management]]}])
        return {"ok": True, "regino": regino, "status": s.get("status", ""),
                "complete": s.get("complete", False)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_tracking_list_worker():
    """「송장추적」 시트 전체 값을 읽어 표시용으로 반환합니다.
    반환 dict: ok, values(헤더 포함 2차원 리스트) — 또는 ok False, error.
    """
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as e:
        return {"ok": False, "error": str(e)}
    try:
        gc = get_authorized_gspread_client()
        ws = _standalone_open_tracking_ws(gc)
        return {"ok": True, "values": read_tracking_values(ws)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_tracking_management_update_worker(reginos, management):
    """선택된 송장의 관리상태 갱신을 repository에 위임한다."""
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    if management not in (TRACKING_MANAGEMENT_ACTIVE, TRACKING_MANAGEMENT_MANUAL_STOP):
        return {"ok": False, "error": "지원하지 않는 관리상태입니다."}
    if not {_normalize_tracking_no(regino) for regino in reginos or []} - {""}:
        return {"ok": False, "error": "등기번호가 비어 있습니다."}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as error:
        return {"ok": False, "error": str(error)}
    try:
        worksheet = _standalone_open_tracking_ws(get_authorized_gspread_client())
        return {"ok": True, **update_tracking_management(worksheet, reginos, management)}
    except Exception as error:
        return {"ok": False, "error": str(error)}


def run_tracking_notes_update_worker(notes):
    """등기번호별 메모 K열 갱신을 repository에 위임한다."""
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    if not {_normalize_tracking_no(regino) for regino in (notes or {})} - {""}:
        return {"ok": True, "updated": 0}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as error:
        return {"ok": False, "error": str(error)}
    try:
        worksheet = _standalone_open_tracking_ws(get_authorized_gspread_client())
        return {"ok": True, "updated": update_tracking_notes(worksheet, notes)}
    except Exception as error:
        return {"ok": False, "error": str(error)}


def run_tracking_config_read_worker():
    """공유 「설정」 탭에서 회사 공통 우체국 regkey 를 읽습니다.
    반환 dict: ok, regkey — 또는 ok False, error.
    """
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as e:
        return {"ok": False, "error": str(e)}
    try:
        gc = get_authorized_gspread_client()
        ws = _standalone_open_config_ws(gc)
        cfg = _read_config_values_map(ws)
        return {
            "ok": True,
            "regkey": cfg.get(CONFIG_KEY_KPOST_REGKEY, ""),
            "slack_webhook": cfg.get(CONFIG_KEY_SLACK_WEBHOOK, ""),
            "stale_hub": cfg.get(CONFIG_KEY_STALE_HUB, ""),
            "stale_pickup": cfg.get(CONFIG_KEY_STALE_PICKUP, ""),
            "stale_transit": cfg.get(CONFIG_KEY_STALE_TRANSIT, ""),
            "stale_remote_bonus": cfg.get(CONFIG_KEY_STALE_REMOTE_BONUS, ""),
            "inquiry_work_start": cfg.get(CONFIG_KEY_INQUIRY_WORK_START, ""),
            "inquiry_work_end": cfg.get(CONFIG_KEY_INQUIRY_WORK_END, ""),
            "auto_refresh_min": cfg.get(CONFIG_KEY_AUTO_REFRESH_MIN, ""),
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_tracking_config_write_worker(updates):
    """공유 「설정」 탭에 {설정키: 값} 들을 upsert 합니다."""
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as e:
        return {"ok": False, "error": str(e)}
    try:
        gc = get_authorized_gspread_client()
        ws = _standalone_open_config_ws(gc)
        _write_config_values(ws, updates)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}



class TrackingUpsertThread(QThread):
    """송장 등기번호를 「송장추적」 시트에 백그라운드로 upsert."""

    result_ready = Signal(dict)

    def __init__(self, records, parent=None):
        super().__init__(parent)
        self._records = records

    def run(self):
        self.result_ready.emit(run_tracking_upsert_worker(self._records))


class TrackingRefreshThread(QThread):
    """「송장추적」 시트 미완료 행을 우체국 종추적조회로 조회·갱신(백그라운드)."""

    result_ready = Signal(dict)
    progress = Signal(int, int)  # (done, total)

    def __init__(self, regkey, parent=None, auto=False, interval_min=0, scope="all"):
        super().__init__(parent)
        self._regkey = regkey
        self._auto = auto
        self._interval_min = interval_min
        self._scope = scope

    def run(self):
        self.result_ready.emit(
            run_tracking_refresh_worker(
                self._regkey, progress_cb=self.progress.emit,
                auto=self._auto, interval_min=self._interval_min, scope=self._scope,
            )
        )


class TrackingListThread(QThread):
    """「송장추적」 시트 전체를 백그라운드로 읽어 표 표시용으로 반환합니다."""

    result_ready = Signal(dict)

    def run(self):
        self.result_ready.emit(run_tracking_list_worker())


class TrackingManagementUpdateThread(QThread):
    """선택 송장의 수동 추적 중지/재개를 백그라운드로 저장합니다."""

    result_ready = Signal(dict)

    def __init__(self, reginos, management, parent=None):
        super().__init__(parent)
        self._reginos = reginos
        self._management = management

    def run(self):
        self.result_ready.emit(
            run_tracking_management_update_worker(self._reginos, self._management))


class TrackingNotesUpdateThread(QThread):
    """배송추적 비고 편집값을 백그라운드로 저장합니다."""

    result_ready = Signal(dict)

    def __init__(self, notes, parent=None):
        super().__init__(parent)
        self._notes = notes

    def run(self):
        self.result_ready.emit(run_tracking_notes_update_worker(self._notes))


class CourierReceiptExportThread(QThread):
    """당일 송장추적 등록분을 택배사 제출용 xlsx 로 백그라운드 생성합니다."""

    result_ready = Signal(dict)

    def __init__(self, save_path, today_str, parent=None):
        super().__init__(parent)
        self._save_path = save_path
        self._today_str = today_str

    def run(self):
        self.result_ready.emit(
            run_courier_receipt_export_worker(self._save_path, self._today_str))


class TrackingRefreshOneThread(QThread):
    """단일 등기번호만 우체국으로 조회·갱신(백그라운드)."""

    result_ready = Signal(dict)

    def __init__(self, regkey, regino, parent=None):
        super().__init__(parent)
        self._regkey = regkey
        self._regino = regino

    def run(self):
        self.result_ready.emit(run_tracking_refresh_one_worker(self._regkey, self._regino))


class TrackingConfigReadThread(QThread):
    """공유 「설정」 탭에서 회사 공통 regkey 를 백그라운드로 읽습니다."""

    result_ready = Signal(dict)

    def run(self):
        self.result_ready.emit(run_tracking_config_read_worker())


class TrackingConfigWriteThread(QThread):
    """공유 「설정」 탭에 {설정키: 값} 들을 백그라운드로 저장합니다."""

    result_ready = Signal(dict)

    def __init__(self, updates, parent=None):
        super().__init__(parent)
        self._updates = updates

    def run(self):
        self.result_ready.emit(run_tracking_config_write_worker(self._updates))
