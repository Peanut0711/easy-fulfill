"""Qt UI와 외부 I/O에 의존하지 않는 배송추적 규칙."""

from __future__ import annotations

import hashlib
from datetime import datetime, timedelta
from typing import Mapping, Sequence


TRACKING_MANAGEMENT_COL = 12
TRACKING_MANAGEMENT_ACTIVE = "추적중"
TRACKING_MANAGEMENT_CANDIDATE = "폐기후보"
TRACKING_MANAGEMENT_DISCARDED = "폐기(미발송)"
TRACKING_MANAGEMENT_MANUAL_STOP = "수동 중지"
TRACKING_MANAGEMENT_EXCLUDED = {
    TRACKING_MANAGEMENT_CANDIDATE,
    TRACKING_MANAGEMENT_DISCARDED,
    TRACKING_MANAGEMENT_MANUAL_STOP,
}
TRACKING_COMPLETED_LOOKBACK_DAYS = 14

TRACKING_DISCARD_CANDIDATE_HOURS = 48
TRACKING_DISCARD_CONFIRM_HOURS = 72
TRACKING_DISCARD_STALE_ARRIVAL_LOCATION = "부평물류센터"
TRACKING_DISCARD_STALE_RECEIPT_LOCATION = "부평우체국"

CONFIG_KEY_STALE_HOURS = "stale_hours"
CONFIG_KEY_STALE_HUB = "stale_hub_hours"
CONFIG_KEY_STALE_PICKUP = "stale_pickup_hours"
CONFIG_KEY_STALE_TRANSIT = "stale_transit_hours"
STALE_DEFAULTS = {"허브정체": 12, "수거누락": 24, "이동정체": 48}
STALE_CONFIG_KEYS = {
    "허브정체": CONFIG_KEY_STALE_HUB,
    "수거누락": CONFIG_KEY_STALE_PICKUP,
    "이동정체": CONFIG_KEY_STALE_TRANSIT,
}
CONFIG_KEY_STALE_REMOTE_BONUS = "stale_remote_bonus_hours"
STALE_REMOTE_BONUS_DEFAULT = 24
CONFIG_KEY_INQUIRY_WORK_START = "inquiry_work_start_hour"
CONFIG_KEY_INQUIRY_WORK_END = "inquiry_work_end_hour"
NAVER_INQUIRY_WORK_START_HOUR = 10
NAVER_INQUIRY_WORK_END_HOUR = 19

RISK_HUB_KEYWORDS = ("물류", "허브", "터미널", "집중국", "교환센터")
RISK_EXCLUDE_STATUS = ("배달준비",)
RISK_REMOTE_KEYWORDS = ("제주", "서귀포", "울릉", "백령", "연평", "흑산", "추자", "거문")
RISK_CATEGORY_ORDER = {"허브정체": 0, "수거누락": 1, "이동정체": 2}
RISK_CATEGORY_LABELS = {
    "허브정체": "🔴 허브 정체(분실·사고 의심)",
    "수거누락": "🟠 수거 누락 의심",
    "이동정체": "🟠 이동 정체",
}


def is_weekday(now_dt: datetime | None = None) -> bool:
    """월~금이면 True, 토·일이면 False."""
    return (now_dt or datetime.now()).weekday() < 5


def business_elapsed_hours(ref_dt: datetime, now_dt: datetime) -> float:
    """두 시각 사이에서 토·일을 제외한 경과 시간을 시간 단위로 반환한다."""
    if now_dt <= ref_dt:
        return 0.0
    weekend = 0.0
    current = ref_dt
    while current < now_dt:
        day_end = datetime(current.year, current.month, current.day) + timedelta(days=1)
        segment_end = min(day_end, now_dt)
        if current.weekday() >= 5:
            weekend += (segment_end - current).total_seconds()
        current = segment_end
    return ((now_dt - ref_dt).total_seconds() - weekend) / 3600.0


def parse_timestamp(value: object) -> datetime | None:
    """`YYYY-MM-DD HH:MM:SS` 문자열을 datetime으로 파싱한다."""
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def parse_tracking_list_timestamp(value: object) -> datetime | None:
    """배송추적 목록의 등록·이벤트 시각 표기를 datetime으로 파싱한다."""
    text = str(value or "").strip()
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y.%m.%d %H:%M:%S",
        "%Y.%m.%d %H:%M",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def select_tracking_list_row_numbers(
    registrations: Sequence[object],
    completions: Sequence[object],
    events: Sequence[object],
    managements: Sequence[object],
    mode: str,
    now_dt: datetime | None = None,
    completed_lookback_days: int = TRACKING_COMPLETED_LOOKBACK_DAYS,
) -> list[int]:
    """목록 모드에 맞는 Sheet 행 번호만 골라 상세 행 읽기 범위를 줄인다."""
    now_dt = now_dt or datetime.now()
    today = now_dt.date()
    cutoff = today - timedelta(days=max(0, completed_lookback_days - 1))
    count = max(len(registrations), len(completions), len(events), len(managements))

    def cell(values: Sequence[object], index: int) -> str:
        return str(values[index] if index < len(values) else "" or "").strip()

    selected = []
    for index in range(count):
        registered = cell(registrations, index)
        complete = cell(completions, index).upper() == "Y"
        management = cell(managements, index)
        if mode == "오늘":
            include = registered.startswith(today.strftime("%Y-%m-%d"))
        elif mode == "완료":
            finished_at = parse_tracking_list_timestamp(cell(events, index))
            if finished_at is None:
                finished_at = parse_tracking_list_timestamp(registered)
            include = complete and finished_at is not None and finished_at.date() >= cutoff
        elif mode == "추적 중지":
            include = management == TRACKING_MANAGEMENT_MANUAL_STOP
        elif mode == "폐기 후보":
            include = management == TRACKING_MANAGEMENT_CANDIDATE
        elif mode == "폐기":
            include = management == TRACKING_MANAGEMENT_DISCARDED
        else:  # 배송중
            include = not complete and management not in TRACKING_MANAGEMENT_EXCLUDED
        if include:
            selected.append(index + 2)  # 헤더 다음 행부터 시작
    return selected


def tracking_management_state(
    row: Sequence[object], status: str, where: str = "", error_code: str = "",
    now_dt: datetime | None = None,
) -> str:
    """출력 뒤 미발송 또는 지정 거점 정체 송장의 자동 관리상태를 계산한다."""
    current = str(row[TRACKING_MANAGEMENT_COL] if len(row) > TRACKING_MANAGEMENT_COL else "").strip()
    if current == TRACKING_MANAGEMENT_MANUAL_STOP:
        return current
    eligible = (
        status in ("운송장출력", "추적정보 없음")
        or error_code == "ERR-125"
        or (status == "도착" and TRACKING_DISCARD_STALE_ARRIVAL_LOCATION in (where or ""))
        or (status == "인수완료" and TRACKING_DISCARD_STALE_RECEIPT_LOCATION in (where or ""))
    )
    if current == TRACKING_MANAGEMENT_DISCARDED:
        return current if eligible else TRACKING_MANAGEMENT_ACTIVE
    if not eligible:
        return TRACKING_MANAGEMENT_ACTIVE
    registered_at = parse_timestamp(row[1] if len(row) > 1 else "")
    if registered_at is None:
        return current or TRACKING_MANAGEMENT_ACTIVE
    elapsed_h = business_elapsed_hours(registered_at, now_dt or datetime.now())
    if elapsed_h >= TRACKING_DISCARD_CONFIRM_HOURS:
        return TRACKING_MANAGEMENT_DISCARDED
    if elapsed_h >= TRACKING_DISCARD_CANDIDATE_HOURS:
        return TRACKING_MANAGEMENT_CANDIDATE
    return TRACKING_MANAGEMENT_ACTIVE


def inquiry_alerts_allowed(
    now_dt: datetime, start_hour: int = NAVER_INQUIRY_WORK_START_HOUR,
    end_hour: int = NAVER_INQUIRY_WORK_END_HOUR,
) -> bool:
    """평일 근무시간(start_hour:00~end_hour:00) 안이면 True."""
    return is_weekday(now_dt) and start_hour <= now_dt.hour < end_hour


def read_inquiry_work_hours(cfg: Mapping[str, object]) -> tuple[int, int]:
    """설정 map에서 문의 알림 시간대를 읽고 유효하지 않으면 기본값을 쓴다."""
    def read_hour(key: str, default: int) -> int:
        try:
            value = int(str(cfg.get(key, "")).strip())
            return value if 0 <= value <= 23 else default
        except (TypeError, ValueError):
            return default

    start = read_hour(CONFIG_KEY_INQUIRY_WORK_START, NAVER_INQUIRY_WORK_START_HOUR)
    end = read_hour(CONFIG_KEY_INQUIRY_WORK_END, NAVER_INQUIRY_WORK_END_HOUR)
    if start >= end:
        return NAVER_INQUIRY_WORK_START_HOUR, NAVER_INQUIRY_WORK_END_HOUR
    return start, end


def is_remote_location(where: str) -> bool:
    """마지막 위치가 도서산간 키워드를 포함하면 True."""
    return any(keyword in (where or "") for keyword in RISK_REMOTE_KEYWORDS)


def risk_bucket(status: str, where: str, done: bool, management: str = "") -> str | None:
    """정체 시간과 무관하게 배송 상태를 위험 분류로 나눈다."""
    if done or management in TRACKING_MANAGEMENT_EXCLUDED:
        return None
    status = (status or "").strip()
    if status in RISK_EXCLUDE_STATUS:
        return None
    if status == "운송장출력":
        return "수거누락"
    if any(keyword in (where or "") for keyword in RISK_HUB_KEYWORDS):
        return "허브정체"
    return "이동정체"


def effective_threshold(bucket: str, where: str, threshold_hours: int, remote_bonus_hours: int) -> int:
    """분류 기준에 이동정체 도서산간 가산을 반영한다."""
    if bucket == "이동정체" and is_remote_location(where):
        return threshold_hours + remote_bonus_hours
    return threshold_hours


def evaluate_risk(
    status: str, where: str, done: bool, ref_dt: datetime | None, now_dt: datetime,
    threshold_hours: Mapping[str, int], remote_bonus_hours: int, management: str = "",
) -> tuple[str | None, float]:
    """위험 분류와 영업일 기준 정체 시간을 결합해 위험 여부를 반환한다."""
    bucket = risk_bucket(status, where, done, management)
    if bucket is None or ref_dt is None:
        return None, 0.0
    elapsed_h = business_elapsed_hours(ref_dt, now_dt)
    threshold = effective_threshold(
        bucket, where, int(threshold_hours.get(bucket, STALE_DEFAULTS[bucket])), remote_bonus_hours,
    )
    if elapsed_h > threshold:
        return bucket, elapsed_h
    return None, elapsed_h


def risk_signature(risks: Sequence[Mapping[str, object]]) -> str:
    """위험 목록 내용의 서명을 계산해 같은 목록의 재발송을 막는다."""
    keys = sorted(
        f"{item['regino']}|{item['status']}|{item['where']}|{item['event_time']}"
        for item in risks
    )
    return hashlib.md5("\n".join(keys).encode("utf-8")).hexdigest()


def build_risk_digest_text(
    risks: Sequence[Mapping[str, object]], threshold_hours: Mapping[str, int],
    now_dt: datetime | None = None,
) -> str:
    """위험 목록을 Slack 일일 다이제스트 본문으로 변환한다."""
    timestamp = (now_dt or datetime.now()).strftime("%Y-%m-%d %H:%M")
    lines = [f"⚠️ 배송 위험 {len(risks)}건 (평일 기준 무이동 · {timestamp})"]
    current_category = None
    shown = 0
    for item in risks:
        if shown >= 25:
            lines.append(f"… 외 {len(risks) - shown}건")
            break
        category = str(item["category"])
        if category != current_category:
            current_category = category
            threshold = threshold_hours.get(category, STALE_DEFAULTS.get(category, 0))
            lines.append(f"[{RISK_CATEGORY_LABELS.get(category, category)} · 기준 {threshold}h+]")
        elapsed = f"{float(item['elapsed_h']):.0f}시간째"
        where = str(item["where"] or "위치미상")
        remote_tag = " 🏝️도서산간" if category == "이동정체" and is_remote_location(where) else ""
        event_time = str(item["event_time"] or "")
        last = f", 마지막 {event_time}" if event_time else ""
        lines.append(
            f"• {item['regino']} {item['name']} — {item['status']} @ {where}{remote_tag} "
            f"({elapsed} 무이동{last})"
        )
        shown += 1
    return "\n".join(lines)
