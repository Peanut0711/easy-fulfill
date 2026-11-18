"""주문 마크다운의 사은품 작업 건을 일별로 집계한다."""

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from hashlib import sha256
import json
import re

from orders.bulk import consolidate_11st_orders, consolidate_gmarket_orders, work_order_amount
from orders.gift_event import GiftEvent


KST = timezone(timedelta(hours=9))
DELIVERY_FIELDS = {
    "naver": ("수취인명", "통합배송지", "수취인연락처1", "수취인연락처1"),
    "gmarket": ("수령인명", "주소", "수령인 휴대폰", "수령인 전화번호"),
    "11st": ("수취인명", "주소", "휴대폰번호", "전화번호"),
}


@dataclass(frozen=True)
class GiftRecord:
    day: date
    store: str
    tier: int
    minimum: int
    gift: str
    key: str
    recorded_at: str


def _order_number(value) -> str:
    number = str(value).strip()
    # 숫자 열이 실수로 읽힌 주문 엑셀과 문자열을 돌려주는 API의 식별자를 맞춘다.
    return number[:-2] if re.fullmatch(r"[0-9]+\.0", number) else number


def _delivery_key_fields(store, source):
    """같은 원주문번호에도 배송지가 나뉘는 스토어의 작업 묶음을 구별한다."""
    if store not in DELIVERY_FIELDS:
        return ()
    name_col, address_col, mobile_col, phone_col = DELIVERY_FIELDS[store]
    name = " ".join(source.get(name_col, "").split())
    address = " ".join(source.get(address_col, "").split())
    phone = source.get(mobile_col) or source.get(phone_col, "")
    zipcode = _order_number(source.get("우편번호", ""))
    return name, address, zipcode.zfill(5) if zipcode.isdigit() else zipcode, re.sub(r"[\s-]", "", phone)


def build_gift_records(store: str, orders, event: GiftEvent, now: datetime) -> tuple[GiftRecord, ...]:
    """마크다운과 같은 묶음·금액으로 가장 높은 구간만 기록한다."""
    if not event.enabled or store not in event.stores:
        return ()
    consolidate = {"gmarket": consolidate_gmarket_orders, "11st": consolidate_11st_orders}.get(store)
    work_orders = consolidate(orders) if consolidate else orders
    tier_numbers = {minimum: number for number, (minimum, _) in enumerate(event.tiers, start=1)}
    stamp = now.astimezone(KST)
    day, recorded_at = stamp.date(), stamp.isoformat(timespec="milliseconds")
    records = []
    for order_key, info in work_orders.items():
        tier = event.matching_tier(store, work_order_amount(store, info))
        if tier is None:
            continue
        minimum, gift = tier
        source_ids = info.get("주문번호목록") or [info.get("원주문번호", order_key)]
        delivery = _delivery_key_fields(store, orders[order_key])
        identity = json.dumps([store, sorted({_order_number(value) for value in source_ids}), delivery],
                              ensure_ascii=False)
        key = sha256(identity.encode("utf-8")).hexdigest()
        records.append(GiftRecord(day, store, tier_numbers[minimum], minimum, gift, key, recorded_at))
    return tuple(records)


def summarize_gift_records(records, start: date, end: date, store: str = ""):
    """조회 범위의 전체·일별 (구간, 최소금액, 사은품명) 수량을 센다."""
    totals = Counter()
    daily = defaultdict(Counter)
    for record in records:
        if start <= record.day <= end and (not store or record.store == store):
            gift = (record.tier, record.minimum, record.gift)
            totals[gift] += 1
            daily[record.day][gift] += 1
    return totals, sorted(daily.items(), reverse=True)
