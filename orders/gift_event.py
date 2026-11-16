"""공유 사은품 행사 설정과 주문별 마크다운 표기."""

import json
import math
import re
from dataclasses import dataclass

from tracking.repository import ConfigValueChangedError, read_config_values_map, write_config_values


GIFT_EVENT_KEY = "order_gift_event"
GIFT_EVENT_STORES = {
    "naver": "네이버", "coupang": "쿠팡", "gmarket": "지마켓", "11st": "11번가",
}


@dataclass(frozen=True)
class GiftEvent:
    enabled: bool
    stores: tuple[str, ...]
    tiers: tuple[tuple[int, str], ...]

    def __post_init__(self):
        if type(self.enabled) is not bool or any(store not in GIFT_EVENT_STORES for store in self.stores):
            raise ValueError("이벤트 사용 여부 또는 적용 스토어를 확인해 주세요.")
        amounts = set()
        for minimum, gift in self.tiers:
            if type(minimum) is not int or minimum <= 0:
                raise ValueError("최소 구매 금액은 1원 이상의 정수로 입력해 주세요.")
            if minimum in amounts:
                raise ValueError(f"최소 구매 금액 {minimum:,}원이 중복되어 있습니다.")
            if not isinstance(gift, str) or not gift.strip() or gift.splitlines() != [gift]:
                raise ValueError("각 구간의 지급 사은품을 한 줄로 입력해 주세요.")
            amounts.add(minimum)
        if self.enabled and (not self.stores or not self.tiers):
            raise ValueError("이벤트를 켜려면 적용 스토어와 사은품 구간을 추가해 주세요.")
        object.__setattr__(self, "tiers", tuple(sorted(self.tiers)))

    def markdown_suffix(self, store: str, order_amount: float) -> str:
        if not self.enabled or store not in self.stores or not math.isfinite(order_amount):
            return ""
        for minimum, gift in reversed(self.tiers):
            if order_amount >= minimum:
                safe_gift = re.sub(r"([\\`*_~\[\]<>!&|])", r"\\\1", gift)
                return f" · **🎁 사은품: {safe_gift}**"
        return ""


DEFAULT_GIFT_EVENT = GiftEvent(False, tuple(GIFT_EVENT_STORES), (
    (60_000, "볼펜 또는 박스커터 (랜덤 1개)"),
    (100_000, "에코백"),
    (150_000, "볼펜 + 박스커터 + 에코백"),
    (250_000, "볼펜 + 박스커터 + 에코백 + USB-C 30W 충전기"),
))


def read_gift_event(worksheet) -> tuple[GiftEvent, str]:
    raw = read_config_values_map(worksheet).get(GIFT_EVENT_KEY, "")
    if not raw:
        return DEFAULT_GIFT_EVENT, raw
    try:
        data = json.loads(raw)
        event = GiftEvent(data["enabled"], tuple(data["stores"]), tuple(tuple(tier) for tier in data["tiers"]))
    except (KeyError, TypeError, ValueError) as error:
        raise ValueError("공유 시트의 사은품 행사 설정 형식을 확인해 주세요.") from error
    return event, raw


def save_gift_event(worksheet, event: GiftEvent, original: str) -> None:
    raw = json.dumps({"enabled": event.enabled, "stores": event.stores, "tiers": event.tiers}, ensure_ascii=False)
    try:
        write_config_values(worksheet, {GIFT_EVENT_KEY: raw}, expected={GIFT_EVENT_KEY: original})
    except ConfigValueChangedError as error:
        raise ValueError("다른 작업자가 행사 설정을 변경했습니다. 공유 설정을 다시 연 뒤 수정해 주세요.") from error
