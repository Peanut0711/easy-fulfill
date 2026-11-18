"""공유 사은품 행사 설정과 주문별 마크다운·송장 표기."""

import json
import math
import re
from dataclasses import dataclass

from tracking.repository import ConfigValueChangedError, config_values_map_from_rows, write_config_values


GIFT_EVENT_KEY = "order_gift_event"
GIFT_EVENT_STORES = {
    "naver": "네이버", "coupang": "쿠팡", "gmarket": "지마켓", "11st": "11번가",
}
DEFAULT_INVOICE_LABELS = ((60_000, "펜/커터"), (100_000, "에코백"), (150_000, "3종세트"), (250_000, "4종세트"))


@dataclass(frozen=True)
class GiftEvent:
    enabled: bool
    stores: tuple[str, ...]
    tiers: tuple[tuple[int, str], ...]
    invoice_labels: tuple[tuple[int, str], ...] = ()

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
        for minimum, label in self.invoice_labels:
            if minimum not in amounts or not isinstance(label, str) or not label.strip() or label.splitlines() != [label]:
                raise ValueError("송장 표기를 해당 구간에 한 줄로 입력해 주세요.")
        object.__setattr__(self, "tiers", tuple(sorted(self.tiers)))
        object.__setattr__(self, "invoice_labels", tuple(sorted(self.invoice_labels)))

    def matching_tier(self, store: str, order_amount: float) -> tuple[int, str] | None:
        if not self.enabled or store not in self.stores or not math.isfinite(order_amount):
            return None
        for minimum, gift in reversed(self.tiers):
            if order_amount >= minimum:
                return minimum, gift
        return None

    def gift_name(self, store: str, order_amount: float) -> str:
        """주문 금액에 맞는 사은품 이름(해당 없으면 빈 문자열). 미리보기 표의 사은품 열에 쓴다."""
        tier = self.matching_tier(store, order_amount)
        return tier[1] if tier is not None else ""

    def markdown_suffix(self, store: str, order_amount: float) -> str:
        gift = self.gift_name(store, order_amount)
        if not gift:
            return ""
        safe_gift = re.sub(r"([\\`*_~\[\]<>!&|])", r"\\\1", gift)
        return f" · **🎁: {safe_gift}**"

    def invoice_prefix(self, store: str, order_amount: float) -> str:
        tier = self.matching_tier(store, order_amount)
        if tier is None:
            return ""
        label = dict(self.invoice_labels).get(tier[0], "")
        return f"[{label}] " if label else ""


DEFAULT_GIFT_EVENT = GiftEvent(False, tuple(GIFT_EVENT_STORES), (
    (60_000, "볼펜 또는 박스커터 (랜덤 1개)"),
    (100_000, "에코백"),
    (150_000, "볼펜 + 박스커터 + 에코백"),
    (250_000, "볼펜 + 박스커터 + 에코백 + USB-C 30W 충전기"),
), DEFAULT_INVOICE_LABELS)


def read_gift_event(worksheet) -> tuple[GiftEvent, str]:
    return gift_event_from_config_rows(worksheet.get_all_values())


def gift_event_from_config_rows(rows) -> tuple[GiftEvent, str]:
    """이미 읽은 공유 설정 시트 값(헤더 포함)에서 사은품 행사 설정을 해석한다."""
    raw = config_values_map_from_rows(rows).get(GIFT_EVENT_KEY, "")
    if not raw:
        return DEFAULT_GIFT_EVENT, raw
    try:
        data = json.loads(raw)
        event = GiftEvent(
            data["enabled"], tuple(data["stores"]), tuple(tuple(tier) for tier in data["tiers"]),
            tuple((int(minimum), label) for minimum, label in data.get("invoice_labels", {}).items()),
        )
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        raise ValueError("공유 시트의 사은품 행사 설정 형식을 확인해 주세요.") from error
    return event, raw


def save_gift_event(worksheet, event: GiftEvent, original: str) -> None:
    raw = json.dumps({
        "enabled": event.enabled, "stores": event.stores, "tiers": event.tiers,
        "invoice_labels": dict(event.invoice_labels),
    }, ensure_ascii=False)
    try:
        write_config_values(worksheet, {GIFT_EVENT_KEY: raw}, expected={GIFT_EVENT_KEY: original})
    except ConfigValueChangedError as error:
        raise ValueError("다른 작업자가 행사 설정을 변경했습니다. 공유 설정을 다시 연 뒤 수정해 주세요.") from error
