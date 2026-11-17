"""공유 사은품 행사 설정과 주문별 마크다운·송장 표기."""

import json
import math
import re
from dataclasses import dataclass

from tracking.repository import ConfigValueChangedError, config_values_map_from_rows, write_config_values


GIFT_EVENT_KEY = "order_gift_event"
GIFT_MARKDOWN_KEY = "order_gift_markdown"
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
    markdown_labels: tuple[tuple[int, str], ...] = ()

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
        for minimum, label in self.markdown_labels:
            if minimum not in amounts or not isinstance(label, str) or (
                label and (not label.strip() or label.splitlines() != [label])
            ):
                raise ValueError("마크다운 문구를 해당 구간에 한 줄로 입력해 주세요.")
        object.__setattr__(self, "tiers", tuple(sorted(self.tiers)))
        object.__setattr__(self, "invoice_labels", tuple(sorted(self.invoice_labels)))
        object.__setattr__(self, "markdown_labels", tuple(sorted(self.markdown_labels)))

    def _matching_tier(self, store: str, order_amount: float) -> tuple[int, str] | None:
        if not self.enabled or store not in self.stores or not math.isfinite(order_amount):
            return None
        for minimum, gift in reversed(self.tiers):
            if order_amount >= minimum:
                return minimum, gift
        return None

    def gift_name(self, store: str, order_amount: float) -> str:
        """주문 금액에 맞는 사은품 이름(해당 없으면 빈 문자열). 미리보기 표의 사은품 열에 쓴다."""
        tier = self._matching_tier(store, order_amount)
        return tier[1] if tier is not None else ""

    def markdown_suffix(self, store: str, order_amount: float) -> str:
        tier = self._matching_tier(store, order_amount)
        if tier is None:
            return ""
        minimum, gift = tier
        label = dict(self.markdown_labels).get(minimum, f"사은품: {gift}")
        safe_label = re.sub(r"([\\`*_~\[\]<>!&|])", r"\\\1", label)
        return f" · **🎁 {safe_label}**" if safe_label else " · **🎁**"

    def invoice_prefix(self, store: str, order_amount: float) -> str:
        tier = self._matching_tier(store, order_amount)
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


def read_gift_event(worksheet) -> tuple[GiftEvent, dict[str, str]]:
    return gift_event_from_config_rows(worksheet.get_all_values())


def gift_event_from_config_rows(rows) -> tuple[GiftEvent, dict[str, str]]:
    """이미 읽은 공유 설정 시트 값(헤더 포함)에서 사은품 행사 설정을 해석한다."""
    config = config_values_map_from_rows(rows)
    revision = {key: config.get(key, "") for key in (GIFT_EVENT_KEY, GIFT_MARKDOWN_KEY)}
    raw = revision[GIFT_EVENT_KEY]
    if not raw:
        return DEFAULT_GIFT_EVENT, revision
    try:
        data = json.loads(raw)
        tiers = tuple(tuple(tier) for tier in data["tiers"])
        active_amounts = {minimum for minimum, _ in tiers}
        markdown_labels = json.loads(revision[GIFT_MARKDOWN_KEY] or "{}")
        event = GiftEvent(
            data["enabled"], tuple(data["stores"]), tiers,
            tuple((int(minimum), label) for minimum, label in data.get("invoice_labels", {}).items()),
            # 구버전 앱에서 삭제한 금액의 문구는 건너뛴다.
            tuple((int(minimum), label) for minimum, label in markdown_labels.items()
                  if int(minimum) in active_amounts),
        )
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        raise ValueError("공유 시트의 사은품 행사 설정 형식을 확인해 주세요.") from error
    return event, revision


def save_gift_event(worksheet, event: GiftEvent, original: dict[str, str]) -> None:
    raw = json.dumps({
        "enabled": event.enabled, "stores": event.stores, "tiers": event.tiers,
        "invoice_labels": dict(event.invoice_labels),
    }, ensure_ascii=False)
    updates = {}
    if raw != original[GIFT_EVENT_KEY]:
        updates[GIFT_EVENT_KEY] = raw
    # 기존 앱이 행사 설정을 저장해도 마크다운 문구는 별도 키에 남는다.
    if event.markdown_labels or original[GIFT_MARKDOWN_KEY]:
        markdown = json.dumps(dict(event.markdown_labels), ensure_ascii=False)
        if markdown != original[GIFT_MARKDOWN_KEY]:
            updates[GIFT_MARKDOWN_KEY] = markdown
    try:
        write_config_values(worksheet, updates, expected=original, atomic=True)
    except ConfigValueChangedError as error:
        raise ValueError("다른 작업자가 행사 설정을 변경했습니다. 공유 설정을 다시 연 뒤 수정해 주세요.") from error
