"""옵션명에 적힌 묶음 개수(예: 「3개」 옵션)를 읽어 포장할 낱개 수를 계산한다.

판매처 엑셀의 수량은 옵션을 몇 번 샀는지(세트 수)라서, 「3개」 옵션을 1번 산 주문은 수량 1로 온다.
포장 담당자가 마크다운의 수량만 보고 1개를 넣는 실수를 막으려고 옵션명의 묶음 개수를 곱해 보여 준다.
쿠팡에만 쓴다(옵션 하나가 묶음 하나인 「2개」·「5개」 옵션이 쿠팡에만 있다). 네이버·지마켓·11번가는
옵션명의 「N개」가 묶음이 아닌 경우(예: 수량이 이미 낱개 수인 11번가 「회전각:360도-12개」 옵션)가 있어
작업자 요청으로 주문 수량을 그대로 보여 준다.
"""

from __future__ import annotations

import re

# 숫자 바로 뒤의 「개」만 본다. 「12개월」은 기간이라 제외하고, 소수·천 단위 일부(1.5개, 1,000개)는 잡지 않는다.
_BUNDLE_PATTERN = re.compile(r"(?<![\d.,])(\d{1,4})\s*개(?!월)")


def bundle_size(option) -> int:
    """옵션명에서 묶음 개수를 읽는다. 없거나 1이거나 서로 다른 개수가 섞여 애매하면 1."""
    text = "" if option is None else str(option)
    sizes = {int(match) for match in _BUNDLE_PATTERN.findall(text)}
    if len(sizes) != 1:
        return 1
    size = sizes.pop()
    return size if size >= 2 else 1


def packed_quantity(quantity, option) -> int:
    """포장할 낱개 수(세트 수 × 묶음 개수)."""
    return int(quantity or 0) * bundle_size(option)


def bundle_note(quantity, option) -> str:
    """묶음 옵션이면 「1세트×3개」, 아니면 빈 문자열."""
    size = bundle_size(option)
    return f"{int(quantity or 0)}세트×{size}개" if size > 1 else ""


def quantity_markdown(quantity, option, *, bundles: bool = False) -> str:
    """작업 목록 마크다운의 수량 칸. 기본은 주문 수량만 적는다.

    bundles=True(쿠팡)면 묶음 옵션일 때 낱개 수와 계산식을 함께 적는다.
    """
    note = bundle_note(quantity, option) if bundles else ""
    if not note:
        return f"**[ {quantity} 개 ]**"
    return f"**[ {packed_quantity(quantity, option)} 개 ]** ({note})"
