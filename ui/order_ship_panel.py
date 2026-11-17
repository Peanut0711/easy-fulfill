"""주문·발송 탭 패널의 진행 단계와 버튼 강조.

작업 순서: 주문 불러오기 → 우체국 접수 엑셀(Ctrl+E, 하이제니스 폼) → 송장 불러오기(일괄발송 파일 자동 생성)
→ 택배사 접수목록. 다음에 할 동작 하나만 강조색 버튼으로 보여 준다.
"""

from __future__ import annotations

STAGE_ORDER = "order"        # 주문을 불러올 차례
STAGE_INTAKE = "intake"      # 우체국 접수 엑셀을 만들 차례
STAGE_INVOICE = "invoice"    # 송장을 불러올 차례
STAGE_RECEIPT = "receipt"    # 발송 파일까지 만들었으니 접수목록을 내보낼 차례

PRIMARY_BUTTON_STYLE = (
    "QPushButton { background: #1a73e8; color: white; font-weight: 600;"
    " border: 1px solid #1765cc; border-radius: 4px; padding: 4px 10px; }"
    "QPushButton:hover { background: #1765cc; }"
    "QPushButton:pressed { background: #155ab6; }"
    "QPushButton:disabled { background: #9fbff0; border-color: #9fbff0; }"
)


def next_stage(*, has_orders: bool, intake_done: bool, batch_done: bool) -> str:
    if batch_done:
        return STAGE_RECEIPT
    if not has_orders:
        return STAGE_ORDER
    if not intake_done:
        return STAGE_INTAKE
    return STAGE_INVOICE


def apply_primary(buttons: dict, stage: str) -> None:
    """buttons: {단계: QPushButton}. 해당 단계 버튼만 강조하고 나머지는 기본 모양으로 둔다."""
    for key, button in buttons.items():
        if button is None:
            continue
        primary = key == stage
        button.setStyleSheet(PRIMARY_BUTTON_STYLE if primary else "")
        button.setDefault(False)
        button.setProperty("primaryAction", primary)
