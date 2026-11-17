"""주문·발송 탭 패널의 진행 단계와 버튼 강조.

작업 순서: 주문 불러오기 → 우체국 접수 엑셀(Ctrl+E, 하이제니스 폼) → 송장 불러오기(일괄발송 파일 자동 생성)
→ 택배사 접수목록. 다음에 할 동작 하나만 강조색 버튼으로 보여 준다.
"""

from __future__ import annotations

from html import escape
from pathlib import Path

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPen, QPixmap

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


STORE_NAMES = {"naver": "네이버", "coupang": "쿠팡", "gmarket": "지마켓", "11st": "11번가"}
STORE_LOGO_FILES = {"naver": "naver-logo.png", "coupang": "coupang-logo.png", "gmarket": "gmarket-logo.png"}
STORE_TEXT_COLORS = {"11st": "#f43142"}

DONE_COLOR = "#2e7d32"
WARN_COLOR = "#e65100"


def order_summary_html(store_type, counts) -> str:
    """① 제목 옆 요약. counts는 order_preview.summary_counts 결과."""
    if not counts or not counts.get("orders"):
        return ""
    name = STORE_NAMES.get(store_type or "", "주문")
    text = f"✓ {escape(name)} {counts['orders']}건 · 사은품 {counts['gifts']}건"
    html = f'<span style="color:{DONE_COLOR}; font-weight:600;">{text}</span>'
    if counts.get("code_missing"):
        html += f' <span style="color:{WARN_COLOR};">· 코드 없음 {counts["code_missing"]}건</span>'
    return html


def ship_summary_html(*, batch=None, invoice_rows=0) -> str:
    """② 제목 옆 요약. batch=(스토어, 주문 건수, 매칭 수, 송장 수)면 발송 파일까지 만든 묶음."""
    if batch:
        store_type, orders, matched, total = batch
        name = STORE_NAMES.get(store_type or "", "주문")
        head = f'<span style="color:{DONE_COLOR}; font-weight:600;">✓ {escape(name)} {orders}건</span>'
        if total is None:
            return head
        color = DONE_COLOR if matched == total else WARN_COLOR
        return head + f' <span style="color:{color}; font-weight:600;">· 송장 매칭 {matched}/{total}</span>'
    if invoice_rows:
        return f'<span style="color:#555555;">송장 {invoice_rows}건</span>'
    return ""


def _logo(path: Path):
    pixmap = QPixmap(str(path))
    return None if pixmap.isNull() else pixmap


def store_badge_pixmap(store_type, image_dir, size, dpr=1.0):
    """스토어 로고 배지. 로고 파일이 없으면(11번가 등) 글자 배지로 그린다. 미선택이면 로고를 흐리게 모아 그린다."""
    width, height = int(size[0]), int(size[1])
    canvas = QPixmap(int(width * dpr), int(height * dpr))
    canvas.setDevicePixelRatio(dpr)
    canvas.fill(Qt.GlobalColor.transparent)
    painter = QPainter(canvas)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
    image_dir = Path(image_dir)

    def draw_logo(pixmap, rect: QRectF):
        scale = min(rect.width() / pixmap.width(), rect.height() / pixmap.height())
        w, h = pixmap.width() * scale, pixmap.height() * scale
        target = QRectF(rect.x() + (rect.width() - w) / 2, rect.y() + (rect.height() - h) / 2, w, h)
        painter.drawPixmap(target, pixmap, QRectF(pixmap.rect()))

    def draw_text(store, rect: QRectF, point_size):
        font = QFont()
        font.setBold(True)
        font.setPointSizeF(point_size)
        painter.setFont(font)
        painter.setPen(QColor(STORE_TEXT_COLORS.get(store, "#1a73e8")))
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, STORE_NAMES.get(store, store))

    if store_type in STORE_NAMES:
        painter.setPen(QPen(QColor("#dadce0"), 1))
        painter.setBrush(QColor("#ffffff"))
        painter.drawRoundedRect(QRectF(0.5, 0.5, width - 1, height - 1), 6, 6)
        inner = QRectF(10, 7, width - 20, height - 14)
        logo_file = STORE_LOGO_FILES.get(store_type)
        logo = _logo(image_dir / logo_file) if logo_file else None
        if logo is not None:
            draw_logo(logo, inner)
        else:
            draw_text(store_type, inner, 12)
    else:
        painter.setOpacity(0.35)
        cols, rows = 2, 2
        cell_w, cell_h = width / cols, height / rows
        for i, store in enumerate(("naver", "coupang", "gmarket", "11st")):
            rect = QRectF((i % cols) * cell_w + 4, (i // cols) * cell_h + 3, cell_w - 8, cell_h - 6)
            logo_file = STORE_LOGO_FILES.get(store)
            logo = _logo(image_dir / logo_file) if logo_file else None
            if logo is not None:
                draw_logo(logo, rect)
            else:
                draw_text(store, rect, 8)
    painter.end()
    return canvas


def next_stage(*, has_orders: bool, intake_done: bool, batch_done: bool) -> str:
    if batch_done:
        return STAGE_RECEIPT
    if not has_orders:
        return STAGE_ORDER
    if not intake_done:
        return STAGE_INTAKE
    return STAGE_INVOICE


def stabilize_widths(buttons) -> None:
    """강조 모양(굵은 글씨·여백)으로 바뀌어도 버튼 폭이 달라져 줄이 흔들리지 않게 두 모양 중 넓은 폭을 최소 폭으로 둔다."""
    for button in buttons:
        if button is None:
            continue
        button.setStyleSheet(PRIMARY_BUTTON_STYLE)
        primary_width = button.sizeHint().width()
        button.setStyleSheet("")
        button.setMinimumWidth(max(primary_width, button.sizeHint().width()))


def apply_primary(buttons: dict, stage: str) -> None:
    """buttons: {단계: QPushButton}. 해당 단계 버튼만 강조하고 나머지는 기본 모양으로 둔다."""
    for key, button in buttons.items():
        if button is None:
            continue
        primary = key == stage
        button.setStyleSheet(PRIMARY_BUTTON_STYLE if primary else "")
        button.setDefault(False)
        button.setProperty("primaryAction", primary)
