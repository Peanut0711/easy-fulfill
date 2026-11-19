"""주문·발송 탭 주문 미리보기 표.

스토어별 처리 함수가 마크다운(클립보드용)을 만들 때 같은 반복문에서 PreviewOrder를 모으고,
여기서 그 목록으로 펼칠 수 있는 표(QTreeView) 모델을 만든다. 마크다운 생성은 건드리지 않는다.
"""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import QEvent, QObject, Qt
from PySide6.QtGui import QBrush, QColor, QKeySequence, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import QAbstractItemView, QHeaderView, QTreeView

from orders.bundle import bundle_note, packed_quantity
from shared import theme

STANDARD_DELIVERY = "택배,등기,소포"
COLUMNS = ("번호", "수취인", "금액", "사은품", "배송방법", "상품 수")
COL_NUMBER, COL_RECIPIENT, COL_AMOUNT, COL_GIFT, COL_DELIVERY, COL_COUNT = range(len(COLUMNS))

CODE_MISSING_BACKGROUND = theme.ROW_DANGER_BG
NON_PARCEL_BACKGROUND = theme.ROW_WARN_BG
CODE_MISSING_TEXT = theme.DANGER


@dataclass(frozen=True)
class PreviewProduct:
    code: str
    quantity: int
    name: str
    option: str
    extra: str = ""


@dataclass(frozen=True)
class PreviewOrder:
    """작업 목록 한 건(마크다운의 [ ] 줄 하나)."""

    number: int
    recipient: str
    amount: float
    amount_text: str
    gift: str
    delivery: str
    products: tuple
    check_codes: bool  # 상품코드 매칭을 쓰는 스토어(네이버·쿠팡)인지

    @property
    def code_missing(self) -> bool:
        return self.check_codes and any(not str(p.code).strip() for p in self.products)

    @property
    def non_parcel(self) -> bool:
        return bool(self.delivery) and self.delivery != STANDARD_DELIVERY

    @property
    def delivery_text(self) -> str:
        return self.delivery if self.non_parcel else "택배"


def product_line(product: PreviewProduct, check_codes: bool) -> str:
    code = str(product.code).strip()
    head = f"[{code or '코드 없음'}] " if check_codes else ""
    note = bundle_note(product.quantity, product.option)
    count = f"{packed_quantity(product.quantity, product.option)}개({note})" if note else f"{product.quantity}개"
    line = f"{head}{count} · {product.name}"
    option = str(product.option).strip()
    if option and option != "없음":
        line += f" (옵션: {option})"
    extra = str(product.extra).strip()
    if extra:
        line += f" (추가구성: {extra})"
    return line


def order_flags_text(order: PreviewOrder) -> str:
    reasons = []
    if order.code_missing:
        reasons.append("상품코드 매칭 실패")
    if order.non_parcel:
        reasons.append(f"택배 외 배송: {order.delivery}")
    return " · ".join(reasons)


def summary_counts(orders) -> dict:
    orders = list(orders)
    return {
        "orders": len(orders),
        "gifts": sum(1 for o in orders if o.gift),
        "code_missing": sum(1 for o in orders if o.code_missing),
        "non_parcel": sum(1 for o in orders if o.non_parcel),
    }


def _item(text, *, align=None, tooltip=""):
    item = QStandardItem(str(text))
    item.setEditable(False)
    if align is not None:
        item.setTextAlignment(align)
    if tooltip:
        item.setToolTip(tooltip)
    return item


def build_preview_model(orders, parent=None) -> QStandardItemModel:
    model = QStandardItemModel(0, len(COLUMNS), parent)
    model.setHorizontalHeaderLabels(list(COLUMNS))
    right = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
    center = Qt.AlignmentFlag.AlignCenter
    for order in orders:
        flags = order_flags_text(order)
        total_qty = sum(packed_quantity(p.quantity, p.option) for p in order.products)
        cells = [
            _item(order.number, align=right),
            _item(order.recipient),
            _item(order.amount_text, align=right, tooltip=f"{order.amount:,.0f}원"),
            _item(order.gift, tooltip=order.gift),
            _item(order.delivery_text, align=center),
            _item(len(order.products), align=right, tooltip=f"수량 합계 {total_qty}개"),
        ]
        background = None
        if order.code_missing:
            background = QBrush(QColor(CODE_MISSING_BACKGROUND))
        elif order.non_parcel:
            background = QBrush(QColor(NON_PARCEL_BACKGROUND))
        for cell in cells:
            if background is not None:
                cell.setBackground(background)
            if flags:
                cell.setToolTip(flags if not cell.toolTip() else f"{cell.toolTip()}\n{flags}")
        for product in order.products:
            child = _item(product_line(product, order.check_codes))
            if order.check_codes and not str(product.code).strip():
                child.setForeground(QBrush(QColor(CODE_MISSING_TEXT)))
            cells[0].appendRow([child])
        model.appendRow(cells)
    return model


class _CopyKeyFilter(QObject):
    """표에 초점이 있을 때 Ctrl+C가 칸 하나가 아니라 주문 전체 마크다운을 복사하게 한다."""

    def __init__(self, view, copy_callback):
        super().__init__(view)
        self._copy = copy_callback

    def eventFilter(self, obj, event):
        kind = event.type()
        if kind in (QEvent.Type.ShortcutOverride, QEvent.Type.KeyPress) and event.matches(QKeySequence.StandardKey.Copy):
            event.accept()
            if kind == QEvent.Type.KeyPress:
                self._copy()
            return True
        return False


def setup_preview_view(view: QTreeView, copy_callback=None) -> None:
    view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    view.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    view.setUniformRowHeights(True)
    view.setAllColumnsShowFocus(True)
    view.setRootIsDecorated(True)
    view.setExpandsOnDoubleClick(True)
    view.setModel(build_preview_model([], view))
    _size_columns(view)
    if copy_callback is not None:
        view._order_copy_filter = _CopyKeyFilter(view, copy_callback)
        view.installEventFilter(view._order_copy_filter)


def _size_columns(view: QTreeView) -> None:
    header = view.header()
    header.setStretchLastSection(False)
    for col in range(len(COLUMNS)):
        mode = QHeaderView.ResizeMode.Stretch if col == COL_GIFT else QHeaderView.ResizeMode.ResizeToContents
        header.setSectionResizeMode(col, mode)
    header.setMinimumSectionSize(36)


def show_preview(view: QTreeView, orders) -> QStandardItemModel:
    old = view.model()
    model = build_preview_model(orders, view)
    view.setModel(model)
    if old is not None and old is not model:
        old.deleteLater()
    for row in range(model.rowCount()):
        parent = model.index(row, 0)
        for child in range(model.rowCount(parent)):
            view.setFirstColumnSpanned(child, parent, True)
    _size_columns(view)
    return model
