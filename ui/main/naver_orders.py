"""메인 창 「발주」 탭: 네이버 발송 전 주문(신규주문·발주확인)을 API로 조회해 목록과 내용을 보여 준다(조회 전용).

주문번호 인덱스·공유 시트에는 쓰지 않는다. 다음 단계(조회한 주문으로 바로 노션 정리·우체국 송장)를 위해
주문마다 엑셀 처리 경로와 같은 행(rows, naver_commerce.order_detail_to_row)을 함께 들고 있다.
"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from shared import theme
from shared.ui_components import EmptyState, align_numeric, fixed_width_button, setup_table
from shared.ui_style import set_button_role, set_status
from ui.main import host as _host
from ui.main.rich_cards import TONES, block, card, chip, esc, facts, muted, section

NAVER_ORDER_FILTERS = ("전체", "신규주문", "발주확인")
NAVER_ORDER_COLUMNS = ("구분", "결제", "수취인", "상품", "금액")
NAVER_ORDER_EMPTY_TEXT = "「새로고침」을 누르면 네이버 발송 전 주문(신규주문·발주확인)을 불러옵니다."
STAGE_TONES = {"신규주문": "warn", "발주확인": "info"}  # 신규주문 = 아직 발주확인 안 함(처리할 일)


def won(amount):
    return f"{int(amount or 0):,}원"


def naver_order_product_summary(order):
    """목록 「상품」 칸: 첫 상품 이름, 여럿이면 「외 N건」."""
    items = order.get("items") or []
    if not items:
        return ""
    first = " ".join(str(items[0].get("product") or "").split())
    return first + (f" 외 {len(items) - 1}건" if len(items) > 1 else "")


def naver_order_detail_html(order):
    """오른쪽 상세 HTML: 구분·클레임 배지, 수취인 제목, 발송기한, 상품 카드, 배송메모 카드, 받는 분 표."""
    stage = order.get("stage") or ""
    chips = [chip(stage, STAGE_TONES.get(stage, "muted"))]
    claims = sorted({i["claim"] for i in order.get("items") or [] if i.get("claim")})
    chips += [chip(c, "danger") for c in claims]
    if order.get("delivery_method") and order["delivery_method"] != "택배,등기,소포":
        chips.append(chip(order["delivery_method"], "muted"))
    meta = [f"결제 {esc(order.get('paid_at'))}"]
    if order.get("orderer"):
        meta.append(f"구매자 {esc(order['orderer'])}")
    out = [
        f'<p style="margin-bottom:{theme.SPACE_XS}px;">{"&nbsp;".join(chips)}</p>',
        f'<p style="margin-top:0; margin-bottom:{theme.SPACE_XS}px; font-size:{theme.FONT_SIZE + 3}pt; '
        f'font-weight:600; color:{theme.TEXT};">{esc(order.get("receiver")) or "(수취인 없음)"}</p>',
        f'<p style="margin-top:0;">{muted(" · ".join(meta))}</p>',
        facts([("주문번호", esc(order.get("order_id"))),
               ("발송기한", f'<b style="color:{TONES["warn"][0]};">{esc(order.get("shipping_due"))}</b>'
                if order.get("shipping_due") else "")]),
    ]
    items = order.get("items") or []
    out.append(section(f"상품 {len(items)}건", "primary",
                       note=f"수량 {order.get('total_quantity', 0)} · 합계 {won(order.get('total_amount'))}"))
    mixed = len({i.get("stage") for i in items}) > 1
    for item in items:
        lines = [f"<b>{esc(item.get('product')) or '(상품명 없음)'}</b>"]
        if item.get("option"):
            lines.append(muted(esc(item["option"])))
        lines.append(f"× {esc(item.get('quantity'))} · {won(item.get('amount'))}")
        item_chips = []
        if mixed:
            item_chips.append(chip(item.get("stage"), STAGE_TONES.get(item.get("stage"), "muted")))
        if item.get("claim"):
            item_chips.append(chip(item["claim"], "danger"))
        if item_chips:
            lines.append("&nbsp;".join(item_chips))
        tone = "danger" if item.get("claim") else STAGE_TONES.get(item.get("stage"), "muted")
        out.append(card("<br>".join(lines), tone, plain_bg=True))
    if order.get("memo"):
        out.append(section("배송메모", "warn"))
        out.append(card(block(order["memo"]), "warn"))
    zip_html = f"({esc(order.get('zip'))}) " if order.get("zip") else ""
    receiver = facts([("수령인", f"<b>{esc(order.get('receiver'))}</b>" if order.get("receiver") else ""),
                      ("연락처", esc(order.get("tel"))),
                      ("주소", f"{zip_html}{esc(order.get('address'))}" if order.get("address") else ""),
                      ("배송방법", esc(order.get("delivery_method")))])
    if receiver:
        out.append(section("받는 분"))
        out.append(receiver)
    return "".join(out)


class NaverOrdersMixin:
    """MainWindow 믹스인. 메인 창 「발주」 탭: 네이버 발송 전 주문 조회(읽기 전용)."""

    def _mount_naver_orders_tab(self):
        self._naver_orders = []  # 마지막으로 불러온 주문 전부(다음 단계에서 rows를 처리 경로로 넘긴다)
        self._naver_orders_shown = []  # 필터를 적용해 표에 보이는 주문
        self._naver_orders_thread = None
        self._naver_orders_loaded_once = False
        self._naver_orders_days = 0
        tabs = self.ui.tabWidget
        page = QWidget()
        page.setObjectName("tab_naver_orders")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
        layout.setSpacing(theme.SPACE_SM)

        bar = QHBoxLayout()
        bar.setSpacing(theme.SPACE_SM)
        self.comboBox_naver_order_filter = QComboBox()
        self.comboBox_naver_order_filter.setObjectName("comboBox_naver_order_filter")
        self.comboBox_naver_order_filter.addItems(list(NAVER_ORDER_FILTERS))
        self.label_naver_order_status = QLabel()
        self.label_naver_order_status.setObjectName("label_naver_order_status")
        self.label_naver_order_status.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.pushButton_naver_order_refresh = QPushButton("새로고침")
        self.pushButton_naver_order_refresh.setObjectName("pushButton_naver_order_refresh")
        set_button_role(self.pushButton_naver_order_refresh, "primary")
        bar.addWidget(self.comboBox_naver_order_filter)
        bar.addWidget(self.label_naver_order_status, 1)
        bar.addWidget(fixed_width_button(self.pushButton_naver_order_refresh))
        layout.addLayout(bar)

        self.stackedWidget_naver_orders = QStackedWidget()
        self.stackedWidget_naver_orders.setObjectName("stackedWidget_naver_orders")
        self.emptyState_naver_orders = EmptyState(NAVER_ORDER_EMPTY_TEXT)
        table = self.tableWidget_naver_orders = QTableWidget(0, len(NAVER_ORDER_COLUMNS))
        table.setObjectName("tableWidget_naver_orders")
        table.setHorizontalHeaderLabels(list(NAVER_ORDER_COLUMNS))
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        setup_table(table, stretch_last=False)
        header = table.horizontalHeader()
        for column in (0, 1, 2, 4):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.stackedWidget_naver_orders.addWidget(self.emptyState_naver_orders)
        self.stackedWidget_naver_orders.addWidget(table)

        self.textBrowser_naver_order_detail = QTextBrowser()
        self.textBrowser_naver_order_detail.setObjectName("textBrowser_naver_order_detail")
        self.textBrowser_naver_order_detail.setOpenLinks(False)
        self.textBrowser_naver_order_detail.setPlaceholderText("목록에서 주문을 고르면 내용이 보입니다.")

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setObjectName("splitter_naver_orders")
        splitter.setChildrenCollapsible(False)
        self.stackedWidget_naver_orders.setMinimumWidth(200)
        self.textBrowser_naver_order_detail.setMinimumWidth(160)
        splitter.addWidget(self.stackedWidget_naver_orders)
        splitter.addWidget(self.textBrowser_naver_order_detail)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter, 1)

        tabs.addTab(page, "네이버 발주")
        self.ui.tab_naver_orders = page

        self.pushButton_naver_order_refresh.clicked.connect(self.refresh_naver_orders)
        self.comboBox_naver_order_filter.currentIndexChanged.connect(lambda _i: self._fill_naver_order_table())
        table.itemSelectionChanged.connect(self._show_selected_naver_order)
        set_status(self.label_naver_order_status, "idle", "")

    def _enter_naver_orders_tab(self):
        """처음 들어올 때 한 번 자동으로 불러온다(조회 전용)."""
        if not self._naver_orders_loaded_once:
            self.refresh_naver_orders()

    def refresh_naver_orders(self):
        if self._naver_orders_thread is not None:
            return
        self._naver_orders_loaded_once = True
        self.pushButton_naver_order_refresh.setEnabled(False)
        set_status(self.label_naver_order_status, "idle", "불러오는 중…")
        thread = _host(self).NaverPendingOrdersThread(self)
        thread.result_ready.connect(self._on_naver_orders_loaded)
        thread.finished.connect(self._on_naver_orders_thread_finished)
        self._naver_orders_thread = thread
        thread.start()

    def _on_naver_orders_thread_finished(self):
        self._naver_orders_thread = None
        self.pushButton_naver_order_refresh.setEnabled(True)

    def _on_naver_orders_loaded(self, payload):
        if not payload.get("ok"):
            error = str(payload.get("error") or "알 수 없는 오류")
            set_status(self.label_naver_order_status, "error", f"불러오지 못함: {error}", detail=error)
            print(f"! 네이버 발주 조회 실패: {error}")
            return
        self._naver_orders = list(payload.get("orders") or [])
        self._naver_orders_days = payload.get("days", 0)
        new = sum(1 for o in self._naver_orders if o.get("stage") == "신규주문")
        confirmed = len(self._naver_orders) - new
        claims = sum(1 for o in self._naver_orders if o.get("has_claim"))
        text = (f"신규 {new} · 발주확인 {confirmed}건 · {payload.get('fetched_at', '')} 조회"
                f" (최근 {payload.get('days', '')}일)")
        if claims:
            set_status(self.label_naver_order_status, "warn", text + f" · 클레임 {claims}건",
                       detail="취소·반품 요청 등 클레임이 걸린 주문이 있습니다. 처리 전에 확인하세요.")
        else:
            set_status(self.label_naver_order_status, "ok", text)
        self._fill_naver_order_table()

    def _fill_naver_order_table(self):
        table = self.tableWidget_naver_orders
        selected = self._selected_naver_order()
        selected_id = selected.get("order_id") if selected else None
        wanted = self.comboBox_naver_order_filter.currentText()
        orders = [o for o in self._naver_orders if wanted == "전체" or o.get("stage") == wanted]
        self._naver_orders_shown = orders
        table.blockSignals(True)
        table.setRowCount(len(orders))
        restore = -1
        for index, order in enumerate(orders):
            stage = order.get("stage") or ""
            values = (stage + (" · 클레임" if order.get("has_claim") else ""), order.get("paid_at"),
                      order.get("receiver"), naver_order_product_summary(order), won(order.get("total_amount")))
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                if column == 0:
                    tone = "danger" if order.get("has_claim") else STAGE_TONES.get(stage, "muted")
                    item.setForeground(QColor(TONES[tone][0]))
                if column == 3:
                    item.setToolTip("\n".join(f"{i.get('product')} × {i.get('quantity')}"
                                              for i in order.get("items") or []))
                if column == 4:
                    align_numeric(item)
                table.setItem(index, column, item)
            if order.get("order_id") == selected_id:
                restore = index
        table.blockSignals(False)
        self.stackedWidget_naver_orders.setCurrentWidget(table if orders else self.emptyState_naver_orders)
        if not orders:
            self.emptyState_naver_orders.set_text(
                f"최근 {self._naver_orders_days}일 안에 발송 전 주문이 없습니다." if wanted == "전체"
                else f"{wanted} 주문이 없습니다.")
        if restore >= 0:
            table.selectRow(restore)
        else:
            table.clearSelection()
        self._show_selected_naver_order()

    def _selected_naver_order(self):
        table = self.tableWidget_naver_orders
        index = table.currentRow()
        if table.selectionModel().hasSelection() and 0 <= index < len(self._naver_orders_shown):
            return self._naver_orders_shown[index]
        return None

    def _show_selected_naver_order(self):
        order = self._selected_naver_order()
        if order is None:
            self.textBrowser_naver_order_detail.clear()
            return
        self.textBrowser_naver_order_detail.setHtml(naver_order_detail_html(order))
