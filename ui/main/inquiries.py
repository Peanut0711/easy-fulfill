"""메인 창 「문의」 탭: 네이버 상품문의·고객문의를 API로 조회해 목록과 내용을 보여 준다(조회 전용).

답변 등록은 아직 없다. 「문의알림」 시트·슬랙 알림과는 따로 API 응답을 그대로 보여 주고 시트에는 쓰지 않는다.
"""

import html

from datetime import datetime

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QHBoxLayout,
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

from shared import app_log, theme
from shared.ui_components import EmptyState, fixed_width_button, setup_table
from shared.ui_style import set_button_role, set_status
from ui.main import host as _host

INQUIRY_PERIODS = (("최근 7일", 7), ("최근 30일", 30))
INQUIRY_COLUMNS = ("유형", "등록", "작성자", "상품", "내용")
INQUIRY_EMPTY_TEXT = "「새로고침」을 누르면 네이버 상품문의·고객문의를 불러옵니다."
# 측면 툴바 「문의」 배지(네이버 미답변 수): 시작 후 잠시 뒤 한 번, 그 뒤 5분마다 센다.
INQUIRY_BADGE_FIRST_DELAY_MS = 15_000
INQUIRY_BADGE_POLL_MS = 300_000


def inquiry_summary(row):
    """목록 「내용」 칸: 고객문의는 제목, 없으면 본문 첫 줄."""
    text = row.get("title") or row.get("content") or ""
    return " ".join(str(text).split())


def _esc(value):
    return html.escape(str(value or ""))


def inquiry_order_html(state):
    """고객문의 상세 아래 「주문·배송」 구역. state: None(해당 없음)·{"loading"}·작업 결과 dict."""
    if state is None:
        return ""
    head = "<p><b>주문·배송</b>"
    if state.get("loading"):
        return head + "<br>주문 정보를 불러오는 중…</p>"
    if not state.get("ok"):
        return head + f"<br>주문 정보를 불러오지 못했습니다: {_esc(state.get('error'))}</p>"
    items = state.get("items") or []
    if not items:
        return head + "<br>주문 정보를 찾지 못했습니다.</p>"
    if items[0].get("order_date"):
        head += f" (주문 {_esc(items[0]['order_date'])})"
    out = [head + "</p>"]
    for item in items:
        title = _esc(item.get("product")) or "(상품명 없음)"
        if item.get("option"):
            title += f" / {_esc(item['option'])}"
        if item.get("quantity"):
            title += f" × {_esc(item['quantity'])}"
        mark = "<b>[문의 상품]</b> " if item.get("is_inquiry_target") and len(items) > 1 else ""
        lines = [mark + title]
        status = f"상태: {_esc(item.get('status')) or '-'}"
        if item.get("claim"):
            status += f" · 클레임: {_esc(item['claim'])}"
        if item.get("delivery_method"):
            status += f" · {_esc(item['delivery_method'])}"
        lines.append(status)
        if item.get("tracking") or item.get("company"):
            ship = "배송: " + " ".join(x for x in (_esc(item.get("company")), _esc(item.get("tracking"))) if x)
            if item.get("delivery_status"):
                ship += f" ({_esc(item['delivery_status'])})"
            dates = []
            if item.get("send_date"):
                dates.append(f"발송 {_esc(item['send_date'])}")
            if item.get("delivered_date"):
                dates.append(f"도착 {_esc(item['delivered_date'])}")
            if dates:
                ship += " · " + " · ".join(dates)
            lines.append(ship)
            if item.get("wrong_tracking"):
                lines.append("<b>! 네이버에서 송장번호 오류로 표시된 건입니다</b>")
        elif item.get("send_date"):  # 방문수령 등 송장 없이 발송 처리된 건
            lines.append(f"발송 {_esc(item['send_date'])}"
                         + (f" · 도착 {_esc(item['delivered_date'])}" if item.get("delivered_date") else ""))
        elif item.get("status_code") == "PAYED":
            due = f" · 발송기한 {_esc(item['shipping_due'])}" if item.get("shipping_due") else ""
            lines.append(f"아직 발송 전{due}")
        if item.get("claim_tracking"):
            lines.append(f"클레임 송장: {_esc(item['claim_tracking'])}")
        out.append("<p>" + "<br>".join(lines) + "</p>")
    receivers = {(i.get("receiver"), i.get("zip"), i.get("address"), i.get("memo")) for i in items}
    if len(receivers) == 1:
        receiver, zip_code, address, memo = next(iter(receivers))
        lines = []
        if receiver:
            lines.append(f"수령인: {_esc(receiver)}")
        if address:
            lines.append(f"주소: {'(' + _esc(zip_code) + ') ' if zip_code else ''}{_esc(address)}")
        if memo:
            lines.append(f"배송메모: {_esc(memo)}")
        if lines:
            out.append("<p>" + "<br>".join(lines) + "</p>")
    else:
        for item in items:
            out.append(f"<p>{_esc(item.get('product'))} 수령인: {_esc(item.get('receiver'))}<br>"
                       f"주소: ({_esc(item.get('zip'))}) {_esc(item.get('address'))}"
                       + (f"<br>배송메모: {_esc(item.get('memo'))}" if item.get("memo") else "") + "</p>")
    return "".join(out)


def inquiry_detail_html(row, order_state=None):
    """오른쪽 상세 영역 HTML. 값은 모두 이스케이프한다. 고객문의는 아래에 주문·배송 구역을 붙인다."""
    esc = _esc

    def block(value):
        return esc(value).replace("\n", "<br>")

    head = f"[{esc(row.get('kind'))}]"
    if row.get("title"):
        head += " " + esc(row["title"])
    meta = [f"등록 {esc(row.get('created'))}"]
    if row.get("writer"):
        meta.append(f"작성자 {esc(row['writer'])}")
    if row.get("category"):
        meta.append(f"분류 {esc(row['category'])}")
    lines = [f"<p><b>{head}</b><br>{' · '.join(meta)}</p>"]
    facts = []
    if row.get("product"):
        no = f" ({esc(row['product_no'])})" if row.get("product_no") else ""
        facts.append(f"상품: {esc(row['product'])}{no}")
    if row.get("order_id"):
        facts.append(f"주문번호: {esc(row['order_id'])}")
    if facts:
        lines.append("<p>" + "<br>".join(facts) + "</p>")
    lines.append("<hr>")
    lines.append(f"<p>{block(row.get('content')) or '(내용 없음)'}</p>")
    lines.append("<hr>")
    if row.get("answered"):
        when = f" ({esc(row['answered_at'])})" if row.get("answered_at") else ""
        answer = block(row.get("answer")) or "(답변 내용을 응답에서 찾지 못했습니다)"
        lines.append(f"<p><b>답변{when}</b><br>{answer}</p>")
    elif row.get("answered") is False:
        lines.append("<p><b>미답변</b></p>")
    else:
        lines.append("<p><b>답변 여부 알 수 없음</b></p>")
    order = inquiry_order_html(order_state)
    if order:
        lines.append("<hr>")
        lines.append(order)
    return "".join(lines)


class InquiriesMixin:
    """MainWindow 믹스인. 메인 창 「문의」 탭: 네이버 문의 조회(답변 등록 없음)."""

    def _mount_inquiries_tab(self):
        self._inquiry_rows = []
        self._inquiry_thread = None
        self._inquiries_loaded_once = False
        self._inquiry_refresh_pending = False
        # 고객문의 주문·배송 정보: 주문번호 → 작업 결과. 목록을 새로 불러오면 비운다.
        self._inquiry_orders = {}
        self._inquiry_order_thread = None
        self._inquiry_order_wanted = None  # (주문번호, 상품주문번호 목록) — 지금 고른 문의
        tabs = self.ui.tabWidget
        page = QWidget()
        page.setObjectName("tab_inquiries")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
        layout.setSpacing(theme.SPACE_SM)

        bar = QHBoxLayout()
        bar.setSpacing(theme.SPACE_SM)
        self.comboBox_inquiry_period = QComboBox()
        self.comboBox_inquiry_period.setObjectName("comboBox_inquiry_period")
        for label, days in INQUIRY_PERIODS:
            self.comboBox_inquiry_period.addItem(label, days)
        self.comboBox_inquiry_period.setCurrentIndex(len(INQUIRY_PERIODS) - 1)
        self.checkBox_inquiry_unanswered = QCheckBox("미답변만")
        self.checkBox_inquiry_unanswered.setObjectName("checkBox_inquiry_unanswered")
        self.checkBox_inquiry_unanswered.setChecked(True)
        self.label_inquiry_status = QLabel()
        self.label_inquiry_status.setObjectName("label_inquiry_status")
        # 긴 오류 문구가 창 최소 폭을 늘리지 않게 한다(원문은 툴팁).
        self.label_inquiry_status.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.pushButton_inquiry_refresh = QPushButton("새로고침")
        self.pushButton_inquiry_refresh.setObjectName("pushButton_inquiry_refresh")
        set_button_role(self.pushButton_inquiry_refresh, "primary")
        bar.addWidget(self.comboBox_inquiry_period)
        bar.addWidget(self.checkBox_inquiry_unanswered)
        bar.addWidget(self.label_inquiry_status, 1)
        bar.addWidget(fixed_width_button(self.pushButton_inquiry_refresh))
        layout.addLayout(bar)

        self.stackedWidget_inquiry_list = QStackedWidget()
        self.stackedWidget_inquiry_list.setObjectName("stackedWidget_inquiry_list")
        self.emptyState_inquiry = EmptyState(INQUIRY_EMPTY_TEXT)
        self.tableWidget_inquiries = QTableWidget(0, len(INQUIRY_COLUMNS))
        self.tableWidget_inquiries.setObjectName("tableWidget_inquiries")
        self.tableWidget_inquiries.setHorizontalHeaderLabels(list(INQUIRY_COLUMNS))
        self.tableWidget_inquiries.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        setup_table(self.tableWidget_inquiries)
        for column, width in enumerate((64, 112, 72, 120)):
            self.tableWidget_inquiries.setColumnWidth(column, width)
        self.stackedWidget_inquiry_list.addWidget(self.emptyState_inquiry)
        self.stackedWidget_inquiry_list.addWidget(self.tableWidget_inquiries)

        self.textBrowser_inquiry_detail = QTextBrowser()
        self.textBrowser_inquiry_detail.setObjectName("textBrowser_inquiry_detail")
        self.textBrowser_inquiry_detail.setOpenLinks(False)
        self.textBrowser_inquiry_detail.setPlaceholderText("목록에서 문의를 고르면 내용이 보입니다.")

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setObjectName("splitter_inquiries")
        splitter.setChildrenCollapsible(False)
        self.stackedWidget_inquiry_list.setMinimumWidth(200)
        self.textBrowser_inquiry_detail.setMinimumWidth(160)
        splitter.addWidget(self.stackedWidget_inquiry_list)
        splitter.addWidget(self.textBrowser_inquiry_detail)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter, 1)

        tabs.addTab(page, "네이버 문의")
        self.ui.tab_inquiries = page

        self.pushButton_inquiry_refresh.clicked.connect(self.refresh_inquiries)
        self.tableWidget_inquiries.itemSelectionChanged.connect(self._show_selected_inquiry)
        self.comboBox_inquiry_period.currentIndexChanged.connect(lambda _i: self.refresh_inquiries())
        self.checkBox_inquiry_unanswered.toggled.connect(lambda _on: self.refresh_inquiries())
        set_status(self.label_inquiry_status, "idle", "")

        self._inquiry_badge_thread = None
        self._inquiry_badge_timer = QTimer(self)
        self._inquiry_badge_timer.setInterval(INQUIRY_BADGE_POLL_MS)
        self._inquiry_badge_timer.timeout.connect(self._poll_inquiry_badge)
        if _host(self).gspread is not None:
            self._inquiry_badge_timer.start()
            QTimer.singleShot(INQUIRY_BADGE_FIRST_DELAY_MS, self._poll_inquiry_badge)

    # ── 측면 툴바 「문의」 배지 ──────────────────────────────────────────
    def _poll_inquiry_badge(self):
        """네이버 미답변 문의 수를 백그라운드로 센다(이미 세는 중이거나 종료 중이면 건너뜀)."""
        if (self._inquiry_badge_thread is not None or _host(self).gspread is None
                or getattr(self, "_close_ready", False) or getattr(self, "_close_flush_timer", None) is not None):
            return
        thread = _host(self).NaverInquiryCountThread(self)
        thread.result_ready.connect(self._on_inquiry_badge_counted)
        thread.finished.connect(self._on_inquiry_badge_thread_finished)
        self._inquiry_badge_thread = thread
        thread.start()

    def _on_inquiry_badge_thread_finished(self):
        self._inquiry_badge_thread = None

    def _on_inquiry_badge_counted(self, payload):
        checked_at = payload.get("checked_at", "")
        if payload.get("ok"):
            self._set_inquiry_badge(payload.get("count", 0), checked_at, partial=payload.get("partial", False))
        else:
            error = str(payload.get("error") or "")
            app_log.log(f"[문의 배지] 미답변 수 확인 실패: {error}")
            # 틀린 숫자를 남기지 않도록 배지는 떼고 툴팁에 실패만 알린다.
            self.set_nav_rail_badge("tab_inquiries", 0, f"미답변 수 확인 실패 · {checked_at}")

    def _set_inquiry_badge(self, count, checked_at, partial=False):
        note = f"미답변 {count}건 · {checked_at} 확인" + (" (일부 실패)" if partial else "")
        self.set_nav_rail_badge("tab_inquiries", count, note)

    def _enter_inquiries_tab(self):
        """처음 들어올 때 한 번 자동으로 불러온다(조회 전용이라 바로 호출해도 된다)."""
        if not self._inquiries_loaded_once:
            self.refresh_inquiries()

    def refresh_inquiries(self):
        if self._inquiry_thread is not None:
            # 조회 중에 기간·필터를 바꾸면 끝난 뒤 바뀐 조건으로 한 번 더 부른다.
            self._inquiry_refresh_pending = True
            return
        self._inquiries_loaded_once = True
        self._inquiry_orders = {}  # 주문·배송 상태도 다시 읽게 한다
        days = self.comboBox_inquiry_period.currentData() or INQUIRY_PERIODS[-1][1]
        unanswered_only = self.checkBox_inquiry_unanswered.isChecked()
        self.pushButton_inquiry_refresh.setEnabled(False)
        set_status(self.label_inquiry_status, "idle", "불러오는 중…")
        thread = _host(self).NaverInquiryListThread(days, unanswered_only, self)
        thread.result_ready.connect(self._on_inquiries_loaded)
        thread.finished.connect(self._on_inquiry_thread_finished)
        self._inquiry_thread = thread
        thread.start()

    def _on_inquiry_thread_finished(self):
        self._inquiry_thread = None
        self.pushButton_inquiry_refresh.setEnabled(True)
        if self._inquiry_refresh_pending:
            self._inquiry_refresh_pending = False
            self.refresh_inquiries()

    def _on_inquiries_loaded(self, payload):
        if not payload.get("ok"):
            error = str(payload.get("error") or "알 수 없는 오류")
            set_status(self.label_inquiry_status, "error", f"불러오지 못함: {error}", detail=error)
            print(f"! 문의 조회 실패: {error}")
            return
        for kind, keys in (payload.get("sample_keys") or {}).items():
            app_log.log(f"[문의 탭] {kind} 응답 키: {', '.join(keys)}")
        rows = list(payload.get("rows") or [])
        self._fill_inquiry_table(rows)
        unanswered_only = self.checkBox_inquiry_unanswered.isChecked()
        label = "미답변" if unanswered_only else "문의"
        text = f"{label} {len(rows)}건 · {payload.get('fetched_at', '')} 조회"
        errors = payload.get("errors") or []
        if errors:
            set_status(self.label_inquiry_status, "warn", text + " (일부 실패)", detail="\n".join(errors))
            print("! 문의 조회 일부 실패: " + " / ".join(errors))
        else:
            set_status(self.label_inquiry_status, "ok", text)
        # 배지와 같은 조건(미답변만·같은 기간)으로 빠짐없이 불러왔으면 배지도 바로 맞춘다.
        if (payload.get("unanswered_only") and not errors
                and payload.get("days") == _host(self).NAVER_INQUIRY_LOOKBACK_DAYS):
            self._set_inquiry_badge(len(rows), payload.get("fetched_at") or datetime.now().strftime("%H:%M"))
        if not rows:
            self.emptyState_inquiry.set_text(
                "미답변 문의가 없습니다." if unanswered_only else "이 기간에 들어온 문의가 없습니다.")

    def _fill_inquiry_table(self, rows):
        table = self.tableWidget_inquiries
        selected = self._selected_inquiry_key()
        self._inquiry_rows = rows
        table.blockSignals(True)
        table.setRowCount(len(rows))
        restore = -1
        for index, row in enumerate(rows):
            values = (row.get("kind"), row.get("created"), row.get("writer"), row.get("product"),
                      inquiry_summary(row))
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                if column == len(values) - 1:
                    item.setToolTip(str(row.get("content") or ""))
                table.setItem(index, column, item)
            if (row.get("kind"), row.get("id")) == selected:
                restore = index
        table.blockSignals(False)
        self.stackedWidget_inquiry_list.setCurrentWidget(table if rows else self.emptyState_inquiry)
        if restore >= 0:
            table.selectRow(restore)
        else:
            table.clearSelection()
        self._show_selected_inquiry()

    def _selected_inquiry_key(self):
        row = self._selected_inquiry_row()
        return (row.get("kind"), row.get("id")) if row else None

    def _selected_inquiry_row(self):
        index = self.tableWidget_inquiries.currentRow()
        selected = self.tableWidget_inquiries.selectionModel().hasSelection()
        if selected and 0 <= index < len(self._inquiry_rows):
            return self._inquiry_rows[index]
        return None

    def _show_selected_inquiry(self):
        row = self._selected_inquiry_row()
        if row is None:
            self._inquiry_order_wanted = None
            self.textBrowser_inquiry_detail.clear()
            return
        order_state = None
        order_id = str(row.get("order_id") or "")
        if order_id:
            order_state = self._inquiry_orders.get(order_id)
            if order_state is None or not order_state.get("ok"):
                # 처음 고른 주문이거나 지난번에 실패했으면 (다시) 불러온다
                order_state = {"loading": True}
                self._inquiry_orders.pop(order_id, None)
                self._request_inquiry_order(order_id, row.get("product_order_ids") or [])
        else:
            self._inquiry_order_wanted = None
        self._render_inquiry_detail(row, order_state)

    def _render_inquiry_detail(self, row, order_state, keep_scroll=False):
        bar = self.textBrowser_inquiry_detail.verticalScrollBar()
        position = bar.value()
        self.textBrowser_inquiry_detail.setHtml(inquiry_detail_html(row, order_state))
        if keep_scroll:  # 같은 문의에 주문 정보만 채울 때는 읽던 위치를 지킨다
            bar.setValue(position)

    def _request_inquiry_order(self, order_id, product_order_ids):
        """고른 문의의 주문 정보를 불러온다. 이미 불러오는 중이면 끝난 뒤 마지막으로 고른 것을 불러온다."""
        self._inquiry_order_wanted = (order_id, list(product_order_ids))
        if self._inquiry_order_thread is not None:
            return
        thread = _host(self).NaverInquiryOrderThread(order_id, product_order_ids, self)
        thread.result_ready.connect(self._on_inquiry_order_loaded)
        thread.finished.connect(self._on_inquiry_order_thread_finished)
        self._inquiry_order_thread = thread
        thread.start()

    def _on_inquiry_order_loaded(self, payload):
        order_id = str(payload.get("order_id") or "")
        self._inquiry_orders[order_id] = payload
        if not payload.get("ok"):
            print(f"! 문의 주문 정보 조회 실패({order_id}): {payload.get('error')}")
        row = self._selected_inquiry_row()
        if row is not None and str(row.get("order_id") or "") == order_id:
            self._render_inquiry_detail(row, payload, keep_scroll=True)

    def _on_inquiry_order_thread_finished(self):
        self._inquiry_order_thread = None
        wanted = self._inquiry_order_wanted
        if wanted is not None and wanted[0] not in self._inquiry_orders:
            self._request_inquiry_order(*wanted)
