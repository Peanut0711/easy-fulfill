"""메인 창 「문의」 탭: 네이버 상품문의·고객문의와 쿠팡 상품문의·CS문의를 API로 조회해 목록과 내용을 보여 주고,
네이버 미답변 문의에 답변을 등록한다(쿠팡 답변은 WING에서).

「문의알림」 시트·슬랙 알림과는 따로 API 응답을 그대로 보여 주고 시트에는 쓰지 않는다.
"""

from datetime import datetime

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
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
from shared.ui_components import EmptyState, SectionCard, confirm_action, fixed_width_button, setup_table
from shared.ui_style import set_button_role, set_status
from ui.main import host as _host
from ui.main.rich_cards import (  # noqa: F401  (아래 이름들은 문의 상세 HTML에서 쓴다)
    TONES,
    block as _block,
    card as _card,
    chip as _chip,
    esc as _esc,
    facts as _facts,
    muted as _muted,
    section as _section,
)

# (이름, 일 수). 0 = 오늘 0시부터. 첫 항목이 기본값이다.
INQUIRY_PERIODS = (("오늘", 0), ("최근 7일", 7), ("최근 30일", 30))
# 「상품」 열 처음 폭(글자 수). 유형·등록·작성자는 내용에 맞추고 「내용」은 남는 폭을 쓴다.
INQUIRY_PRODUCT_COLUMN_CHARS = 10
INQUIRY_COLUMNS = ("유형", "등록", "작성자", "상품", "내용")
INQUIRY_EMPTY_TEXT = "「새로고침」을 누르면 네이버·쿠팡 문의를 불러옵니다."
# 측면 툴바 「문의」 배지(네이버 미답변 수, 쿠팡은 세지 않음): 시작 후 잠시 뒤 한 번, 그 뒤 5분마다 센다.
INQUIRY_BADGE_FIRST_DELAY_MS = 15_000
INQUIRY_BADGE_POLL_MS = 300_000


def inquiry_summary(row):
    """목록 「내용」 칸: 고객문의는 제목, 없으면 본문 첫 줄."""
    text = row.get("title") or row.get("content") or ""
    return " ".join(str(text).split())


KIND_TONES = {"고객문의": "info", "상품문의": "primary", "쿠팡 CS문의": "info", "쿠팡 상품문의": "primary"}


def inquiry_store(row):
    """행의 스토어. 네이버 행에는 store 칸이 없다."""
    return row.get("store") or "네이버"


def inquiry_kind_label(row):
    """목록 「유형」 칸·상세 배지 글자: 「네이버 상품문의」, 「쿠팡 CS문의」."""
    kind = row.get("kind") or ""
    store = inquiry_store(row)
    return kind if kind.startswith(store) else f"{store} {kind}".strip()


def inquiry_answerable(row):
    """앱에서 답변을 등록할 수 있는 문의인지(네이버 미답변만. 쿠팡 답변은 WING에서 한다)."""
    return row is not None and inquiry_store(row) == "네이버" and row.get("answered") is False


# 상품주문 상태 → 색: 발송 전은 처리할 일(주의), 배송 중은 정보, 끝난 건 정상, 취소·반품·교환은 위험
ORDER_STATUS_TONES = {
    "PAYMENT_WAITING": "muted",
    "PAYED": "warn",
    "DELIVERING": "info",
    "DELIVERED": "ok",
    "PURCHASE_DECIDED": "ok",
    "EXCHANGED": "danger",
    "CANCELED": "danger",
    "RETURNED": "danger",
    "CANCELED_BY_NOPAYMENT": "danger",
}


def _order_item_card(item, mark_target):
    """상품주문 1건 카드: 상품·상태 배지·배송(송장 굵게)·날짜·클레임."""
    tone = ORDER_STATUS_TONES.get(item.get("status_code") or "", "muted")
    head = ""
    if mark_target:
        head += _chip("문의 상품", "primary") + "&nbsp; "
    head += f'<b>{_esc(item.get("product")) or "(상품명 없음)"}</b>'
    sub = []
    if item.get("option"):
        sub.append(_esc(item["option"]))
    if item.get("quantity"):
        sub.append(f"× {_esc(item['quantity'])}")
    lines = [head]
    if sub:
        lines.append(_muted(" · ".join(sub)))
    chips = [_chip(item.get("status") or "상태 모름", tone)]
    if item.get("claim"):
        chips.append(_chip(item["claim"], "danger"))
    if item.get("delivery_method"):
        chips.append(_chip(item["delivery_method"], "muted"))
    lines.append("&nbsp;".join(chips))
    dates = []
    if item.get("send_date"):
        dates.append(f"발송 {_esc(item['send_date'])}")
    if item.get("delivered_date"):
        dates.append(f"도착 {_esc(item['delivered_date'])}")
    if item.get("tracking") or item.get("company"):
        ship = " ".join(x for x in (_esc(item.get("company")), f'<b>{_esc(item.get("tracking"))}</b>'
                                    if item.get("tracking") else "") if x)
        if item.get("delivery_status"):
            ship += f" · {_esc(item['delivery_status'])}"
        lines.append(f"배송 {ship}")
        if dates:
            lines.append(_muted(" → ".join(dates)))
        if item.get("wrong_tracking"):
            lines.append(f'<span style="color:{theme.DANGER}; font-weight:600;">'
                         f'! 네이버에서 송장번호 오류로 표시된 건입니다</span>')
    elif dates:  # 방문수령 등 송장 없이 발송 처리된 건
        lines.append(_muted(" → ".join(dates)))
    elif item.get("status_code") == "PAYED":
        due = f" · 발송기한 {_esc(item['shipping_due'])}" if item.get("shipping_due") else ""
        lines.append(f'<span style="color:{theme.WARN}; font-weight:600;">아직 발송 전{due}</span>')
    if item.get("claim_tracking"):
        lines.append(_muted(f"클레임 송장 {_esc(item['claim_tracking'])}"))
    return _card("<br>".join(lines), tone, plain_bg=True)


def inquiry_order_html(state):
    """고객문의 상세 아래 「주문·배송」 구역. state: None(해당 없음)·{"loading"}·작업 결과 dict."""
    if state is None:
        return ""
    if state.get("loading"):
        return _section("주문·배송") + _muted("주문 정보를 불러오는 중…")
    if not state.get("ok"):
        return _section("주문·배송") + _card(
            f"주문 정보를 불러오지 못했습니다: {_esc(state.get('error'))}", "danger")
    items = state.get("items") or []
    if not items:
        return _section("주문·배송") + _muted("주문 정보를 찾지 못했습니다.")
    note = f"주문 {items[0]['order_date']}" if items[0].get("order_date") else ""
    out = [_section("주문·배송", note=note)]
    for item in items:
        out.append(_order_item_card(item, item.get("is_inquiry_target") and len(items) > 1))
    receivers = []
    for item in items:
        key = (item.get("receiver"), item.get("zip"), item.get("address"), item.get("memo"))
        if key not in [r[0] for r in receivers]:
            receivers.append((key, item.get("product")))
    for (receiver, zip_code, address, memo), product in receivers:
        title = "받는 분" if len(receivers) == 1 else f"받는 분 · {product}"
        zip_html = f"({_esc(zip_code)}) " if zip_code else ""
        facts = _facts([("수령인", f"<b>{_esc(receiver)}</b>" if receiver else ""),
                        ("주소", f"{zip_html}{_esc(address)}" if address else ""),
                        ("배송메모", _esc(memo))])
        if facts:
            out.append(_section(title))
            out.append(facts)
    return "".join(out)


def inquiry_detail_html(row, order_state=None):
    """오른쪽 상세 영역 HTML. 값은 모두 이스케이프한다. 고객문의는 아래에 주문·배송 구역을 붙인다."""
    kind = row.get("kind") or ""
    chips = [_chip(inquiry_kind_label(row), KIND_TONES.get(kind, "muted"))]
    if row.get("answered"):
        chips.append(_chip("답변 완료", "ok"))
    elif row.get("answered") is False:
        chips.append(_chip("미답변", "danger"))
    else:
        chips.append(_chip(row.get("status_label") or "답변 여부 알 수 없음", "muted"))
    title = row.get("title") or row.get("product") or "(제목 없음)"
    meta = [f"등록 {_esc(row.get('created'))}"]
    if row.get("writer"):
        meta.append(f"작성자 {_esc(row['writer'])}")
    if row.get("category"):
        meta.append(f"분류 {_esc(row['category'])}")
    product = ""
    if row.get("product") and row.get("title"):
        product = _esc(row["product"]) + (_muted(f" ({_esc(row['product_no'])})") if row.get("product_no") else "")
    facts = [("상품", product), ("주문번호", _esc(row.get("order_id")))]
    if not row.get("title"):  # 상품문의는 제목 자리에 상품명이 있어 상품번호만 보인다
        facts = [("상품번호", _esc(row.get("product_no")))] + facts
    out = [
        f'<p style="margin-bottom:{theme.SPACE_XS}px;">{"&nbsp;".join(chips)}</p>',
        f'<p style="margin-top:0; margin-bottom:{theme.SPACE_XS}px; font-size:{theme.FONT_SIZE + 3}pt; '
        f'font-weight:600; color:{theme.TEXT};">{_esc(title)}</p>',
        f'<p style="margin-top:0;">{_muted(" · ".join(meta))}</p>',
        _facts(facts),
        _section("문의", "primary"),
        _card(_block(row.get("content")) or _muted("(내용 없음)"), "primary"),
    ]
    if row.get("replies"):
        out += _replies_html(row)
    elif row.get("answered"):
        when = row.get("answered_at") or ""
        answer = _block(row.get("answer")) or _muted("(답변 내용을 응답에서 찾지 못했습니다)")
        out += [_section("답변", "ok", note=when), _card(answer, "ok")]
    elif row.get("answered") is False:
        out += [_section("답변", "danger"), _card("아직 답변하지 않았습니다.", "danger")]
    if inquiry_store(row) == "쿠팡" and row.get("answered") is False:
        out.append(f'<p>{_muted("쿠팡 문의 답변은 쿠팡 WING에서 등록합니다.")}</p>')
    out.append(inquiry_order_html(order_state))
    return "".join(out)


def _replies_html(row):
    """쿠팡 CS문의 상담 기록: 쿠팡 상담원(판매자 답변 요청은 주의 색)과 판매자 답변을 시간 순 카드로."""
    out = [_section("상담 기록", "ok" if row.get("answered") else "muted")]
    for reply in row["replies"]:
        if reply.get("vendor"):
            tone = "ok"
        elif "답변 요청" in (reply.get("who") or ""):
            tone = "warn"
        else:
            tone = "info"
        head = f'<b>{_esc(reply.get("who"))}</b>'
        if reply.get("at"):
            head += " " + _muted(_esc(reply["at"]))
        out.append(_card(head + "<br>" + (_block(reply.get("content")) or _muted("(내용 없음)")), tone))
    if row.get("answered") is False:
        out.append(_card("아직 판매자가 답변하지 않았습니다.", "danger"))
    return out


class InquiriesMixin:
    """MainWindow 믹스인. 메인 창 「문의」 탭: 네이버·쿠팡 문의 조회와 네이버 문의 답변 등록."""

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
        self.comboBox_inquiry_period.setCurrentIndex(0)
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
        # 픽셀 고정 폭은 화면 배율·글꼴에 따라 「고객...」처럼 잘려서, 짧은 열은 내용에 맞춘다.
        header = self.tableWidget_inquiries.horizontalHeader()
        for column in (0, 1, 2):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        self._inquiry_product_column_sized = False
        self.stackedWidget_inquiry_list.addWidget(self.emptyState_inquiry)
        self.stackedWidget_inquiry_list.addWidget(self.tableWidget_inquiries)

        self.textBrowser_inquiry_detail = QTextBrowser()
        self.textBrowser_inquiry_detail.setObjectName("textBrowser_inquiry_detail")
        self.textBrowser_inquiry_detail.setOpenLinks(False)
        self.textBrowser_inquiry_detail.setPlaceholderText("목록에서 문의를 고르면 내용이 보입니다.")

        # 오른쪽: 위 상세 보기 · 아래 답변 작성(미답변 문의를 골랐을 때만 보인다)
        self.textBrowser_inquiry_detail.setMinimumHeight(80)
        self._build_inquiry_answer_panel()
        detail_split = QSplitter(Qt.Orientation.Vertical)
        detail_split.setObjectName("splitter_inquiry_detail")
        detail_split.setChildrenCollapsible(False)
        detail_split.addWidget(self.textBrowser_inquiry_detail)
        detail_split.addWidget(self.card_inquiry_answer)
        detail_split.setStretchFactor(0, 3)
        detail_split.setStretchFactor(1, 1)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setObjectName("splitter_inquiries")
        splitter.setChildrenCollapsible(False)
        self.stackedWidget_inquiry_list.setMinimumWidth(200)
        detail_split.setMinimumWidth(160)
        splitter.addWidget(self.stackedWidget_inquiry_list)
        splitter.addWidget(detail_split)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter, 1)

        tabs.addTab(page, "문의")
        self.ui.tab_inquiries = page

        self.pushButton_inquiry_refresh.clicked.connect(self.refresh_inquiries)
        self.tableWidget_inquiries.itemSelectionChanged.connect(self._show_selected_inquiry)
        self.comboBox_inquiry_period.currentIndexChanged.connect(lambda _i: self.refresh_inquiries())
        self.checkBox_inquiry_unanswered.toggled.connect(lambda _on: self.refresh_inquiries())
        set_status(self.label_inquiry_status, "idle", "")

        self._inquiry_badge_thread = None
        self._inquiry_known_ids = None  # 마지막으로 센 미답변 「종류:번호」(None = 아직 안 셈)
        self._inquiry_auto_refreshing = False
        self._inquiry_new_pending = False  # 다른 탭에 있을 때 새 문의가 생김 → 들어오면 오늘·미답변만으로
        self._inquiry_badge_timer = QTimer(self)
        self._inquiry_badge_timer.setInterval(INQUIRY_BADGE_POLL_MS)
        self._inquiry_badge_timer.timeout.connect(self._poll_inquiry_badge)
        if _host(self).gspread is not None:
            self._inquiry_badge_timer.start()
            QTimer.singleShot(INQUIRY_BADGE_FIRST_DELAY_MS, self._poll_inquiry_badge)

    # ── 답변 작성(고객에게 보이는 쓰기) ───────────────────────────────────
    def _build_inquiry_answer_panel(self):
        self._inquiry_drafts = {}  # (종류, 번호) → 쓰던 답변
        self._inquiry_answer_key = None  # 지금 답변 칸이 가리키는 문의
        self._inquiry_answer_thread = None
        self._inquiry_answer_sending_key = None
        self._inquiry_templates = None  # None = 아직 안 읽음
        self._inquiry_templates_thread = None

        self.comboBox_inquiry_template = QComboBox()
        self.comboBox_inquiry_template.setObjectName("comboBox_inquiry_template")
        self.comboBox_inquiry_template.addItem("템플릿 넣기…")
        self.comboBox_inquiry_template.setEnabled(False)
        self.card_inquiry_answer = SectionCard("답변 작성", actions=[self.comboBox_inquiry_template])
        self.card_inquiry_answer.setObjectName("card_inquiry_answer")
        self.plainTextEdit_inquiry_answer = QPlainTextEdit()
        self.plainTextEdit_inquiry_answer.setObjectName("plainTextEdit_inquiry_answer")
        self.plainTextEdit_inquiry_answer.setPlaceholderText("고객에게 보낼 답변을 입력하세요.")
        self.plainTextEdit_inquiry_answer.setMinimumHeight(60)
        self.label_inquiry_answer_status = QLabel()
        self.label_inquiry_answer_status.setObjectName("label_inquiry_answer_status")
        self.label_inquiry_answer_status.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.pushButton_inquiry_answer_send = QPushButton("답변 등록")
        self.pushButton_inquiry_answer_send.setObjectName("pushButton_inquiry_answer_send")
        set_button_role(self.pushButton_inquiry_answer_send, "primary")
        self.pushButton_inquiry_answer_send.setEnabled(False)
        row = QHBoxLayout()
        row.setSpacing(theme.SPACE_SM)
        row.addWidget(self.label_inquiry_answer_status, 1)
        row.addWidget(fixed_width_button(self.pushButton_inquiry_answer_send))
        self.card_inquiry_answer.body.addWidget(self.plainTextEdit_inquiry_answer, 1)
        self.card_inquiry_answer.body.addLayout(row)
        self.card_inquiry_answer.setVisible(False)
        set_status(self.label_inquiry_answer_status, "idle", "")

        self.plainTextEdit_inquiry_answer.textChanged.connect(self._on_inquiry_answer_text_changed)
        self.pushButton_inquiry_answer_send.clicked.connect(self._send_inquiry_answer)
        self.comboBox_inquiry_template.activated.connect(self._insert_inquiry_template)

    def _sync_inquiry_answer_panel(self, row):
        """고른 문의에 맞춰 답변 칸을 보이거나 숨기고, 문의마다 쓰던 답변을 바꿔 끼운다."""
        key = (row.get("kind"), row.get("id")) if row else None
        show = inquiry_answerable(row)
        # 같은 문의라도 새로고침 뒤 답변됨으로 바뀌면 칸을 숨겨야 해서 보임 여부까지 비교한다
        if key == self._inquiry_answer_key and show == self.card_inquiry_answer.isVisibleTo(self.card_inquiry_answer.parentWidget()):
            return
        if self._inquiry_answer_key is not None and self._inquiry_answer_key != key:  # 앞 문의에 쓰던 글을 기억
            self._inquiry_drafts[self._inquiry_answer_key] = self.plainTextEdit_inquiry_answer.toPlainText()
        self._inquiry_answer_key = key
        self.card_inquiry_answer.setVisible(show)
        self.plainTextEdit_inquiry_answer.blockSignals(True)
        self.plainTextEdit_inquiry_answer.setPlainText(self._inquiry_drafts.get(key, "") if show else "")
        self.plainTextEdit_inquiry_answer.blockSignals(False)
        if show:
            if row.get("kind") == "상품문의":
                public, detail = "상품 페이지에 공개", "상품문의 답변은 상품 페이지에 공개되어 다른 고객도 봅니다."
            else:
                public, detail = "고객에게만 보임", "고객문의 답변은 문의한 고객에게만 보입니다."
            sending = self._inquiry_answer_sending_key == key
            set_status(self.label_inquiry_answer_status, "idle", "보내는 중…" if sending else public, detail=detail)
            self._load_inquiry_templates()
        self._update_inquiry_answer_button()

    def _on_inquiry_answer_text_changed(self):
        if self._inquiry_answer_key is not None:
            self._inquiry_drafts[self._inquiry_answer_key] = self.plainTextEdit_inquiry_answer.toPlainText()
        self._update_inquiry_answer_button()

    def _update_inquiry_answer_button(self):
        has_text = bool(self.plainTextEdit_inquiry_answer.toPlainText().strip())
        self.pushButton_inquiry_answer_send.setEnabled(has_text and self._inquiry_answer_thread is None)

    def _load_inquiry_templates(self):
        if self._inquiry_templates is not None or self._inquiry_templates_thread is not None:
            return
        if _host(self).gspread is None:
            return
        thread = _host(self).NaverQnaTemplatesThread(self)
        thread.result_ready.connect(self._on_inquiry_templates_loaded)
        thread.finished.connect(lambda: setattr(self, "_inquiry_templates_thread", None))
        self._inquiry_templates_thread = thread
        thread.start()

    def _on_inquiry_templates_loaded(self, payload):
        combo = self.comboBox_inquiry_template
        if not payload.get("ok"):
            self._inquiry_templates = None  # 다음에 답변 칸을 열 때 다시 읽는다
            combo.setEnabled(False)
            combo.setToolTip(f"템플릿을 읽지 못했습니다: {payload.get('error')}")
            return
        self._inquiry_templates = list(payload.get("templates") or [])
        while combo.count() > 1:
            combo.removeItem(1)
        for template in self._inquiry_templates:
            combo.addItem(template["subject"])
        combo.setEnabled(bool(self._inquiry_templates))
        combo.setToolTip("" if self._inquiry_templates else "판매자센터에 등록된 상품문의 답변 템플릿이 없습니다.")

    def _insert_inquiry_template(self, index):
        """고른 템플릿 내용을 입력란의 커서 자리에 넣는다(덮어쓰지 않는다)."""
        if index <= 0 or not self._inquiry_templates or index > len(self._inquiry_templates):
            return
        self.plainTextEdit_inquiry_answer.insertPlainText(self._inquiry_templates[index - 1]["content"])
        self.plainTextEdit_inquiry_answer.setFocus()
        self.comboBox_inquiry_template.setCurrentIndex(0)

    def _send_inquiry_answer(self):
        row = self._selected_inquiry_row()
        text = self.plainTextEdit_inquiry_answer.toPlainText().strip()
        if row is None or not text or row.get("answered") is not False or self._inquiry_answer_thread is not None:
            return
        kind = row.get("kind")
        title = row.get("title") or row.get("product") or ""
        public = ("상품문의 답변은 상품 페이지에 공개되어 다른 고객도 봅니다."
                  if kind == "상품문의" else "문의한 고객에게 보입니다.")
        preview = text if len(text) <= 300 else text[:300] + "…"
        if not confirm_action(
                self, "답변 등록", f"{kind} 「{title}」({row.get('writer') or '-'})에 이 답변을 등록할까요?",
                "답변 등록", destructive=True,
                detail=f"{public}\n등록한 답변은 앱에서 고칠 수 없습니다(판매자센터에서 수정).\n\n{preview}"):
            return
        key = (kind, row.get("id"))
        self._inquiry_answer_sending_key = key
        set_status(self.label_inquiry_answer_status, "idle", "보내는 중…")
        thread = _host(self).NaverInquiryAnswerThread(kind, row.get("id"), row.get("created"), text, self)
        thread.result_ready.connect(self._on_inquiry_answer_sent)
        thread.finished.connect(self._on_inquiry_answer_thread_finished)
        self._inquiry_answer_thread = thread
        self._update_inquiry_answer_button()
        thread.start()

    def _on_inquiry_answer_thread_finished(self):
        self._inquiry_answer_thread = None
        self._inquiry_answer_sending_key = None
        self._update_inquiry_answer_button()

    def _on_inquiry_answer_sent(self, payload):
        key = (payload.get("kind"), payload.get("id"))
        current = key == self._inquiry_answer_key
        if payload.get("ok"):
            self._inquiry_drafts.pop(key, None)
            if current:
                self.plainTextEdit_inquiry_answer.blockSignals(True)
                self.plainTextEdit_inquiry_answer.clear()
                self.plainTextEdit_inquiry_answer.blockSignals(False)
                set_status(self.label_inquiry_answer_status, "ok", "답변을 등록했습니다.")
            print(f"✓ 문의 답변 등록: {key[0]} {key[1]}")
            self.notify("ok", f"{key[0]} 답변을 등록했습니다.")
        else:
            error = str(payload.get("error") or "알 수 없는 오류")
            if payload.get("uncertain"):
                error = ("보내는 중 연결이 끊겨 등록됐는지 알 수 없습니다. 목록을 새로고침해 답변 여부를 확인한 뒤 "
                         f"다시 보내세요. ({error})")
            kind = "warn" if payload.get("already_answered") else "error"
            if current:
                set_status(self.label_inquiry_answer_status, kind, error, detail=error)
            print(f"! 문의 답변 등록 실패({key[0]} {key[1]}): {error}")
            self.notify(kind, f"답변을 등록하지 못했습니다: {error}")
        # 답변 상태를 다시 읽어 목록·배지를 맞춘다(이미 답변됨·알 수 없음일 때도).
        self.refresh_inquiries()
        self._poll_inquiry_badge()

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
            partial = payload.get("partial", False)
            self._set_inquiry_badge(payload.get("count", 0), checked_at, partial=partial)
            if payload.get("ids") is not None:
                self._note_inquiry_ids(set(payload["ids"]), partial=partial, auto_refresh=True)
        else:
            error = str(payload.get("error") or "")
            app_log.log(f"[문의 배지] 미답변 수 확인 실패: {error}")
            # 틀린 숫자를 남기지 않도록 배지는 떼고 툴팁에 실패만 알린다.
            self.set_nav_rail_badge("tab_inquiries", 0, f"미답변 수 확인 실패 · {checked_at}")

    def _note_inquiry_ids(self, ids, partial=False, auto_refresh=False):
        """미답변 문의 번호 목록(「종류:번호」)을 기억하고, 처음 보는 번호가 생기면 목록에 반영한다.

        - 「문의」 탭을 보고 있으면 지금 기간·필터 그대로 바로 새로고침한다(어느 기간이든 오늘이 들어 있어 새 문의가 보인다).
        - 다른 탭에 있으면 표시만 해 두고, 「문의」 탭에 들어올 때 오늘·미답변만으로 바꿔 새로고침한다
          (_enter_inquiries_tab).
        건수가 아니라 번호로 비교해 「한 건 답변 + 한 건 새로 옴」(건수 그대로)도 잡는다. 앱을 켠 뒤 처음
        센 결과는 기준으로만 삼는다. 한 종류만 센 결과(partial)는 빠진 번호를 「답변됨」으로 보지 않도록
        기존 목록에 더하기만 한다.
        """
        previous = self._inquiry_known_ids
        new = ids - previous if previous is not None else set()
        self._inquiry_known_ids = (previous or set()) | ids if partial else set(ids)
        if not new or not auto_refresh:
            return
        if self._inquiries_tab_shown():
            app_log.log(f"[문의 탭] 새 미답변 문의 {len(new)}건 → 목록 바로 새로고침: {', '.join(sorted(new))}")
            self._inquiry_auto_refreshing = True
            self.refresh_inquiries()
        else:
            app_log.log(f"[문의 탭] 새 미답변 문의 {len(new)}건 → 탭에 들어오면 오늘·미답변만으로 새로고침: "
                        f"{', '.join(sorted(new))}")
            self._inquiry_new_pending = True

    def _inquiries_tab_shown(self):
        tabs = getattr(self.ui, "tabWidget", None)
        return tabs is not None and tabs.currentWidget() is self.ui.tab_inquiries

    def _set_inquiry_badge(self, count, checked_at, partial=False):
        note = f"네이버 미답변 {count}건 · {checked_at} 확인" + (" (일부 실패)" if partial else "")
        self.set_nav_rail_badge("tab_inquiries", count, note)

    def _enter_inquiries_tab(self):
        """처음 들어올 때 한 번 자동으로 불러온다. 다른 탭에 있는 동안 새 미답변 문의가 생겼으면
        기간을 「오늘」, 「미답변만」을 켬으로 바꿔 새로고침한다(새 문의가 바로 보이게)."""
        if self._inquiry_new_pending:
            self._inquiry_new_pending = False
            for widget in (self.comboBox_inquiry_period, self.checkBox_inquiry_unanswered):
                widget.blockSignals(True)  # 바꿀 때마다 따로 조회하지 않고 아래에서 한 번만
            self.comboBox_inquiry_period.setCurrentIndex(0)
            self.checkBox_inquiry_unanswered.setChecked(True)
            for widget in (self.comboBox_inquiry_period, self.checkBox_inquiry_unanswered):
                widget.blockSignals(False)
            self._inquiry_auto_refreshing = True
            self.refresh_inquiries()
            return
        if not self._inquiries_loaded_once:
            self.refresh_inquiries()

    def refresh_inquiries(self):
        if self._inquiry_thread is not None:
            # 조회 중에 기간·필터를 바꾸면 끝난 뒤 바뀐 조건으로 한 번 더 부른다.
            self._inquiry_refresh_pending = True
            return
        self._inquiries_loaded_once = True
        self._inquiry_orders = {}  # 주문·배송 상태도 다시 읽게 한다
        days = self.comboBox_inquiry_period.currentData()
        if days is None:  # 0(오늘)도 올바른 값이라 `or`로 기본값을 고르지 않는다
            days = INQUIRY_PERIODS[0][1]
        unanswered_only = self.checkBox_inquiry_unanswered.isChecked()
        self.pushButton_inquiry_refresh.setEnabled(False)
        set_status(self.label_inquiry_status, "idle", "불러오는 중…")
        thread = _host(self).InquiryListThread(days, unanswered_only, self)
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
        missing = payload.get("missing_keys") or []
        if missing:
            text += f" · {'·'.join(missing)} 키 없음"
        if self._inquiry_auto_refreshing:
            text += " · 새 문의 자동 반영"
            self._inquiry_auto_refreshing = False
        errors = payload.get("errors") or []
        if errors:
            set_status(self.label_inquiry_status, "warn", text + " (일부 실패)", detail="\n".join(errors))
            print("! 문의 조회 일부 실패: " + " / ".join(errors))
        else:
            set_status(self.label_inquiry_status, "ok", text)
        # 배지와 같은 조건(네이버 미답변만·같은 기간)으로 네이버를 빠짐없이 불러왔으면 배지도 바로 맞춘다.
        # 배지는 네이버만 세므로 쿠팡 행은 빼고, 쿠팡 쪽 실패와는 상관없다.
        if (payload.get("unanswered_only") and payload.get("naver_complete")
                and payload.get("days") == _host(self).NAVER_INQUIRY_LOOKBACK_DAYS):
            naver_rows = [r for r in rows if inquiry_store(r) == "네이버"]
            self._set_inquiry_badge(len(naver_rows), payload.get("fetched_at") or datetime.now().strftime("%H:%M"))
            self._note_inquiry_ids({f"{r.get('kind', '')}:{r.get('id', '')}" for r in naver_rows})
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
            product = row.get("product") or (f"상품 {row['product_no']}" if row.get("product_no") else "")
            values = (inquiry_kind_label(row), row.get("created"), row.get("writer"), product,
                      inquiry_summary(row))
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                if column == len(values) - 1:
                    item.setToolTip(str(row.get("content") or ""))
                table.setItem(index, column, item)
            if (row.get("kind"), row.get("id")) == selected:
                restore = index
        table.blockSignals(False)
        if rows and not self._inquiry_product_column_sized:
            # 처음 채울 때 한 번만 글자 폭으로 정한다(그 뒤로는 사용자가 바꾼 폭을 지킨다)
            self._inquiry_product_column_sized = True
            table.setColumnWidth(3, table.fontMetrics().horizontalAdvance("가" * INQUIRY_PRODUCT_COLUMN_CHARS) + 16)
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
        self._sync_inquiry_answer_panel(row)
        if row is None:
            self._inquiry_order_wanted = None
            self.textBrowser_inquiry_detail.clear()
            return
        order_state = None
        # 주문·배송은 네이버 고객문의만 불러온다(쿠팡 문의의 주문번호는 상세에 번호만 보인다).
        order_id = str(row.get("order_id") or "") if inquiry_store(row) == "네이버" else ""
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
