"""메인 창 「문의」 탭: 네이버 상품문의·고객문의를 API로 조회해 목록과 내용을 보여 준다(조회 전용).

답변 등록은 아직 없다. 「문의알림」 시트·슬랙 알림과는 따로 API 응답을 그대로 보여 주고 시트에는 쓰지 않는다.
"""

import html

from PySide6.QtCore import Qt
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


def inquiry_summary(row):
    """목록 「내용」 칸: 고객문의는 제목, 없으면 본문 첫 줄."""
    text = row.get("title") or row.get("content") or ""
    return " ".join(str(text).split())


def inquiry_detail_html(row):
    """오른쪽 상세 영역 HTML. 값은 모두 이스케이프한다."""
    def esc(value):
        return html.escape(str(value or ""))

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
    return "".join(lines)


class InquiriesMixin:
    """MainWindow 믹스인. 메인 창 「문의」 탭: 네이버 문의 조회(답변 등록 없음)."""

    def _mount_inquiries_tab(self):
        self._inquiry_rows = []
        self._inquiry_thread = None
        self._inquiries_loaded_once = False
        self._inquiry_refresh_pending = False
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
            self.textBrowser_inquiry_detail.clear()
            return
        self.textBrowser_inquiry_detail.setHtml(inquiry_detail_html(row))
