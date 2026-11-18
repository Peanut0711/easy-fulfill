"""마크다운 사은품 집계를 동기화하고 관리자에게 보여 준다."""

from collections import defaultdict
from datetime import date, datetime, timedelta
import logging
from threading import Thread

from PySide6.QtCore import QDate, QObject, Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QDateEdit, QDialog, QDialogButtonBox,
    QHBoxLayout, QHeaderView, QLabel, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
)

from orders.gift_event import GIFT_EVENT_STORES
from orders.gift_history import KST, summarize_gift_records
from shared.qt_background import run_in_background
from sheets.gift_history import append_gift_records, read_gift_records


class GiftHistorySync(QObject):
    def __init__(self, open_spreadsheet, outbox, parent):
        super().__init__(parent)
        self._open_spreadsheet = open_spreadsheet
        self._outbox = outbox
        self._worker = None
        self._timer = QTimer(self)
        self._timer.setInterval(60_000)
        self._timer.timeout.connect(self.flush)
        self._timer.start()
        QTimer.singleShot(4000, self, self.flush)

    def enqueue(self, records):
        # 주문을 닫거나 앱을 종료해도 이어서 반영할 수 있게 디스크에 먼저 저장한다.
        self._outbox.add(records)
        self.flush()

    def flush(self):
        if not self._outbox.path.exists() or (self._worker is not None and self._worker.is_alive()):
            return
        self._worker = Thread(target=self._sync, name="gift-history-sync", daemon=True)
        self._worker.start()

    def _sync(self):
        try:
            records = self._outbox.list_pending()
            if records:
                append_gift_records(self._open_spreadsheet(), records)
                self._outbox.remove(records)
        except Exception as error:
            logging.warning("사은품 기록 공유 시트 동기화 실패(재시도 예정): %s", error)

    def read_report(self):
        records = read_gift_records(self._open_spreadsheet())
        stored = {record.key: record.recorded_at for record in records}
        pending_count = sum(record.key not in stored or record.recorded_at < stored[record.key]
                            for record in self._outbox.list_pending())
        return records, pending_count


class GiftHistoryDialog(QDialog):
    def __init__(self, read_report, parent=None):
        super().__init__(parent)
        self._read_report = read_report
        self._records = None
        self._pending_count = 0
        self._loading = False
        self.setWindowTitle("사은품 일별 집계")
        self.resize(760, 480)

        layout = QVBoxLayout(self)
        description = QLabel("한국시간 기준으로 주문 마크다운에 처음 표시된 사은품의 구간별 수량입니다.")
        description.setWordWrap(True)
        layout.addWidget(description)

        filters = QHBoxLayout()
        today = datetime.now(KST).date()
        self._start = self._date_edit(today - timedelta(days=29))
        self._end = self._date_edit(today)
        self._start.setAccessibleName("조회 시작일")
        self._end.setAccessibleName("조회 종료일")
        self._all_dates = QCheckBox("전체 기간")
        self._store = QComboBox()
        self._store.addItem("전체 스토어", "")
        for key, name in GIFT_EVENT_STORES.items():
            self._store.addItem(name, key)
        self._refresh = QPushButton("새로고침")
        self._refresh.clicked.connect(self._reload)
        self._start.dateChanged.connect(self._render)
        self._end.dateChanged.connect(self._render)
        self._all_dates.toggled.connect(self._render)
        self._store.currentIndexChanged.connect(self._render)
        for widget in (self._start, QLabel("~"), self._end, self._all_dates, self._store):
            filters.addWidget(widget)
        filters.addStretch()
        filters.addWidget(self._refresh)
        layout.addLayout(filters)

        self._status = QLabel()
        self._status.setTextFormat(Qt.TextFormat.PlainText)
        self._status.setWordWrap(True)
        layout.addWidget(self._status)
        self._table = QTableWidget()
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().hide()
        self._table.horizontalHeader().setMinimumSectionSize(90)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self._table, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.button(QDialogButtonBox.StandardButton.Close).setText("닫기")
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._render()
        QTimer.singleShot(0, self, self._reload)

    @staticmethod
    def _date_edit(value):
        widget = QDateEdit(QDate(value.year, value.month, value.day))
        widget.setDisplayFormat("yyyy-MM-dd")
        widget.setCalendarPopup(True)
        return widget

    def _reload(self):
        self._loading = True
        self._refresh.setEnabled(False)
        self._status.setText("공유 시트에서 읽는 중…")
        try:
            self._records, self._pending_count = run_in_background(self._read_report)
        except Exception as error:
            self._status.setText(f"사은품 기록을 읽지 못했습니다. 관리자에서 Google Sheets 연결을 확인해 주세요.\n{error}")
            self._records = None
            self._table.setRowCount(0)
        finally:
            self._loading = False
            self._refresh.setEnabled(True)
        if self._records is not None:
            self._render()

    def reject(self):
        if not self._loading:
            super().reject()

    def _render(self):
        all_dates = self._all_dates.isChecked()
        self._start.setEnabled(not all_dates)
        self._end.setEnabled(not all_dates)
        count = max((record.tier for record in self._records or ()), default=4)
        count = max(count, 4)
        self._table.setColumnCount(count + 2)
        self._table.setHorizontalHeaderLabels(["날짜", *(f"사은품 {i}" for i in range(1, count + 1)), "합계"])
        if self._records is None:
            return

        start = date.min if all_dates else self._start.date().toPython()
        end = date.max if all_dates else self._end.date().toPython()
        if start > end:
            self._status.setText("시작일을 종료일과 같거나 이전으로 설정해 주세요.")
            self._table.setRowCount(0)
            return

        totals, daily = summarize_gift_records(self._records, start, end, self._store.currentData())
        status = "기록이 있는 날짜를 표시합니다. 수량에 마우스를 올리면 사은품 구성을 볼 수 있습니다."
        if not daily:
            status = "선택한 범위에 사은품 기록이 없습니다."
        if self._pending_count:
            status += f"\n이 PC에 공유 시트 반영 대기 중인 기록 {self._pending_count:,}건이 있습니다."
        self._status.setText(status)
        rows = [("기간 합계", totals), *((day.isoformat(), counts) for day, counts in daily)]
        self._table.setRowCount(len(rows))
        for index, (title, counts) in enumerate(rows):
            details = defaultdict(list)
            per_tier = defaultdict(int)
            for (tier, minimum, gift), quantity in sorted(counts.items()):
                details[tier].append(f"{gift} ({minimum:,}원 이상) · {quantity:,}개")
                per_tier[tier] += quantity
            values = [title, *(f"{per_tier[i]:,}" for i in range(1, count + 1)), f"{sum(counts.values()):,}"]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    item.setToolTip("\n".join(details.get(column, ())))
                if index == 0:
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                self._table.setItem(index, column, item)
