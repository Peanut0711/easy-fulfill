"""마크다운 사은품 집계를 동기화하고 관리자에게 보여 준다."""

from collections import defaultdict
from datetime import date, datetime, timedelta
import logging
from pathlib import Path
from threading import Thread

from PySide6.QtCore import QDate, QObject, QStandardPaths, Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView, QComboBox, QDateEdit, QDialog, QDialogButtonBox, QFileDialog,
    QHBoxLayout, QHeaderView, QLabel, QMessageBox, QPushButton, QSpinBox, QTableWidget,
    QTableWidgetItem, QVBoxLayout,
)

from orders.gift_event import GIFT_EVENT_STORES
from orders.gift_history import KST, MONTHLY_TIER_COUNT, gift_tier_quantities, summarize_gift_month, summarize_gift_records
from orders.gift_history_export import export_gift_month
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
        self._range_mode = QComboBox()
        for label, mode in (("월별", "month"), ("기간 지정", "custom"), ("전체 기간", "all")):
            self._range_mode.addItem(label, mode)
        self._range_mode.setAccessibleName("조회 범위")
        self._year = QSpinBox()
        self._year.setRange(1900, 9999)
        self._year.setSuffix("년")
        self._year.setValue(today.year)
        self._year.setAccessibleName("조회 연도")
        self._month = QComboBox()
        for month in range(1, 13):
            self._month.addItem(f"{month}월", month)
        self._month.setCurrentIndex(today.month - 1)
        self._month.setAccessibleName("조회 월")
        self._month_widgets = (self._year, self._month)
        self._start = self._date_edit(today - timedelta(days=29))
        self._end = self._date_edit(today)
        self._start.setAccessibleName("조회 시작일")
        self._end.setAccessibleName("조회 종료일")
        self._custom_widgets = (self._start, QLabel("~"), self._end)
        self._store = QComboBox()
        self._store.addItem("전체 스토어", "")
        for key, name in GIFT_EVENT_STORES.items():
            self._store.addItem(name, key)
        self._refresh = QPushButton("새로고침")
        self._export = QPushButton("엑셀 저장")
        self._export.setToolTip("월별 조회에서 사은품 1~4의 일별 수량과 월 합계를 저장합니다.")
        self._export.clicked.connect(self._export_month)
        self._refresh.clicked.connect(self._reload)
        self._range_mode.currentIndexChanged.connect(self._render)
        self._year.valueChanged.connect(self._render)
        self._month.currentIndexChanged.connect(self._render)
        self._start.dateChanged.connect(self._render)
        self._end.dateChanged.connect(self._render)
        self._store.currentIndexChanged.connect(self._render)
        for widget in (self._range_mode, *self._month_widgets, *self._custom_widgets, self._store):
            filters.addWidget(widget)
        filters.addStretch()
        filters.addWidget(self._refresh)
        filters.addWidget(self._export)
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
        self._export.setEnabled(False)
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

    def _choose_month_path(self, year, month):
        dialog = QFileDialog(self, "월별 사은품 집계 저장")
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        dialog.setNameFilter("Excel 파일 (*.xlsx)")
        dialog.setDefaultSuffix("xlsx")
        directory = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        dialog.setDirectory(directory or str(Path.home()))
        dialog.selectFile(f"사은품 집계 {year:04d}-{month:02d}.xlsx")
        path = dialog.selectedFiles()[0] if dialog.exec() == QDialog.DialogCode.Accepted else ""
        dialog.deleteLater()
        return path

    def _export_month(self):
        if self._records is None or self._loading or self._range_mode.currentData() != "month":
            return
        year, month, store = self._year.value(), self._month.currentData(), self._store.currentData()
        filename = self._choose_month_path(year, month)
        if not filename:
            return
        path = Path(filename)
        if path.suffix.lower() != ".xlsx":
            QMessageBox.warning(self, self.windowTitle(), "파일 이름 끝을 .xlsx로 입력해 주세요.")
            return
        try:
            export_gift_month(path, self._records, year, month, store)
        except Exception as error:
            QMessageBox.warning(self, self.windowTitle(), f"엑셀을 저장하지 못했습니다.\n{error}")
            return
        QMessageBox.information(self, self.windowTitle(), f"월별 사은품 집계를 저장했습니다.\n{path}")

    def _render(self):
        mode = self._range_mode.currentData()
        for widget in self._month_widgets:
            widget.setVisible(mode == "month")
        for widget in self._custom_widgets:
            widget.setVisible(mode == "custom")
        self._export.setEnabled(mode == "month" and self._records is not None and not self._loading)
        count = MONTHLY_TIER_COUNT
        if mode != "month":
            count = max(count, max((record.tier for record in self._records or ()), default=count))
        self._table.setColumnCount(count + 2)
        self._table.setHorizontalHeaderLabels(["날짜", *(f"사은품 {i}" for i in range(1, count + 1)), "합계"])
        if self._records is None:
            return

        store = self._store.currentData()
        if mode == "month":
            totals, daily = summarize_gift_month(self._records, self._year.value(), self._month.currentData(), store)
            total_label = "월 합계"
            status = "선택한 달의 모든 날짜를 표시합니다. 기록이 없는 날의 수량은 0입니다."
        else:
            start = date.min if mode == "all" else self._start.date().toPython()
            end = date.max if mode == "all" else self._end.date().toPython()
            if start > end:
                self._status.setText("시작일을 종료일과 같거나 이전으로 설정해 주세요.")
                self._table.setRowCount(0)
                return
            totals, daily = summarize_gift_records(self._records, start, end, store)
            total_label = "기간 합계"
            status = "기록이 있는 날짜를 표시합니다."

        if totals:
            status += " 수량에 마우스를 올리면 사은품 구성을 볼 수 있습니다."
        else:
            status = "선택한 범위에 사은품 기록이 없습니다."
        if self._pending_count:
            status += f"\n이 PC에 공유 시트 반영 대기 중인 기록 {self._pending_count:,}건이 있습니다."
        self._status.setText(status)
        rows = [(total_label, totals), *((day.isoformat(), counts) for day, counts in daily)]
        self._table.setRowCount(len(rows))
        for index, (title, counts) in enumerate(rows):
            details = defaultdict(list)
            for (tier, minimum, gift), quantity in sorted(counts.items()):
                details[tier].append(f"{gift} ({minimum:,}원 이상) · {quantity:,}개")
            quantities = gift_tier_quantities(counts, count)
            values = [title, *(f"{quantity:,}" for quantity in quantities), f"{sum(quantities):,}"]
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
