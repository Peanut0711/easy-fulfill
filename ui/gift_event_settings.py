"""구매 금액별 사은품 공유 설정 편집기."""

import re

from PySide6.QtCore import QThread
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QDialog, QDialogButtonBox, QHBoxLayout, QHeaderView,
    QLabel, QMessageBox, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
)

from orders.gift_event import GIFT_EVENT_STORES, GiftEvent, read_gift_event, save_gift_event


class _SheetThread(QThread):
    def __init__(self, operation, parent):
        super().__init__(parent)
        self.operation = operation
        self.result = None
        self.error = None

    def run(self):
        try:
            self.result = self.operation()
        except Exception as error:
            self.error = str(error)


class GiftEventSettingsDialog(QDialog):
    def __init__(self, open_worksheet, parent=None):
        super().__init__(parent)
        self._open_worksheet = open_worksheet
        self._original = None
        self._thread = None
        self.setWindowTitle("구매 금액별 사은품 공유 설정")
        self.resize(640, 420)

        layout = QVBoxLayout(self)
        self._status = QLabel("공유 스프레드시트에서 설정을 읽는 중입니다…")
        self._status.setWordWrap(True)
        layout.addWidget(self._status)
        self._enabled = QCheckBox("이벤트 사용")
        layout.addWidget(self._enabled)

        stores = QHBoxLayout()
        stores.addWidget(QLabel("적용 스토어"))
        self._stores = {}
        for store, name in GIFT_EVENT_STORES.items():
            checkbox = QCheckBox(name)
            self._stores[store] = checkbox
            stores.addWidget(checkbox)
        stores.addStretch()
        layout.addLayout(stores)

        self._table = QTableWidget(0, 2)
        self._table.setHorizontalHeaderLabels(["최소 구매 금액 (원)", "지급 사은품"])
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        layout.addWidget(self._table)

        rows = QHBoxLayout()
        self._add = QPushButton("구간 추가")
        self._remove = QPushButton("선택 구간 삭제")
        self._add.clicked.connect(self._add_tier)
        self._remove.clicked.connect(self._remove_tier)
        self._table.itemSelectionChanged.connect(self._update_remove_button)
        rows.addWidget(self._add)
        rows.addWidget(self._remove)
        rows.addStretch()
        layout.addLayout(rows)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        self._save = buttons.button(QDialogButtonBox.StandardButton.Save)
        self._save.setText("공유 설정 저장")
        self._cancel = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        self._cancel.setText("취소")
        buttons.accepted.connect(self._save_settings)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._start_sheet_task(lambda: read_gift_event(self._open_worksheet()))

    def _append_tier(self, minimum="", gift=""):
        row = self._table.rowCount()
        self._table.insertRow(row)
        self._table.setItem(row, 0, QTableWidgetItem(minimum))
        self._table.setItem(row, 1, QTableWidgetItem(gift))
        return row

    def _add_tier(self):
        row = self._append_tier()
        self._table.setCurrentCell(row, 0)
        self._table.editItem(self._table.item(row, 0))

    def _remove_tier(self):
        selected = self._table.selectionModel().selectedRows()
        if selected:
            self._table.removeRow(selected[0].row())

    def _update_remove_button(self):
        self._remove.setEnabled(self._table.isEnabled() and bool(self._table.selectionModel().selectedRows()))

    def _set_editable(self, editable):
        for widget in (self._enabled, *self._stores.values(), self._table, self._add, self._save):
            widget.setEnabled(editable)
        self._update_remove_button()

    def _start_sheet_task(self, operation):
        self._set_editable(False)
        self._cancel.setEnabled(False)
        self._thread = _SheetThread(operation, self)
        self._thread.finished.connect(self._finish_sheet_task)
        self._thread.start()

    def _finish_sheet_task(self):
        thread = self._thread
        self._thread = None
        thread.deleteLater()
        self._cancel.setEnabled(True)
        if thread.error is not None:
            self._set_editable(self._original is not None)
            self._status.setText("공유 설정을 처리하지 못했습니다.")
            QMessageBox.warning(self, self.windowTitle(), thread.error)
        elif self._original is None:
            event, self._original = thread.result
            self._enabled.setChecked(event.enabled)
            for store, checkbox in self._stores.items():
                checkbox.setChecked(store in event.stores)
            for minimum, gift in event.tiers:
                self._append_tier(f"{minimum:,}", gift)
            self._set_editable(True)
            self._status.setText("주문 줄의 원 단위 금액으로 충족한 가장 높은 구간을 적용합니다.")
        else:
            self.accept()

    def _save_settings(self):
        try:
            tiers = []
            for row in range(self._table.rowCount()):
                amount = self._table.item(row, 0).text().strip()
                if not re.fullmatch(r"[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+", amount):
                    raise ValueError(f"{row + 1}번째 구간의 최소 구매 금액을 원 단위 정수로 입력해 주세요.")
                tiers.append((int(amount.replace(",", "")), self._table.item(row, 1).text().strip()))
            stores = tuple(store for store, checkbox in self._stores.items() if checkbox.isChecked())
            event = GiftEvent(self._enabled.isChecked(), stores, tuple(tiers))
        except ValueError as error:
            QMessageBox.warning(self, self.windowTitle(), str(error))
            return

        original = self._original
        self._status.setText("공유 스프레드시트에 저장하는 중입니다…")
        self._start_sheet_task(lambda: save_gift_event(self._open_worksheet(), event, original))

    def reject(self):
        if self._thread is None:
            super().reject()
