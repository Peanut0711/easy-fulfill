"""구매 금액별 사은품 공유 설정 편집기."""

import re

from PySide6.QtCore import Qt, QThread
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
        self.resize(900, 420)

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

        self._table = QTableWidget(0, 5)
        self._table.setHorizontalHeaderLabels([
            "최소 구매 금액 (원)", "지급 사은품", "송장 표기", "마크다운 문구", "아이콘만",
        ])
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self._table.horizontalHeaderItem(2).setToolTip("대괄호는 자동으로 붙습니다. 비워 두면 송장에 표시하지 않습니다.")
        self._table.horizontalHeaderItem(3).setToolTip("🎁 뒤에 붙일 문구입니다. 비워 두면 지급 사은품 이름을 표시합니다.")
        self._table.horizontalHeaderItem(4).setToolTip("선택하면 마크다운에 🎁만 표시합니다.")
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._table.itemChanged.connect(self._update_markdown_editor)
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

    def _append_tier(self, minimum="", gift="", invoice_label="", markdown_label=None):
        row = self._table.rowCount()
        self._table.insertRow(row)
        self._table.setItem(row, 0, QTableWidgetItem(minimum))
        self._table.setItem(row, 1, QTableWidgetItem(gift))
        self._table.setItem(row, 2, QTableWidgetItem(invoice_label))
        self._table.setItem(row, 3, QTableWidgetItem(markdown_label or ""))
        icon_only = QTableWidgetItem()
        icon_only.setFlags(icon_only.flags() & ~Qt.ItemFlag.ItemIsEditable)
        icon_only.setCheckState(Qt.CheckState.Checked if markdown_label == "" else Qt.CheckState.Unchecked)
        self._table.setItem(row, 4, icon_only)
        return row

    def _add_tier(self):
        row = self._append_tier()
        self._table.setCurrentCell(row, 0)
        self._table.editItem(self._table.item(row, 0))

    def _update_markdown_editor(self, item):
        if item.column() == 4:
            label = self._table.item(item.row(), 3)
            if item.checkState() == Qt.CheckState.Checked:
                label.setFlags(label.flags() & ~Qt.ItemFlag.ItemIsEnabled)
            else:
                label.setFlags(label.flags() | Qt.ItemFlag.ItemIsEnabled)

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
            invoice_labels = dict(event.invoice_labels)
            markdown_labels = dict(event.markdown_labels)
            for minimum, gift in event.tiers:
                self._append_tier(f"{minimum:,}", gift, invoice_labels.get(minimum, ""), markdown_labels.get(minimum))
            self._set_editable(True)
            self._status.setText(
                "주문 줄의 원 단위 금액으로 충족한 가장 높은 구간을 적용합니다.\n"
                "마크다운 문구가 비어 있으면 지급 사은품 이름을 쓰고, ‘아이콘만’을 선택하면 🎁만 표시합니다."
            )
        else:
            self.accept()

    def _save_settings(self):
        try:
            tiers = []
            invoice_labels = []
            markdown_labels = []
            for row in range(self._table.rowCount()):
                amount = self._table.item(row, 0).text().strip()
                if not re.fullmatch(r"[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+", amount):
                    raise ValueError(f"{row + 1}번째 구간의 최소 구매 금액을 원 단위 정수로 입력해 주세요.")
                minimum = int(amount.replace(",", ""))
                tiers.append((minimum, self._table.item(row, 1).text().strip()))
                invoice_label = self._table.item(row, 2).text().strip()
                if invoice_label:
                    invoice_labels.append((minimum, invoice_label))
                markdown_label = self._table.item(row, 3).text().strip()
                if self._table.item(row, 4).checkState() == Qt.CheckState.Checked:
                    markdown_labels.append((minimum, ""))
                elif markdown_label:
                    markdown_labels.append((minimum, markdown_label))
            stores = tuple(store for store, checkbox in self._stores.items() if checkbox.isChecked())
            event = GiftEvent(self._enabled.isChecked(), stores, tuple(tiers), tuple(invoice_labels), tuple(markdown_labels))
        except ValueError as error:
            QMessageBox.warning(self, self.windowTitle(), str(error))
            return

        original = self._original
        self._status.setText("공유 스프레드시트에 저장하는 중입니다…")
        self._start_sheet_task(lambda: save_gift_event(self._open_worksheet(), event, original))

    def reject(self):
        if self._thread is None:
            super().reject()
