"""PySide6 editor for direct-trade quotation and statement documents."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import QDate, QThread, QUrl, Qt, Signal
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from quotation_statement import (
    DocumentValidationError,
    ItemInput,
    NEGOTIATION_LIMIT,
    NegotiationRequired,
    PAYMENT_METHODS,
    calculate_document,
    export_xlsx_to_pdf,
    generate_document,
)


COL_STATUS = 0
COL_API_NAME = 1
COL_DOCUMENT_NAME = 2
COL_SPEC = 3
COL_QUANTITY = 4
COL_PRICE = 5
ROLE_API_PRICE = Qt.ItemDataRole.UserRole + 1
SMARTSTORE_PRODUCT_URL_TEMPLATE = "https://smartstore.naver.com/higenis/products/{product_no}"


class ProductSearchThread(QThread):
    result_ready = Signal(dict)

    def __init__(self, worker, query: str, parent=None):
        super().__init__(parent)
        self._worker = worker
        self._query = query

    def run(self):
        try:
            result = self._worker(self._query)
            if not isinstance(result, dict):
                result = {"ok": False, "error": "상품 조회 결과 형식이 올바르지 않습니다."}
        except Exception as exc:
            result = {"ok": False, "error": str(exc)}
        self.result_ready.emit(result)


class PdfExportThread(QThread):
    result_ready = Signal(dict)

    def __init__(self, xlsx_path, parent=None):
        super().__init__(parent)
        self._xlsx_path = Path(xlsx_path)

    def run(self):
        try:
            self.result_ready.emit({"ok": True, "path": str(export_xlsx_to_pdf(self._xlsx_path))})
        except Exception as exc:
            self.result_ready.emit({"ok": False, "error": str(exc)})


class ProductChoiceDialog(QDialog):
    def __init__(self, products, query: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("스마트스토어 상품 선택")
        self.resize(920, 520)
        self._products = list(products)
        self.selected_product = None

        layout = QVBoxLayout(self)
        guide = QLabel("상품명과 가격을 확인한 뒤 행을 선택하고 ‘적용’을 누르세요. 상품명은 더블클릭해도 바로 적용됩니다.")
        guide.setWordWrap(True)
        layout.addWidget(guide)

        self.filter_edit = QLineEdit(query)
        self.filter_edit.setPlaceholderText("결과 안에서 다시 검색")
        layout.addWidget(self.filter_edit)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["상품번호", "스마트스토어 상품명", "기준가격", "상품 페이지"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.table.setStyleSheet(
            "QTableWidget::item:selected { background-color: #eef1f4; color: #111827; }"
        )
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table)

        buttons = QDialogButtonBox()
        self.apply_button = buttons.addButton("적용", QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton("취소", QDialogButtonBox.ButtonRole.RejectRole)
        self.apply_button.setEnabled(False)
        layout.addWidget(buttons)

        self.filter_edit.textChanged.connect(self._populate)
        self.table.itemSelectionChanged.connect(
            lambda: self.apply_button.setEnabled(self.table.currentRow() >= 0)
        )
        self.table.cellClicked.connect(self._open_product_link)
        self.table.cellDoubleClicked.connect(self._apply_from_product_name)
        buttons.accepted.connect(self._accept_selection)
        buttons.rejected.connect(self.reject)
        self._populate(query)

    @staticmethod
    def _matches(product, query: str) -> bool:
        tokens = [token.casefold() for token in query.split() if token]
        name = str(product.get("name") or "").casefold()
        return all(token in name for token in tokens)

    def _populate(self, query: str):
        matches = [p for p in self._products if self._matches(p, query)]
        self.table.setRowCount(len(matches))
        for row, product in enumerate(matches):
            number = QTableWidgetItem(str(product.get("product_no") or ""))
            number.setData(Qt.ItemDataRole.UserRole, product)
            self.table.setItem(row, 0, number)
            self.table.setItem(row, 1, QTableWidgetItem(str(product.get("name") or "")))
            price = int(product.get("price") or 0)
            price_item = QTableWidgetItem(f"{price:,}원")
            price_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row, 2, price_item)
            product_url = self._smartstore_product_url(product)
            link_item = QTableWidgetItem("열기" if product_url else "-")
            link_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if product_url:
                link_item.setData(Qt.ItemDataRole.UserRole, product_url)
                link_item.setToolTip(product_url)
                link_font = link_item.font()
                link_font.setUnderline(True)
                link_item.setFont(link_font)
                link_item.setForeground(QColor("#2563eb"))
            self.table.setItem(row, 3, link_item)
        self.table.clearSelection()
        self.table.setCurrentCell(-1, -1)
        self.apply_button.setEnabled(False)

    @staticmethod
    def _smartstore_product_url(product):
        product_no = str(product.get("product_no") or "").strip()
        if not product_no.isdigit():
            return ""
        return SMARTSTORE_PRODUCT_URL_TEMPLATE.format(product_no=product_no)

    def _open_product_link(self, row, column):
        if column != 3:
            return
        url = self.table.item(row, column).data(Qt.ItemDataRole.UserRole)
        if url and not QDesktopServices.openUrl(QUrl(str(url))):
            QMessageBox.warning(self, "상품 페이지", "기본 브라우저에서 상품 페이지를 열 수 없습니다.")

    def _apply_from_product_name(self, row, column):
        if column == 1:
            self._accept_selection()

    def _accept_selection(self):
        row = self.table.currentRow()
        if row < 0:
            return
        self.selected_product = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        self.accept()


class QuotationStatementWidget(QWidget):
    def __init__(self, product_search_worker=None, on_created=None, base_dir=None, parent=None):
        super().__init__(parent)
        self._product_search_worker = product_search_worker
        self._on_created = on_created
        self._base_dir = Path(base_dir or Path(__file__).resolve().parent)
        self._search_thread = None
        self._pdf_thread = None
        self._pdf_progress = None
        self._prefetch_thread = None
        self._prefetch_result = None
        self._prefetch_complete = False
        self._pending_search = None
        self._search_target_row = -1
        self._updating_table = False
        self._negotiation_mode = False
        self._build_ui()
        self.add_item_row()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(10)

        info_group = QGroupBox("문서 기본정보")
        info_layout = QHBoxLayout(info_group)
        form = QFormLayout()
        self.document_type = QComboBox()
        self.document_type.addItems(["견적서", "거래명세서"])
        self.payment_method = QComboBox()
        self.payment_method.addItems(PAYMENT_METHODS)
        self.delivery_label = QLabel("납기")
        self.delivery_edit = QLineEdit()
        self.delivery_edit.setPlaceholderText("결제 후 즉시")
        self.organization_edit = QLineEdit()
        self.organization_edit.setPlaceholderText("예: 한국대학교 전자공학과")
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("예: 홍길동")
        self.trade_date = QDateEdit(QDate.currentDate())
        self.trade_date.setCalendarPopup(True)
        self.trade_date.setDisplayFormat("yyyy.MM.dd")
        form.addRow("문서 종류", self.document_type)
        form.addRow("소속", self.organization_edit)
        form.addRow("성명", self.name_edit)
        form.addRow("거래일자", self.trade_date)
        form.addRow("결제 방법", self.payment_method)
        form.addRow(self.delivery_label, self.delivery_edit)
        info_layout.addLayout(form, 1)

        search_layout = QVBoxLayout()
        search_layout.addWidget(QLabel("스마트스토어 기준상품 찾기 (선택한 품목 행에 적용)"))
        search_row = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("예: stm32 h7 개발보드")
        self.search_edit.setFixedWidth(350)
        self.search_button = QPushButton("상품 찾기")
        self.search_button.setFixedWidth(82)
        search_row.addWidget(self.search_edit)
        search_row.addWidget(self.search_button)
        search_layout.addLayout(search_row)
        self.search_status = QLabel("API를 사용하지 않아도 아래 표에 상품명과 가격을 직접 입력할 수 있습니다.")
        self.search_status.setWordWrap(True)
        search_layout.addWidget(self.search_status)
        search_layout.addStretch(1)
        info_layout.addLayout(search_layout, 2)
        root.addWidget(info_group)

        item_header = QHBoxLayout()
        item_header.addWidget(QLabel("품목"))
        item_header.addStretch(1)
        self.add_button = QPushButton("품목 추가")
        self.delete_button = QPushButton("선택 품목 삭제")
        item_header.addWidget(self.add_button)
        item_header.addWidget(self.delete_button)
        root.addLayout(item_header)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["상태", "API 상품명 (확인용)", "문서용 상품명", "규격", "수량", "VAT포함 단가"]
        )
        for column in range(self.table.columnCount()):
            self.table.horizontalHeaderItem(column).setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(
            "QTableWidget { selection-background-color: #dcefff; selection-color: #1f2937; }"
            "QTableWidget::item:selected { background-color: #dcefff; color: #1f2937; }"
        )
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(COL_STATUS, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(COL_API_NAME, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(COL_DOCUMENT_NAME, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(COL_SPEC, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(COL_PRICE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(COL_QUANTITY, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setColumnWidth(COL_DOCUMENT_NAME, 240)
        self.table.setColumnWidth(COL_SPEC, 120)
        self.table.setMinimumHeight(260)
        root.addWidget(self.table, 1)

        summary_group = QGroupBox("계산 요약")
        summary_layout = QVBoxLayout(summary_group)
        self.summary_label = QLabel("품목을 입력하면 금액을 계산합니다.")
        self.summary_label.setWordWrap(True)
        self.summary_label.setStyleSheet("font-weight: 600; padding: 4px;")
        self.negotiated_check = QCheckBox("협의 완료 (현재 입력 단가를 협의 확정가로 사용)")
        self.negotiated_check.setVisible(False)
        self.free_shipping_check = QCheckBox("배송비 면제")
        summary_layout.addWidget(self.summary_label)
        summary_layout.addWidget(self.negotiated_check)
        summary_layout.addWidget(self.free_shipping_check)
        root.addWidget(summary_group)

        action_row = QHBoxLayout()
        action_row.addStretch(1)
        self.reset_button = QPushButton("초기화")
        self.create_button = QPushButton("엑셀 만들기")
        self.create_button.setMinimumWidth(120)
        self.pdf_button = QPushButton("PDF 만들기")
        self.pdf_button.setMinimumWidth(120)
        action_row.addWidget(self.reset_button)
        action_row.addWidget(self.create_button)
        action_row.addWidget(self.pdf_button)
        root.addLayout(action_row)

        self.add_button.clicked.connect(self.add_item_row)
        self.delete_button.clicked.connect(self.delete_selected_row)
        self.search_button.clicked.connect(self.search_products)
        self.search_edit.returnPressed.connect(self.search_products)
        self.table.itemChanged.connect(self._on_item_changed)
        self.document_type.currentTextChanged.connect(self._on_document_type_changed)
        self.negotiated_check.toggled.connect(self.refresh_summary)
        self.free_shipping_check.toggled.connect(self.refresh_summary)
        self.reset_button.clicked.connect(self.reset_form)
        self.create_button.clicked.connect(self.create_document)
        self.pdf_button.clicked.connect(lambda: self.create_document(as_pdf=True))

    def _on_document_type_changed(self, document_type):
        quote = document_type == "견적서"
        self.payment_method.setEnabled(quote)
        self.payment_method.setToolTip("" if quote else "결제 방법은 견적서에만 기재됩니다.")
        self.delivery_label.setVisible(quote)
        self.delivery_edit.setVisible(quote)
        self.refresh_summary()

    @staticmethod
    def _editable_item(text=""):
        return QTableWidgetItem(str(text))

    def add_item_row(self):
        self._updating_table = True
        try:
            row = self.table.rowCount()
            self.table.insertRow(row)
            api_item = QTableWidgetItem("")
            api_item.setFlags(api_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, COL_API_NAME, api_item)
            self.table.setItem(row, COL_DOCUMENT_NAME, self._editable_item())
            specification = self._editable_item()
            specification.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, COL_SPEC, specification)
            price = self._editable_item()
            price.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, COL_PRICE, price)
            quantity = self._editable_item("1")
            quantity.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, COL_QUANTITY, quantity)
            status = QTableWidgetItem("수동")
            status.setFlags(status.flags() & ~Qt.ItemFlag.ItemIsEditable)
            status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, COL_STATUS, status)
            self.table.setCurrentCell(row, COL_DOCUMENT_NAME)
        finally:
            self._updating_table = False
        self.refresh_summary()

    def delete_selected_row(self):
        row = self.table.currentRow()
        if row < 0:
            return
        self.table.removeRow(row)
        if self.table.rowCount() == 0:
            self.add_item_row()
        self.refresh_summary()

    def _on_item_changed(self, item):
        if self._updating_table:
            return
        if item.column() == COL_PRICE:
            raw_price = item.text().replace(",", "").replace("원", "").strip()
            if raw_price.isdigit():
                formatted = f"{int(raw_price):,}원"
                if item.text() != formatted:
                    self._updating_table = True
                    item.setText(formatted)
                    self._updating_table = False
        row = item.row()
        self._update_row_status(row)
        self.refresh_summary()

    def _update_row_status(self, row):
        api_name = self.table.item(row, COL_API_NAME).text().strip()
        status_text = "수동"
        if api_name:
            doc_name = self.table.item(row, COL_DOCUMENT_NAME).text().strip()
            specification = self.table.item(row, COL_SPEC).text().strip()
            price_text = self.table.item(row, COL_PRICE).text().replace(",", "").replace("원", "").strip()
            api_price = self.table.item(row, COL_PRICE).data(ROLE_API_PRICE)
            modified = (
                doc_name != api_name
                or bool(specification)
                or (api_price is not None and price_text != str(api_price))
            )
            status_text = "수정" if modified else "자동"
        if self.table.item(row, COL_STATUS).text() != status_text:
            self._updating_table = True
            self.table.item(row, COL_STATUS).setText(status_text)
            self._updating_table = False

    def _collect_items(self, strict=True):
        items = []
        for row in range(self.table.rowCount()):
            name = self.table.item(row, COL_DOCUMENT_NAME).text().strip()
            spec = self.table.item(row, COL_SPEC).text().strip()
            price = self.table.item(row, COL_PRICE).text().replace(",", "").replace("원", "").strip()
            quantity = self.table.item(row, COL_QUANTITY).text().strip()
            if not (name or price or spec) and quantity in ("", "1"):
                continue
            if not strict and (not name or not price or not quantity.isdigit()):
                continue
            items.append(
                ItemInput(
                    api_name=self.table.item(row, COL_API_NAME).text().strip(),
                    document_name=name,
                    specification=spec,
                    gross_unit_price=price,
                    quantity=quantity,
                )
            )
        return items

    def refresh_summary(self):
        items = self._collect_items(strict=False)
        if not items:
            self.summary_label.setText("품목을 입력하면 금액을 계산합니다.")
            self.negotiated_check.setVisible(False)
            return
        try:
            raw_result = calculate_document(items, negotiated=True)
            over_limit = raw_result.goods_total_before_discount > NEGOTIATION_LIMIT
            if over_limit:
                self._negotiation_mode = True
            elif self._negotiation_mode and not self.negotiated_check.isChecked():
                self._negotiation_mode = False

            if self._negotiation_mode:
                result = raw_result
            else:
                result = calculate_document(items, free_shipping=self.free_shipping_check.isChecked())
        except DocumentValidationError as exc:
            self.summary_label.setText(str(exc))
            return
        self.negotiated_check.setVisible(self._negotiation_mode)
        discount = "별도 협의" if self._negotiation_mode else f"{int(result.discount_rate * Decimal('100'))}%"
        self.summary_label.setText(
            f"할인 전 품목 합계 {result.goods_total_before_discount:,}원  |  할인 {discount} "
            f"({result.discount_amount:,}원)  |  공급가액 {result.supply_total:,}원  |  "
            f"세액 {result.tax_total:,}원  |  배송비 {result.shipping_gross:,}원  |  "
            f"최종 합계 {result.grand_total:,}원"
        )

    def search_products(self):
        query = self.search_edit.text().strip()
        row = self.table.currentRow()
        if not query:
            QMessageBox.information(self, "상품 찾기", "상품명 또는 짧은 키워드를 입력해주세요.")
            return
        if row < 0:
            QMessageBox.information(self, "상품 찾기", "상품을 적용할 품목 행을 먼저 선택해주세요.")
            return
        if self._product_search_worker is None:
            QMessageBox.information(self, "상품 찾기", "상품 API를 사용할 수 없습니다. 수동 입력은 계속 사용할 수 있습니다.")
            return
        if self._prefetch_thread and self._prefetch_thread.isRunning():
            self._pending_search = (query, row)
            self.search_status.setText("상품 목록을 준비 중입니다. 완료되면 검색합니다…")
            return
        if self._search_thread and self._search_thread.isRunning():
            return
        self._start_product_search(query, row)

    def prefetch_products(self):
        """문서 탭 첫 진입 시 상품 목록을 미리 세션 캐시에 적재한다."""
        if self._product_search_worker is None:
            return
        if self._prefetch_complete or (self._prefetch_thread and self._prefetch_thread.isRunning()):
            return
        self._prefetch_result = None
        self.search_status.setText("스마트스토어 상품 목록을 준비하고 있습니다…")
        self._prefetch_thread = ProductSearchThread(self._product_search_worker, "", self)
        self._prefetch_thread.result_ready.connect(self._on_prefetch_result)
        self._prefetch_thread.finished.connect(self._on_prefetch_finished)
        self._prefetch_thread.start()

    def _start_product_search(self, query, row):
        self._search_target_row = row
        self.search_button.setEnabled(False)
        self.search_status.setText("스마트스토어 상품을 조회하고 있습니다…")
        self._search_thread = ProductSearchThread(self._product_search_worker, query, self)
        self._search_thread.result_ready.connect(lambda result: self._on_search_result(query, result))
        self._search_thread.finished.connect(lambda: self.search_button.setEnabled(True))
        self._search_thread.start()

    def _on_prefetch_result(self, result):
        self._prefetch_result = result

    def _on_prefetch_finished(self):
        result = self._prefetch_result or {"ok": False, "error": "상품 목록을 불러오지 못했습니다."}
        self._prefetch_thread = None
        if not result.get("ok"):
            self._pending_search = None
            self.search_status.setText(
                f"상품 목록 준비 실패: {result.get('error', '알 수 없는 오류')} — 수동 입력은 가능합니다."
            )
            return
        count = int(result.get("count") or 0)
        self._prefetch_complete = True
        self.search_status.setText(f"스마트스토어 상품 {count:,}개 준비 완료")
        pending = self._pending_search
        self._pending_search = None
        if pending:
            self._start_product_search(*pending)

    def _on_search_result(self, query, result):
        if not result.get("ok"):
            error = result.get("error") or "알 수 없는 오류"
            self.search_status.setText(f"상품 API 조회 실패: {error} — 수동 입력은 계속 사용할 수 있습니다.")
            return
        products = result.get("products") or []
        if not products:
            self.search_status.setText("일치하는 상품이 없습니다. 검색어를 줄이거나 수동으로 입력해주세요.")
            return
        self.search_status.setText(f"검색 결과 {len(products)}건 — 적용 전까지 기존 입력은 유지됩니다.")
        dialog = ProductChoiceDialog(products, query, self)
        if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.selected_product:
            return
        if self._search_target_row >= self.table.rowCount():
            return
        self._apply_product(self._search_target_row, dialog.selected_product)

    def _apply_product(self, row, product):
        name = str(product.get("name") or "").strip()
        price = int(product.get("price") or 0)
        self._updating_table = True
        try:
            self.table.item(row, COL_API_NAME).setText(name)
            self.table.item(row, COL_DOCUMENT_NAME).setText(name)
            price_item = self.table.item(row, COL_PRICE)
            price_item.setText(f"{price:,}원")
            price_item.setData(ROLE_API_PRICE, price)
        finally:
            self._updating_table = False
        self._update_row_status(row)
        self.refresh_summary()

    def reset_form(self):
        confirmation = QMessageBox(QMessageBox.Icon.Question, "초기화", "입력한 문서 내용을 모두 지울까요?", parent=self)
        yes_button = confirmation.addButton("예", QMessageBox.ButtonRole.YesRole)
        confirmation.addButton("아니오", QMessageBox.ButtonRole.NoRole)
        confirmation.exec()
        if confirmation.clickedButton() is not yes_button:
            return
        self.organization_edit.clear()
        self.name_edit.clear()
        self.trade_date.setDate(QDate.currentDate())
        self.payment_method.setCurrentIndex(0)
        self.delivery_edit.clear()
        self.search_edit.clear()
        self.search_status.setText("API를 사용하지 않아도 아래 표에 상품명과 가격을 직접 입력할 수 있습니다.")
        self.negotiated_check.setChecked(False)
        self.free_shipping_check.setChecked(False)
        self._negotiation_mode = False
        self.table.setRowCount(0)
        self.add_item_row()

    def create_document(self, as_pdf=False):
        try:
            items = self._collect_items(strict=True)
            negotiated = self._negotiation_mode and self.negotiated_check.isChecked()
            free_shipping = self.free_shipping_check.isChecked()
            if self._negotiation_mode and not negotiated:
                raise NegotiationRequired("500만 원 초과 주문은 협의 완료를 확인해야 합니다.")
            result = calculate_document(
                items, negotiated=negotiated, free_shipping=free_shipping
            )
            organization = self.organization_edit.text().strip()
            name = self.name_edit.text().strip()
            if not organization or not name:
                raise DocumentValidationError("소속명과 성명을 입력해주세요.")
        except DocumentValidationError as exc:
            QMessageBox.warning(self, "입력 확인", str(exc))
            return

        lines = "\n".join(
            f"- {item.document_name} / {item.specification or '-'} / {item.quantity}개 / 단가 {item.gross_unit_price:,}원"
            for item in result.items
        )
        output_label = "엑셀과 PDF를" if as_pdf else "엑셀을"
        text = (
            f"{self.document_type.currentText()} {output_label} 생성할까요?\n\n{lines}\n\n"
            f"공급가액 {result.supply_total:,}원 + 세액 {result.tax_total:,}원 = "
            f"최종 {result.grand_total:,}원"
        )
        if self.document_type.currentText() == "견적서":
            text += f"\n결제 방법: {PAYMENT_METHODS[self.payment_method.currentText()]}"
        confirmation = QMessageBox(QMessageBox.Icon.Question, "문서 생성 확인", text, parent=self)
        yes_button = confirmation.addButton("예", QMessageBox.ButtonRole.YesRole)
        confirmation.addButton("아니오", QMessageBox.ButtonRole.NoRole)
        confirmation.exec()
        if confirmation.clickedButton() is not yes_button:
            return

        try:
            path, _ = generate_document(
                self.document_type.currentText(),
                organization,
                name,
                self.trade_date.date().toPython(),
                items,
                negotiated=negotiated,
                free_shipping=free_shipping,
                payment_method=self.payment_method.currentText(),
                delivery_term=self.delivery_edit.text(),
                templates_dir=self._base_dir / "templates",
                output_dir=self._base_dir / "output",
            )
        except Exception as exc:
            QMessageBox.critical(self, "문서 생성 실패", str(exc))
            return

        if as_pdf:
            self._start_pdf_export(path, self.document_type.currentText())
            return

        if self._on_created:
            self._on_created(str(path), f"{self.document_type.currentText()} 파일이 생성되었습니다.")
        else:
            QMessageBox.information(self, "완료", f"문서를 생성했습니다.\n{path}")

    def _start_pdf_export(self, xlsx_path, document_type):
        self.create_button.setEnabled(False)
        self.pdf_button.setEnabled(False)
        progress = QProgressDialog("PDF를 변환하고 있습니다. 잠시만 기다려주세요.", None, 0, 0, self)
        progress.setWindowTitle("PDF 만들기")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, False)
        progress.setCancelButton(None)
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.show()
        self._pdf_progress = progress
        self._pdf_thread = PdfExportThread(xlsx_path, self)
        self._pdf_thread.result_ready.connect(
            lambda result: self._finish_pdf_export(result, xlsx_path, document_type)
        )
        self._pdf_thread.finished.connect(self._pdf_thread.deleteLater)
        self._pdf_thread.start()

    def _finish_pdf_export(self, result, xlsx_path, document_type):
        if self._pdf_progress:
            self._pdf_progress.close()
            self._pdf_progress.deleteLater()
            self._pdf_progress = None
        self._pdf_thread = None
        self.create_button.setEnabled(True)
        self.pdf_button.setEnabled(True)
        if not result.get("ok"):
            QMessageBox.warning(
                self,
                "PDF 출력 실패",
                f"PDF 변환에 실패했습니다. 엑셀 파일은 생성되었습니다.\n\n{result.get('error', '알 수 없는 오류')}",
            )
            if self._on_created:
                self._on_created(str(xlsx_path), f"{document_type} 엑셀 파일이 생성되었습니다.")
            return
        pdf_path = result["path"]
        if self._on_created:
            self._on_created(str(pdf_path), f"{document_type} 엑셀 및 PDF 파일이 생성되었습니다.", "PDF 열기")
        else:
            QMessageBox.information(self, "완료", f"엑셀 및 PDF 파일을 생성했습니다.\n{pdf_path}")
