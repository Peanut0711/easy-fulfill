import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import load_workbook
from PySide6.QtCore import QDate
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox

from documents.quotation_statement import generate_document
from documents.quotation_statement_widget import QuotationStatementWidget


class OrderImportWidgetTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.worker = Mock(return_value={
            "ok": True,
            "order": {
                "order_id": "123456789",
                "customer_name": "테스트 고객",
                "payment_date": "2026-10-26",
                "shipping_gross": 2500,
                "items": [{
                    "name": "주문 상품", "specification": "옵션 A",
                    "quantity": 3, "gross_unit_price": 40000,
                    "gross_amount": 119999,
                }],
            },
        })
        self.widget = QuotationStatementWidget(
            order_lookup_worker=self.worker,
            coupang_order_lookup_worker=self.worker,
            on_created=Mock(),
        )

    def tearDown(self):
        thread = self.widget._order_lookup_thread
        if thread is not None:
            thread.wait(5000)
        self.widget.close()
        self.widget.deleteLater()
        self.app.processEvents()

    def load_order(self, source):
        getattr(self.widget, f"{source}_order_radio").click()
        getattr(self.widget, f"{source}_order_id_edit").setText("123456789")
        getattr(self.widget, f"{source}_order_load_button").click()
        deadline = time.monotonic() + 5
        while self.widget._order_lookup_thread is not None and time.monotonic() < deadline:
            QTest.qWait(10)
        self.assertIsNone(self.widget._order_lookup_thread)
        self.assertTrue(getattr(self.widget, f"{source}_order_load_button").isEnabled())
        self.worker.assert_called_with("123456789")

    def test_input_modes_are_available_for_both_document_types(self):
        for document_type in ("견적서", "거래명세서"):
            self.widget.document_type.setCurrentText(document_type)
            self.assertFalse(self.widget.input_mode_label.isHidden())
            self.assertFalse(self.widget.input_mode_widget.isHidden())
            for mode in ("manual", "naver", "coupang"):
                with self.subTest(document_type=document_type, mode=mode):
                    button = (self.widget.direct_input_radio if mode == "manual"
                              else getattr(self.widget, f"{mode}_order_radio"))
                    button.click()
                    self.assertEqual(self.widget.product_search_widget.isHidden(), mode != "manual")
                    self.assertEqual(self.widget.naver_order_group.isHidden(), mode != "naver")
                    self.assertEqual(self.widget.coupang_order_group.isHidden(), mode != "coupang")
                    self.assertEqual(self.widget.discount_controls.isHidden(), mode != "manual")

    def test_quote_import_preserves_amounts_through_type_switch_and_export(self):
        for source in ("naver", "coupang"):
            with self.subTest(source=source):
                self.widget._clear_order_import()
                self.load_order(source)
                self.assertEqual(self.widget.name_edit.text(), "테스트 고객")
                self.assertEqual(self.widget.trade_date.date(), QDate(2026, 10, 26))
                self.assertEqual(self.widget._collect_items()[0].specification, "옵션 A")
                for document_type in ("견적서", "거래명세서", "견적서"):
                    self.widget.document_type.setCurrentText(document_type)
                    self.assertIn("최종 합계 122,499원", self.widget.summary_label.text())
                    self.assertIn("추가 할인 없음", self.widget.summary_label.text())
                    self.assertTrue(self.widget.discount_controls.isHidden())
                    self.assertTrue(self.widget.free_shipping_check.isHidden())

                with tempfile.TemporaryDirectory() as directory:
                    def export(*args, **kwargs):
                        kwargs["output_dir"] = directory
                        return generate_document(*args, **kwargs)

                    with patch("documents.quotation_statement_widget.QMessageBox") as message_box, \
                            patch("documents.quotation_statement_widget.generate_document", side_effect=export):
                        dialog = message_box.return_value
                        dialog.clickedButton.return_value = dialog.addButton.return_value
                        self.widget.create_document()
                        message_box.warning.assert_not_called()
                        message_box.critical.assert_not_called()

                    files = list(Path(directory).glob("견적서_*.xlsx"))
                    self.assertEqual(len(files), 1)
                    workbook = load_workbook(files[0])
                    try:
                        sheet = workbook.active
                        self.assertEqual(sheet["R12"].value, 122499)
                        self.assertEqual(sheet["B16"].value, "주문 상품")
                        self.assertEqual(sheet["H16"].value, "옵션 A")
                        self.assertEqual(sheet["K16"].value, 3)
                        self.assertEqual(sheet["B17"].value, "배송비")
                        self.assertEqual(sheet["R17"].value + sheet["V17"].value, 2500)
                    finally:
                        workbook.close()

    def test_quote_manual_switch_requires_confirmation_and_clears_import(self):
        self.load_order("naver")
        with patch.object(QMessageBox, "question", return_value=QMessageBox.StandardButton.No):
            self.widget.direct_input_radio.click()
        self.assertTrue(self.widget.naver_order_radio.isChecked())
        self.assertIn("최종 합계 122,499원", self.widget.summary_label.text())
        with patch.object(QMessageBox, "question", return_value=QMessageBox.StandardButton.Yes):
            self.widget.direct_input_radio.click()
        self.assertFalse(self.widget._is_order_import())
        self.assertEqual(self.widget._collect_items(), [])
        self.assertFalse(self.widget.discount_controls.isHidden())
        self.assertFalse(self.widget.free_shipping_check.isHidden())


class DocumentActionLayoutTests(unittest.TestCase):
    """문서 만들기 하단 버튼과 필수 칸 표시(알림·상태바 계획 D-3)."""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.widget = QuotationStatementWidget(on_created=Mock())

    def tearDown(self):
        self.widget.close()
        self.widget.deleteLater()

    def test_reset_is_apart_and_pdf_is_primary(self):
        from PySide6.QtWidgets import QHBoxLayout
        row = next(layout for layout in self.widget.findChildren(QHBoxLayout)
                   if layout.indexOf(self.widget.pdf_button) >= 0)
        order = [row.itemAt(i).widget() or "stretch" for i in range(row.count())]
        self.assertEqual(order, [self.widget.reset_button, "stretch", self.widget.create_button, self.widget.pdf_button])
        self.assertEqual(self.widget.pdf_button.property("role"), "primary")
        self.assertFalse(self.widget.create_button.property("role"))

    def test_compact_layout_without_fixed_widths(self):
        # 디자인 개편 4-4: 작성 방식은 한 줄 전체, 입력칸은 고정 폭 없이 남는 폭을 쓰고, 버튼 줄은 위젯 하나.
        from PySide6.QtWidgets import QGridLayout
        grid = next(layout for layout in self.widget.findChildren(QGridLayout)
                    if layout.indexOf(self.widget.input_mode_widget) >= 0)
        row, column, row_span, column_span = grid.getItemPosition(grid.indexOf(self.widget.input_mode_widget))
        self.assertEqual((column, column_span), (1, 3))
        for edit in (self.widget.search_edit, self.widget.naver_order_id_edit, self.widget.coupang_order_id_edit):
            with self.subTest(edit=edit.placeholderText()[:10]):
                self.assertNotEqual(edit.minimumWidth(), edit.maximumWidth())
        self.assertIs(self.widget.pdf_button.parentWidget(), self.widget.action_bar)
        self.assertFalse(self.widget.product_search_widget.isAncestorOf(self.widget.input_mode_widget))

    def test_name_label_marks_required(self):
        label = self.widget.name_label
        self.assertEqual(label.text(), "성명 *")
        self.assertEqual(label.property("status"), "warn")
        self.widget.name_edit.setText("홍길동")
        self.assertFalse(label.property("status"))
        self.widget.name_edit.clear()  # 비우면 다시 주의 색(초기화는 확인 창이 떠서 직접 비운다)
        self.assertEqual(label.property("status"), "warn")


if __name__ == "__main__":
    unittest.main()
