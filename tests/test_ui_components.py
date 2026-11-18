"""공통 부품(shared/ui_components.py)을 검증한다."""

import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class ComponentTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def test_section_card_header(self):
        from PySide6.QtWidgets import QLabel, QPushButton
        from shared.ui_components import SectionCard
        card = SectionCard("주문 불러오기", step=1)
        self.assertTrue(card.property("card"))
        self.assertEqual(card.title_label.text(), "주문 불러오기")
        self.assertEqual(card.step_badge.text(), "1")
        self.assertFalse(card.header_widget.isHidden())
        button = card.add_action(QPushButton("새로고침"))
        self.assertEqual(button.sizePolicy().horizontalPolicy().name, "Fixed")
        card.body.addWidget(QLabel("내용"))
        card.set_step_done(True)
        self.assertEqual((card.step_badge.text(), card.step_badge.property("stepState")), ("✓", "done"))
        card.set_step_done(False)
        self.assertEqual(card.step_badge.text(), "1")

    def test_section_card_without_header_hides_it(self):
        from shared.ui_components import SectionCard
        card = SectionCard()
        self.assertTrue(card.header_widget.isHidden())
        card.set_title("제목")
        self.assertFalse(card.header_widget.isHidden())

    def test_button_row_order_and_fixed_width(self):
        from PySide6.QtWidgets import QPushButton
        from shared.ui_components import ButtonRow
        reset, excel, pdf = QPushButton("초기화"), QPushButton("엑셀 만들기"), QPushButton("PDF 만들기")
        row = ButtonRow(left=[reset], right=[excel, pdf])
        self.assertEqual(row.buttons(), [reset, excel, pdf])
        # 왼쪽 버튼 다음에 늘어나는 빈 칸이 있어 주 동작이 오른쪽 끝에 붙는다.
        self.assertGreater(row.layout_.itemAt(1).spacerItem().expandingDirections().value, 0)
        for button in (reset, excel, pdf):
            self.assertEqual(button.sizePolicy().horizontalPolicy().name, "Fixed")

    def test_form_grid_required_and_error(self):
        from PySide6.QtWidgets import QLineEdit, QPushButton, QWidget
        from shared.ui_components import FormGrid
        host = QWidget()
        grid = FormGrid(host)
        name, zip_code = QLineEdit(), QLineEdit()
        grid.add_row("이름", name, required=True)
        grid.add_row("우편번호", zip_code, extra=QPushButton("주소 검색"))
        label = grid.label_for(name)
        self.assertEqual(label.text(), "이름 *")
        self.assertIs(label.buddy(), name)
        grid.set_missing(name, True)
        self.assertEqual(label.property("status"), "warn")
        grid.set_missing(name, False)
        self.assertEqual(label.property("status"), "")
        grid.set_error(zip_code, "5자리 숫자로 입력하세요")
        error = grid._rows[zip_code][1]
        self.assertFalse(error.isHidden())
        grid.set_error(zip_code, "")
        self.assertTrue(error.isHidden())

    def test_empty_state_action_is_primary(self):
        from PySide6.QtWidgets import QPushButton
        from shared.ui_components import EmptyState
        action = QPushButton("주문 불러오기")
        empty = EmptyState("주문 엑셀을 불러오면 여기에 보입니다.", action)
        self.assertEqual(action.property("role"), "primary")
        self.assertTrue(empty.label.property("emptyText"))

    def test_setup_table(self):
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QTableWidget, QTableWidgetItem
        from shared.ui_components import NUMERIC_ALIGNMENT, align_numeric, setup_table
        table = setup_table(QTableWidget(2, 3))
        self.assertFalse(table.showGrid())
        self.assertFalse(table.verticalHeader().isVisible())
        self.assertEqual(table.textElideMode(), Qt.TextElideMode.ElideRight)
        item = align_numeric(QTableWidgetItem("12,000"))
        self.assertEqual(Qt.AlignmentFlag(item.textAlignment()), NUMERIC_ALIGNMENT)

    def _answer_with_default(self):
        from PySide6.QtCore import QTimer
        from PySide6.QtWidgets import QApplication

        def press():
            box = QApplication.activeModalWidget()
            box.defaultButton().click()
        QTimer.singleShot(0, press)

    def test_confirm_action_destructive_defaults_to_cancel(self):
        from shared.ui_components import confirm_action
        self._answer_with_default()
        self.assertFalse(confirm_action(None, "확인", "실제로 접수할까요?", "접수하기", destructive=True))

    def test_confirm_action_safe_defaults_to_action(self):
        from shared.ui_components import confirm_action
        self._answer_with_default()
        self.assertTrue(confirm_action(None, "확인", "다시 만들까요?", "다시 만들기"))


if __name__ == "__main__":
    unittest.main()
