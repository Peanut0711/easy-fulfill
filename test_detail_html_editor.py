"""상세 HTML 편집창만 로드해 외부 계정 연결 없이 Qt 편집·미리보기를 검증한다."""

import ast
import json
import os
import re
import tempfile
import time
import unittest
from html.parser import HTMLParser
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS", "--disable-gpu")

from PySide6.QtCore import Qt, QUrl, QEvent, QCoreApplication
from PySide6.QtGui import QAction, QColor, QFont, QFontDatabase, QKeySequence, QTextCursor
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFrame, QHBoxLayout, QLabel, QMessageBox,
    QPlainTextEdit, QPushButton, QSpinBox, QSplitter, QTextEdit, QVBoxLayout, QWidget,
)
from PySide6.QtWebEngineCore import QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView


# MainWindow의 주문·설정 초기화를 실행하지 않고 실제 편집창 클래스 전체를 사용한다.
source_path = Path(__file__).with_name("easy-fulfill.py")
tree = ast.parse(source_path.read_text(encoding="utf-8"))
editor_class = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "DetailHtmlEditorDialog")
exec(compile(ast.Module(body=[editor_class], type_ignores=[]), str(source_path), "exec"))


SOURCE = (
    '<div style="max-width:780px;margin:0 auto">'
    '<div><p>제목 😀</p></div>'
    '<blockquote><div><p>케이블 색상은 바뀔 수 있습니다.</p>'
    '<p>1번 핀이 3번 핀에 연결되는 케이블입니다.</p></div></blockquote>'
    '<div><img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7">'
    '<img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"></div>'
    '<div><p>뒤 본문</p></div></div>'
)


class DetailEditorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        font_path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts/malgun.ttf"
        if font_path.exists():
            QFontDatabase.addApplicationFont(str(font_path))
            cls.app.setFont(QFont("Malgun Gothic", 9))

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name) / "detail.html"
        self.path.write_text(SOURCE, encoding="utf-8")
        self.dialog = DetailHtmlEditorDialog(self.path)

    def tearDown(self):
        self.dialog.close()
        self.dialog.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        self.app.processEvents()
        self.directory.cleanup()

    def select(self, *indices):
        self.dialog._selected_blocks = set(indices)
        self.dialog._update_block_controls()

    def javascript(self, expression):
        results = []
        self.dialog.preview.page().runJavaScript(expression, results.append)
        deadline = time.monotonic() + 10
        while not results and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.01)
        self.assertTrue(results, "JavaScript callback timed out")
        return results[0]

    def wait_preview(self, spacer_count):
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            if self.javascript(
                "Boolean(window.__easyFulfillPaintBlocks) && "
                f"document.querySelectorAll('div[data-ef-spacer][data-ef-block]').length === {spacer_count}"
            ):
                return
            self.app.processEvents()
            time.sleep(.02)
        self.fail("Preview selector did not initialize")

    def test_quote_and_image_group_boundaries(self):
        details = self.dialog._html_block_details(SOURCE)
        self.assertEqual(len(details), 6)
        self.assertEqual(details[1]["boundary"], details[2]["boundary"])
        self.assertTrue(SOURCE[slice(*details[1]["boundary"])].startswith("<blockquote>"))
        self.assertEqual(details[3]["boundary"], details[4]["boundary"])
        self.assertEqual(SOURCE[slice(*details[3]["boundary"])].count("<img "), 2)

    def test_add_above_quote_then_undo_preserves_unicode_and_selection(self):
        self.select(2)
        self.dialog.add_spacing.click()
        self.assertIn('</div>' + self.dialog._spacer_html(24) + '<blockquote>', self.dialog.editor.toPlainText())
        self.assertEqual(self.dialog._selected_spacer()["spacer"], "24")
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), SOURCE)
        self.assertEqual(self.dialog._selected_blocks, {2})

    def test_range_inserts_once_below_image_group(self):
        self.select(1, 2, 3)
        self.dialog.spacing_position.setCurrentIndex(1)
        self.dialog.spacing_height.setValue(48)
        self.dialog.add_spacing.click()
        result = self.dialog.editor.toPlainText()
        self.assertEqual(result.count('data-ef-spacer='), 1)
        self.assertIn('</div>' + self.dialog._spacer_html(48) + '<div><p>뒤 본문', result)

    def test_noncontiguous_selection_disabled(self):
        self.select(0, 2)
        self.assertFalse(self.dialog.add_spacing.isEnabled())
        self.dialog._add_spacing()
        self.assertEqual(self.dialog.editor.toPlainText(), SOURCE)

    def test_resize_delete_and_undo_each_edit(self):
        self.select(0)
        self.dialog.add_spacing.click()
        inserted = self.dialog.editor.toPlainText()
        self.dialog.spacing_height.setValue(37)
        self.assertEqual(self.dialog.spacing_preset.currentText(), "직접 입력")
        self.dialog.resize_spacing.click()
        resized = self.dialog.editor.toPlainText()
        self.assertEqual(self.dialog._selected_spacer()["spacer"], "37")
        self.dialog.delete_blocks.click()
        self.assertEqual(self.dialog.editor.toPlainText(), SOURCE)
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), resized)
        self.assertEqual(self.dialog._selected_spacer()["spacer"], "37")
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), inserted)
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), SOURCE)

    def test_list_and_table_are_not_split(self):
        source = '<div><ul><li>가</li><li>나</li></ul><table><tr><td><p>셀</p></td></tr></table></div>'
        details = self.dialog._html_block_details(source)
        self.assertEqual(len(details), 3)
        self.assertEqual(source[slice(*details[0]["boundary"])], '<ul><li>가</li><li>나</li></ul>')
        self.assertTrue(source[slice(*details[2]["boundary"])].startswith('<table>'))

    def test_existing_deletion_and_empty_selection(self):
        self.assertFalse(self.dialog.add_spacing.isEnabled())
        self.assertFalse(self.dialog.resize_spacing.isEnabled())
        self.select(1)
        self.dialog.delete_blocks.click()
        self.assertNotIn('케이블 색상은 바뀔 수 있습니다.', self.dialog.editor.toPlainText())
        self.assertIn('1번 핀이 3번 핀', self.dialog.editor.toPlainText())
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), SOURCE)

    def test_preview_click_height_and_clean_save_copy(self):
        self.dialog.show()
        self.wait_preview(0)
        self.select(0)
        self.dialog.spacing_preset.setCurrentIndex(2)
        self.dialog.add_spacing.click()
        self.wait_preview(1)
        for mobile in (False, True):
            self.dialog._set_preview_mode(mobile)
            self.app.processEvents()
            self.assertEqual(self.javascript("document.querySelector('[data-ef-spacer]').getBoundingClientRect().height"), 48)
        self.dialog._clear_block_selection()
        self.javascript("document.querySelector('[data-ef-spacer]').click()")
        self.app.processEvents()
        self.assertIsNotNone(self.dialog._selected_spacer())
        self.assertTrue(self.dialog.resize_spacing.isEnabled())
        self.assertIn('여백 48px', self.javascript("document.querySelector('[data-ef-spacer]').dataset.efSpacerLabel"))
        with patch.object(QMessageBox, 'information'):
            self.dialog._save()
        saved = self.path.read_text(encoding="utf-8")
        copy_button = next(button for button in self.dialog.findChildren(QPushButton) if button.text() == 'HTML 복사')
        copy_button.click()
        self.assertEqual(QApplication.clipboard().text(), saved)
        self.assertNotIn('클릭하여 조절', saved)
        self.assertNotIn('spacer-label', saved)
        self.assertNotIn('dashed', saved)
        self.assertIn('height:48px', saved)
        if os.environ.get('DETAIL_EDITOR_SCREENSHOT'):
            self.dialog._set_preview_mode(False)
            self.app.processEvents()
            self.dialog.grab().save(os.environ['DETAIL_EDITOR_SCREENSHOT'])


if __name__ == '__main__':
    unittest.main()
