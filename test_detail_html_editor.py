"""상세 HTML 편집창만 로드해 외부 계정 연결 없이 Qt 편집·미리보기를 검증한다."""

import ast
import html
import json
import math
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
    QPlainTextEdit, QPushButton, QSpinBox, QSplitter, QTextEdit, QVBoxLayout, QWidget, QDoubleSpinBox,
)
from PySide6.QtWebEngineCore import QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView
from shiboken6 import isValid
from detail_preview_server import DetailPreviewServer


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
        self.dialog.preview.page().runJavaScript(f"JSON.stringify(({expression}) ?? null)", results.append)
        deadline = time.monotonic() + 10
        while not results and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.01)
        self.assertTrue(results, "JavaScript callback timed out")
        return json.loads(results[0])

    def wait_preview(self, spacer_count):
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            if self.dialog._preview_ready and self.javascript(
                "Boolean(window.__easyFulfillPaintBlocks) && "
                f"document.querySelectorAll('[data-ef-spacer][data-ef-block]').length === {spacer_count}"
            ):
                return
            self.app.processEvents()
            time.sleep(.02)
        self.fail("Preview selector did not initialize")

    def align(self, direction, spacers=0):
        self.wait_preview(spacers)
        self.dialog._update_alignment_controls()
        self.dialog.alignment_buttons[direction].click()
        deadline = time.monotonic() + 10
        while self.dialog._alignment_pending and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.01)
        self.assertFalse(self.dialog._alignment_pending)
        self.wait_preview(spacers)

    def wait_checked_alignment(self, direction):
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            self.app.processEvents()
            if self.dialog.alignment_buttons[direction].isChecked():
                return
            time.sleep(.01)
        self.fail(f"Alignment state did not become {direction}")

    def wait_font_state(self):
        self.wait_preview(0)
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            self.app.processEvents()
            if self.dialog._font_state is not None and not self.dialog._alignment_pending:
                return
            time.sleep(.01)
        self.fail("Font state did not initialize")

    def finish_format_edit(self):
        deadline = time.monotonic() + 10
        while self.dialog._alignment_pending and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.01)
        self.assertFalse(self.dialog._alignment_pending)
        self.wait_font_state()

    def test_bold_toggle_normalizes_nested_bold_and_preserves_links(self):
        self.dialog.show()
        source = '<div><p style="color:red">보통 😀 <strong>굵게 <span style="font-weight:900">매우 굵게</span></strong> <a href="https://example.com">링크</a></p><p>다른 문단</p></div>'
        self.dialog.editor.setPlainText(source)
        self.select(0)
        self.wait_font_state()
        self.assertFalse(self.dialog.bold_button.isChecked())
        self.assertIn('굵기 혼합', self.dialog.font_status.text())
        self.dialog.bold_button.click()
        self.finish_format_edit()
        self.assertTrue(self.dialog.bold_button.isChecked())
        self.assertEqual(self.javascript("[...document.querySelector('p').querySelectorAll('strong,span,a')].map(e=>getComputedStyle(e).fontWeight)"), ['700', '700', '700'])
        bold = self.dialog.editor.toPlainText()
        self.dialog.bold_button.click()
        self.finish_format_edit()
        self.assertFalse(self.dialog.bold_button.isChecked())
        self.assertEqual(self.javascript("[document.querySelector('p'),...document.querySelector('p').querySelectorAll('strong,span,a')].map(e=>getComputedStyle(e).fontWeight)"), ['400'] * 4)
        self.assertIn('href="https://example.com"', self.dialog.editor.toPlainText())
        self.assertIn('<p>다른 문단</p>', self.dialog.editor.toPlainText())
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('p')).color"), 'rgb(255, 0, 0)')
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), bold)
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_font_size_applies_nested_spans_preserving_other_styles(self):
        self.dialog.show()
        source = '<div><h2 style="text-align:right;font-size:24px">제목</h2><p style="line-height:1.75;color:blue">본문 <strong><span style="font-size:12px">작은 글자</span></strong></p><img style="width:50px;height:20px" src="data:,x"></div>'
        self.dialog.editor.setPlainText(source)
        self.select(0, 1, 2)
        self.wait_font_state()
        self.assertEqual(self.dialog.font_size.value(), 0)
        self.assertFalse(self.dialog.apply_font_size.isEnabled())
        self.assertIn('크기 혼합', self.dialog.font_status.text())
        self.dialog.font_size.setValue(27.5)
        self.dialog.apply_font_size.click()
        self.finish_format_edit()
        for mobile in (False, True):
            self.dialog._set_preview_mode(mobile)
            self.app.processEvents()
            self.assertEqual(self.javascript("[...document.querySelectorAll('h2,p,strong,span')].map(e=>getComputedStyle(e).fontSize)"), ['27.5px'] * 4)
        self.assertEqual(self.dialog.font_size.value(), 27.5)
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('h2')).textAlign"), 'right')
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('strong')).fontWeight"), '700')
        self.assertIn('<img style="width:50px;height:20px" src="data:,x">', self.dialog.editor.toPlainText())
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_text_format_disabled_for_images_and_spacing(self):
        self.dialog.show()
        self.wait_preview(0)
        self.select(3)
        self.app.processEvents()
        self.assertFalse(self.dialog.bold_button.isEnabled())
        self.assertFalse(self.dialog.font_size.isEnabled())
        self.dialog.add_spacing.click()
        self.wait_preview(1)
        self.assertFalse(self.dialog.bold_button.isEnabled())
        self.assertFalse(self.dialog.apply_font_size.isEnabled())

    def test_table_text_format_and_save_reload(self):
        self.dialog.show()
        source = '<div><table style="width:120px"><tr><th>제목</th><td><a href="https://example.com">값 &amp; 정보</a></td></tr></table></div>'
        self.dialog.editor.setPlainText(source)
        self.select(0)
        self.wait_font_state()
        self.dialog.bold_button.click()
        self.finish_format_edit()
        self.dialog.font_size.setValue(21)
        self.dialog.apply_font_size.click()
        self.finish_format_edit()
        self.assertEqual(self.javascript("[...document.querySelectorAll('th,td,a')].map(e=>[getComputedStyle(e).fontWeight,getComputedStyle(e).fontSize])"), [['700', '21px']] * 3)
        self.assertIn('width: 120px', self.dialog.editor.toPlainText())
        with patch.object(QMessageBox, 'information'):
            self.dialog._save()
        saved = self.path.read_text(encoding='utf-8')
        copy_button = next(button for button in self.dialog.findChildren(QPushButton) if button.text() == 'HTML 복사')
        copy_button.click()
        self.assertEqual(QApplication.clipboard().text(), saved)
        self.assertNotIn('data-ef-block', saved)
        self.assertNotIn('outline', saved)
        self.dialog.editor.setPlainText(saved)
        self.select(0)
        self.wait_font_state()
        self.assertTrue(self.dialog.bold_button.isChecked())
        self.assertEqual(self.dialog.font_size.value(), 21)

    def test_heading_bold_off_then_alignment_and_size_keep_history(self):
        self.dialog.show()
        source = '<div><h1>큰 제목</h1><p>본문</p></div>'
        self.dialog.editor.setPlainText(source)
        self.select(0)
        self.wait_font_state()
        self.assertTrue(self.dialog.bold_button.isChecked())
        self.dialog.bold_button.click()
        self.finish_format_edit()
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('h1')).fontWeight"), '400')
        self.align('right')
        self.wait_font_state()
        self.dialog.font_size.setValue(30)
        self.dialog.apply_font_size.click()
        self.finish_format_edit()
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('h1')).textAlign"), 'right')
        self.assertFalse(self.dialog.bold_button.isChecked())
        for _ in range(3):
            self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_text_alignment_preserves_quote_content_and_undo(self):
        self.dialog.show()
        self.select(0, 1, 2)
        self.align('center')
        self.wait_checked_alignment('center')
        result = self.dialog.editor.toPlainText()
        self.assertEqual(result.count('text-align: center'), 3)
        self.assertIn('<blockquote><div>', result)
        self.assertIn('제목 😀', result)
        self.assertIn('1번 핀이 3번 핀', result)
        self.assertEqual(self.dialog._selected_blocks, {0, 1, 2})
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('blockquote p')).textAlign"), 'center')
        self.assertNotIn('outline', result)
        self.assertNotIn('cursor', result)
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), SOURCE)
        self.assertEqual(self.dialog._selected_blocks, {0, 1, 2})

    def test_image_placement_preserves_size_and_vertical_margins(self):
        self.dialog.show()
        image_source = SOURCE.replace('<img ', '<img style="width:60px;height:30px;margin:7px auto 19px" ', 1)
        self.dialog.editor.setPlainText(image_source)
        self.select(3)
        for mobile in (False, True):
            self.dialog._set_preview_mode(mobile)
            for direction in ('left', 'center', 'right'):
                self.align(direction)
                self.wait_checked_alignment(direction)
                geometry = self.javascript("""(() => {
                    const image = document.querySelector('img');
                    const r = image.getBoundingClientRect(), p = image.parentElement.getBoundingClientRect();
                    const s = getComputedStyle(image);
                    return [r.left - p.left, p.right - r.right, r.width, r.height, s.marginTop, s.marginBottom];
                })()""")
                self.assertEqual(geometry[2:], [60, 30, '7px', '19px'])
                if direction == 'left':
                    self.assertAlmostEqual(geometry[0], 0, delta=1)
                elif direction == 'right':
                    self.assertAlmostEqual(geometry[1], 0, delta=1)
                else:
                    self.assertAlmostEqual(geometry[0], geometry[1], delta=1)
                self.assertEqual(self.dialog._block_details[4]['attrs'].get('style'), None)

    def test_alignment_mixed_selection_and_spacer_exclusion(self):
        self.dialog.show()
        self.select(0)
        self.dialog.add_spacing.click()
        self.wait_preview(1)
        self.assertTrue(all(not button.isEnabled() for button in self.dialog.alignment_buttons.values()))
        self.select(0, 1, 4)  # 여백, 제목, 첫 이미지의 비연속 선택
        spacer_before = self.dialog.editor.toPlainText()[slice(*self.dialog._block_details[0]['range'])]
        self.align('right', spacers=1)
        self.assertEqual(self.dialog.editor.toPlainText()[slice(*self.dialog._block_details[0]['range'])], spacer_before)
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('p')).textAlign"), 'right')
        self.select(1, 2)  # 오른쪽 제목과 기본 왼쪽 인용구
        self.javascript('true')
        self.app.processEvents()
        self.assertTrue(all(not button.isChecked() for button in self.dialog.alignment_buttons.values()))

    def test_style_attribute_replacement_keeps_other_attributes(self):
        tag = '''<img data-note="style='wrong' >" src='a?x=1&amp;y=2' STYLE='margin:0 auto' />'''
        result = self.dialog._tag_with_style(tag, 'margin-left: 0px; background: url("a;b");')
        self.assertIn('data-note="style=\'wrong\' >"', result)
        self.assertIn("src='a?x=1&amp;y=2'", result)
        parser = HTMLParser()
        attributes = []
        parser.handle_startendtag = lambda _tag, attrs: attributes.extend(attrs)
        parser.feed(result)
        self.assertEqual([value for name, value in attributes if name == 'style'], ['margin-left: 0px; background: url("a;b");'])

    def test_inherited_alignment_and_saved_html(self):
        self.dialog.show()
        source = SOURCE.replace('max-width:780px;', 'text-align:right;max-width:780px;')
        self.dialog.editor.setPlainText(source)
        self.select(0)
        self.wait_preview(0)
        self.wait_checked_alignment('right')
        self.align('left')
        saved = self.dialog.editor.toPlainText()
        with patch.object(QMessageBox, 'information'):
            self.dialog._save()
        copy_button = next(button for button in self.dialog.findChildren(QPushButton) if button.text() == 'HTML 복사')
        copy_button.click()
        self.assertEqual(self.path.read_text(encoding='utf-8'), saved)
        self.assertEqual(QApplication.clipboard().text(), saved)
        self.assertIn('text-align: left', saved)
        self.assertNotIn('data-ef-block', saved)
        # 저장 HTML을 다시 불러와도 정렬 상태를 읽을 수 있다.
        self.dialog.editor.setPlainText(saved)
        self.select(0)
        self.wait_preview(0)
        self.wait_checked_alignment('left')

    def test_css_values_with_semicolons_and_small_table_alignment(self):
        self.dialog.show()
        source = '<div><p style="color:red;background-image:url(&quot;https://example.com/a;b.png&quot;)">본문</p><table style="width:120px;margin:3px auto 9px"><tr><td>셀</td></tr></table></div>'
        # CSSOM 선언 보존만 확인하며 외부 이미지는 요청하지 않는다.
        source = source.replace('https://example.com/a;b.png', 'data:image/svg+xml;base64,PHN2Zy8+')
        self.dialog.editor.setPlainText(source)
        self.select(0, 1)
        self.align('right')
        self.assertIn('data:image/svg+xml;base64,PHN2Zy8+', self.dialog.editor.toPlainText())
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('p')).color"), 'rgb(255, 0, 0)')
        self.assertEqual(self.javascript("document.querySelector('table').getBoundingClientRect().width"), 120)
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('table')).marginBottom"), '9px')

    def test_spacing_boundaries_match_clicked_items(self):
        details = self.dialog._html_block_details(SOURCE)
        self.assertEqual(len(details), 6)
        self.assertEqual(details[1]["boundary"], details[1]["range"])
        self.assertNotEqual(details[1]["boundary"], details[2]["boundary"])
        self.assertNotEqual(details[3]["boundary"], details[4]["boundary"])
        self.assertEqual(SOURCE[slice(*details[3]["boundary"])].count("<img "), 1)

    def test_add_above_quote_then_undo_preserves_unicode_and_selection(self):
        self.select(2)
        self.dialog.add_spacing.click()
        self.assertIn('</p>' + self.dialog._spacer_html(24) + '<p>1번 핀이', self.dialog.editor.toPlainText())
        self.assertEqual(self.dialog._selected_spacer()["spacer"], "24")
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), SOURCE)
        self.assertEqual(self.dialog._selected_blocks, {2})

    def test_range_inserts_once_below_last_selected_image(self):
        self.select(1, 2, 3)
        self.dialog.spacing_position.setCurrentIndex(1)
        self.dialog.spacing_height.setValue(48)
        self.dialog.add_spacing.click()
        result = self.dialog.editor.toPlainText()
        self.assertEqual(result.count('data-ef-spacer='), 1)
        first_image = SOURCE[slice(*DetailHtmlEditorDialog._html_block_ranges(SOURCE)[3])]
        self.assertIn(first_image + self.dialog._spacer_html(48) + '<img ', result)

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
        self.assertEqual(source[slice(*details[0]["boundary"])], '<li>가</li>')
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

    def item_html(self):
        source = self.dialog.editor.toPlainText()
        return [source[start:end] for start, end in self.dialog._block_ranges]

    def test_preview_scroll_preserved_across_edit_operations(self):
        source = '<div style="width:1200px">' + ''.join(f'<p style="height:80px;margin:0">문단 {index}</p>' for index in range(75)) + '</div>'
        self.dialog.show()
        self.dialog.editor.setPlainText(source)
        for mobile in (False, True):
            self.dialog._set_preview_mode(mobile)
            self.wait_preview(0)
            self.javascript('window.scrollTo({left:120,top:1400,behavior:"instant"})')
            expected = self.javascript('[window.scrollX,window.scrollY]')
            self.assertEqual(expected, [120, 1400])

            def unchanged(spacers=0):
                self.wait_preview(spacers)
                actual = self.javascript('[window.scrollX,window.scrollY]')
                self.assertEqual(actual, expected)

            self.select(18)
            self.align('right')
            unchanged()
            self.wait_font_state()
            self.dialog.bold_button.click()
            self.finish_format_edit()
            unchanged()
            self.dialog.font_size.setValue(22)
            self.dialog.apply_font_size.click()
            self.finish_format_edit()
            unchanged()
            self.dialog.move_up.click()
            unchanged()
            self.dialog._undo_editor()
            unchanged()
            # 화면 밖 항목을 선택해도 추가 후 선택 항목으로 강제 이동하지 않는다.
            self.select(50)
            self.dialog.add_spacing.click()
            unchanged(1)
            self.dialog.spacing_height.setValue(48)
            self.dialog.resize_spacing.click()
            unchanged(1)
            self.dialog.delete_blocks.click()
            unchanged()

    def test_rapid_preview_refresh_keeps_scroll_and_latest_html(self):
        self.dialog.show()
        source = '<div>' + ''.join(f'<p style="height:90px">문단 {index}</p>' for index in range(60)) + '</div>'
        self.dialog.editor.setPlainText(source)
        self.wait_preview(0)
        self.javascript('window.scrollTo(0,1200)')
        self.dialog.editor.setPlainText(source.replace('문단 59', '첫 수정'))
        self.dialog.editor.setPlainText(source.replace('문단 59', '마지막 수정'))
        self.wait_preview(0)
        self.assertEqual(self.javascript('window.scrollY'), 1200)
        self.assertIn('마지막 수정', self.javascript('document.body.textContent'))
        self.assertNotIn('첫 수정', self.javascript('document.body.textContent'))
        self.select(30)
        for _ in range(3):
            self.dialog.add_spacing.click()
        self.wait_preview(3)
        self.assertEqual(self.javascript('window.scrollY'), 1200)

    def test_save_confirmation_closes_editor_after_writing(self):
        self.dialog.show()
        changed = SOURCE.replace('뒤 본문', '저장된 본문')
        self.dialog.editor.setPlainText(changed)

        def confirm(*_args):
            self.assertTrue(self.dialog.isVisible())
            self.assertEqual(self.path.read_text(encoding='utf-8'), changed)
            return QMessageBox.Ok

        with patch.object(QMessageBox, 'information', side_effect=confirm) as popup:
            save_action = next(action for action in self.dialog.actions() if action.text() == '저장')
            save_action.trigger()
        popup.assert_called_once()
        self.assertEqual(self.dialog.result(), QDialog.Accepted)
        self.assertFalse(self.dialog.isVisible())

    def test_save_failure_keeps_editor_open(self):
        self.dialog.show()
        with patch.object(Path, 'write_text', side_effect=OSError('write failed')), patch.object(QMessageBox, 'information') as popup:
            with self.assertRaises(OSError):
                self.dialog._save()
        popup.assert_not_called()
        self.assertTrue(self.dialog.isVisible())

    def test_youtube_iframes_survive_text_edit_copy_and_save(self):
        from naver_detail_preview import render_youtube_video

        videos = ''.join(render_youtube_video({
            'embedUrl': f'https://www.youtube.com/embed/{video_id}', 'title': '제품 시연',
        }) for video_id in ('5RydDnx4VE4', 'NaL3tsLzK0o'))
        source = '<div><p>상품 소개</p>' + videos + '<p>제품 사양</p></div>'
        # 원격 재생 성공 여부와 독립적으로 실제 편집/복사/저장 경로를 검증한다.
        with patch.object(self.dialog, '_queue_preview_refresh') as preview:
            self.dialog.editor.setPlainText(source)
            cursor = self.dialog.editor.document().find('상품 소개')
            cursor.insertText('수정된 상품 소개')
            changed = self.dialog.editor.toPlainText()
            self.assertIn(videos, changed)
            self.assertIn(videos, preview.call_args.args[0])
            self.assertIn('수정된 상품 소개', changed)
            copy_button = next(button for button in self.dialog.findChildren(QPushButton) if button.text() == 'HTML 복사')
            copy_button.click()
            self.assertEqual(QApplication.clipboard().text(), changed)
            with patch.object(QMessageBox, 'information', return_value=QMessageBox.Ok):
                self.dialog._save()
            self.assertEqual(self.path.read_text(encoding='utf-8'), changed)

    def test_http_preview_origin_relative_images_and_cleanup(self):
        import base64
        from urllib.error import URLError
        from urllib.request import urlopen

        self.dialog.show()
        images = self.path.parent / 'images'
        images.mkdir()
        (images / 'pixel.gif').write_bytes(base64.b64decode('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'))
        source = '<div><p>미리보기</p><img src="images/pixel.gif"></div>'
        self.dialog.editor.setPlainText(source)
        self.wait_preview(0)
        self.assertEqual(self.javascript('location.protocol'), 'http:')
        self.assertEqual(self.javascript('location.hostname'), '127.0.0.1')
        self.assertEqual(self.javascript('document.querySelector("img").naturalWidth'), 1)
        self.assertEqual(self.dialog.editor.toPlainText(), source)
        url = self.dialog._http_preview.url
        self.dialog.accept()
        self.assertTrue(self.dialog._http_preview.closed)
        with self.assertRaises(URLError):
            urlopen(url, timeout=1)

    def test_http_preview_sends_origin_as_cross_origin_referer(self):
        from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
        from threading import Thread
        from urllib.parse import urlsplit

        self.dialog.show()
        referers = []
        class Receiver(BaseHTTPRequestHandler):
            def do_GET(self):
                referers.append(self.headers.get('Referer'))
                self.send_response(204)
                self.end_headers()
            def log_message(self, *_args):
                pass

        receiver = ThreadingHTTPServer(('127.0.0.1', 0), Receiver)
        thread = Thread(target=receiver.serve_forever, kwargs={'poll_interval': .05}, daemon=True)
        thread.start()
        try:
            self.dialog.editor.setPlainText(
                f'<div><p>출처 확인</p><iframe src="http://127.0.0.1:{receiver.server_port}/probe"></iframe></div>'
            )
            self.wait_preview(0)
            parsed = urlsplit(self.dialog._http_preview.url)
            self.assertIn(f'{parsed.scheme}://{parsed.netloc}/', referers)
        finally:
            receiver.shutdown()
            receiver.server_close()
            thread.join(timeout=1)

    def test_shared_wrapper_spacing_regression(self):
        source = '<div><div style="font-size:18px"><h2>상품 소개</h2><p>A</p><p>B</p><p>C</p><p>제품 사양</p></div></div>'
        self.dialog.editor.setPlainText(source)
        self.select(3)
        self.dialog.add_spacing.click()
        spacer = self.dialog._spacer_html(24)
        self.assertEqual(self.dialog.editor.toPlainText(), source.replace('<p>C</p>', spacer + '<p>C</p>'))
        self.dialog._undo_editor()
        self.select(3)
        self.dialog.spacing_position.setCurrentIndex(1)
        self.dialog.add_spacing.click()
        self.assertEqual(self.dialog.editor.toPlainText(), source.replace('<p>C</p>', '<p>C</p>' + spacer))

    def test_move_single_item_up_down_and_undo(self):
        source = '<div><p>A 😀</p>\n<p style="color:red">B</p>\n<p>C</p></div>'
        self.dialog.editor.setPlainText(source)
        original = self.item_html()
        self.select(2)
        self.assertFalse(self.dialog.move_down.isEnabled())
        self.dialog.move_up.click()
        self.assertEqual(self.item_html(), [original[0], original[2], original[1]])
        self.assertEqual(self.dialog._selected_blocks, {1})
        self.dialog.move_down.click()
        self.assertEqual(self.dialog.editor.toPlainText(), source)
        self.assertEqual(self.dialog._selected_blocks, {2})
        self.dialog._undo_editor()
        self.assertEqual(self.item_html(), [original[0], original[2], original[1]])
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)
        self.select(0)
        self.assertFalse(self.dialog.move_up.isEnabled())

    def test_move_contiguous_selection_and_reject_discontiguous(self):
        source = '<div><p>A</p><p>B</p><p>C</p><p>D</p></div>'
        self.dialog.editor.setPlainText(source)
        original = self.item_html()
        self.select(0, 2)
        self.assertFalse(self.dialog.move_up.isEnabled())
        self.assertFalse(self.dialog.move_down.isEnabled())
        self.select(1, 2)
        self.dialog.move_up.click()
        self.assertEqual(self.item_html(), [original[1], original[2], original[0], original[3]])
        self.assertEqual(self.dialog._selected_blocks, {0, 1})
        self.assertEqual(self.dialog._selection_range, (0, 1))
        self.dialog.move_down.click()
        self.assertEqual(self.dialog.editor.toPlainText(), source)
        self.assertEqual(self.dialog._selected_blocks, {1, 2})

    def test_cross_wrapper_move_keeps_inherited_styles_and_other_items(self):
        self.dialog.show()
        source = '<div><div style="color:red;font-size:14px"><p>A</p><p>B</p></div><!--경계--><div style="color:blue;font-size:23px"><p>C</p><p>D</p></div></div>'
        self.dialog.editor.setPlainText(source)
        original = self.item_html()
        self.select(2)
        self.dialog.move_up.click()
        self.wait_preview(0)
        self.assertEqual(self.item_html(), [original[0], original[2], original[1], original[3]])
        self.assertEqual(self.javascript("[...document.querySelectorAll('p')].map(e=>[e.textContent,getComputedStyle(e).color,getComputedStyle(e).fontSize])"), [
            ['A', 'rgb(255, 0, 0)', '14px'], ['C', 'rgb(0, 0, 255)', '23px'],
            ['B', 'rgb(255, 0, 0)', '14px'], ['D', 'rgb(0, 0, 255)', '23px']])
        self.assertEqual(self.dialog.editor.toPlainText().count('<!--경계-->'), 1)
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_list_item_and_quote_cross_boundaries_without_moving_neighbors(self):
        self.dialog.show()
        source = '<div><p>A</p><ul><li>B</li><li>C</li></ul><blockquote style="border-left:5px solid red"><p>D</p><p>E</p></blockquote><p>F</p></div>'
        self.dialog.editor.setPlainText(source)
        original = self.item_html()
        self.select(3)
        self.dialog.move_up.click()
        self.wait_preview(0)
        self.assertEqual(self.item_html(), [original[0], original[1], original[3], original[2], original[4], original[5]])
        self.assertEqual(self.javascript("[...document.querySelectorAll('li')].map(e=>e.parentElement.tagName)"), ['UL', 'UL'])
        self.assertEqual(self.javascript("[...document.querySelectorAll('blockquote')].map(e=>[e.textContent,getComputedStyle(e).borderLeftWidth])"), [['D', '5px'], ['E', '5px']])
        self.assertEqual(self.javascript("document.querySelectorAll('[data-ef-block]').length"), 6)
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_spacing_between_list_items_then_move_resize_delete(self):
        self.dialog.show()
        source = '<div><ol start="4"><li>A</li><li>B</li><li>C</li></ol><p>D</p></div>'
        self.dialog.editor.setPlainText(source)
        self.select(1)
        self.dialog.add_spacing.click()
        self.wait_preview(1)
        self.assertIn('<li>A</li>' + self.dialog._spacer_html(24, 'li') + '<li>B</li>', self.dialog.editor.toPlainText())
        self.assertEqual(self.javascript("document.querySelector('[data-ef-spacer]').parentElement.tagName"), 'OL')
        self.assertEqual(self.javascript("getComputedStyle(document.querySelector('[data-ef-spacer]')).display"), 'block')
        self.dialog.move_down.click()
        self.wait_preview(1)
        self.assertEqual(self.dialog._selected_blocks, {2})
        self.dialog.spacing_height.setValue(37)
        self.dialog.resize_spacing.click()
        self.wait_preview(1)
        self.assertEqual(self.javascript("document.querySelector('[data-ef-spacer]').getBoundingClientRect().height"), 37)
        self.dialog.delete_blocks.click()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_multi_selection_moves_across_list_and_quote_boundaries(self):
        source = '<div><p>A</p><ul><li>B</li><li>C</li></ul><blockquote><p>D</p><p>E</p></blockquote><p>F</p></div>'
        self.dialog.editor.setPlainText(source)
        original = self.item_html()
        self.select(1, 2, 3)
        self.dialog.move_up.click()
        self.assertEqual(self.item_html(), original[1:4] + [original[0]] + original[4:])
        self.assertEqual(self.dialog._selected_blocks, {0, 1, 2})
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_ordered_list_split_keeps_original_numbers(self):
        self.dialog.show()
        source = '<div><ol start="4"><li>A</li><li value="9">B</li><li>C</li></ol><p>D</p></div>'
        self.dialog.editor.setPlainText(source)
        self.select(2)
        self.dialog.move_down.click()
        self.wait_preview(0)
        self.assertEqual(self.javascript("[...document.querySelectorAll('ol')].map(e=>[e.start,e.textContent])"), [[4, 'AB'], [10, 'C']])
        self.dialog._undo_editor()
        self.assertEqual(self.dialog.editor.toPlainText(), source)

    def test_move_image_spacer_and_save_reload(self):
        self.dialog.show()
        self.select(3)
        self.dialog.add_spacing.click()
        before = self.item_html()
        self.dialog.move_down.click()
        self.wait_preview(1)
        self.assertEqual(self.item_html(), before[:3] + [before[4], before[3]] + before[5:])
        self.assertEqual(self.javascript("document.querySelectorAll('[data-ef-block]').length"), len(before))
        with patch.object(QMessageBox, 'information'):
            self.dialog._save()
        saved = self.dialog.editor.toPlainText()
        self.assertEqual(self.path.read_text(encoding='utf-8'), saved)
        self.assertNotIn('data-ef-block', saved)
        self.dialog.editor.setPlainText(saved)
        self.select(4)
        self.wait_preview(1)
        self.assertIsNotNone(self.dialog._selected_spacer())

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
