"""알림 띠(shared/ui_notice.py)와 파일 생성 완료 알림(계획 C-1)을 검증한다."""

import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class NoticeBarTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def bar(self):
        from shared.ui_notice import NoticeBar
        bar = NoticeBar()
        bar.resize(600, 40)
        return bar

    def test_hidden_until_notice(self):
        bar = self.bar()
        self.assertTrue(bar.isHidden())

    def test_ok_notice_with_actions_and_auto_hide(self):
        from shared.ui_notice import OK_AUTO_HIDE_MS
        bar = self.bar()
        calls = []
        bar.show_notice("ok", "송장 엑셀 파일이 생성되었습니다: a.xlsx",
                        [("엑셀 열기", lambda: calls.append("open")), ("폴더 열기", lambda: calls.append("folder"))],
                        detail="C:/output/a.xlsx")
        self.assertFalse(bar.isHidden())
        self.assertEqual(bar.property("notice"), "ok")
        self.assertEqual(bar.findChild(type(bar._mark), "label_notice_mark").text(), "✓")
        self.assertEqual(bar.action_labels(), ["엑셀 열기", "폴더 열기"])
        bar._buttons[1].click()
        bar._buttons[0].click()
        self.assertEqual(calls, ["folder", "open"])
        self.assertEqual(bar._text.toolTip(), "C:/output/a.xlsx")
        self.assertTrue(bar._timer.isActive())
        self.assertEqual(bar._timer.interval(), OK_AUTO_HIDE_MS)
        bar._timer.timeout.emit()
        self.assertTrue(bar.isHidden())

    def test_hover_pauses_auto_hide(self):
        from PySide6.QtCore import QEvent, QPointF
        from PySide6.QtGui import QEnterEvent
        bar = self.bar()
        bar.show_notice("ok", "완료")
        bar.enterEvent(QEnterEvent(QPointF(1, 1), QPointF(1, 1), QPointF(1, 1)))
        self.assertFalse(bar._timer.isActive())
        bar.leaveEvent(QEvent(QEvent.Type.Leave))
        self.assertTrue(bar._timer.isActive())

    def test_warn_and_error_stay_until_closed(self):
        bar = self.bar()
        bar.show_notice("warn", "주의")
        self.assertFalse(bar._timer.isActive())
        bar.show_notice("error", "실패")
        self.assertFalse(bar._timer.isActive())
        self.assertEqual(bar.property("notice"), "error")
        bar._close.click()
        self.assertTrue(bar.isHidden())

    def test_new_notice_replaces_previous(self):
        bar = self.bar()
        bar.show_notice("ok", "첫째", [("열기", lambda: None), ("폴더 열기", lambda: None)])
        bar.show_notice("warn", "둘째")
        self.assertEqual(bar.text(), "둘째")
        self.assertEqual(bar.action_labels(), [])
        self.assertTrue(all(button.isHidden() for button in bar._buttons))
        with self.assertRaises(ValueError):
            bar.show_notice("info", "x")

    def test_long_text_is_elided_but_kept(self):
        bar = self.bar()
        bar.show()
        long_name = "일괄발송_쿠팡_" + "가" * 200 + ".xlsx"
        bar.show_notice("ok", long_name)
        self.app.processEvents()
        bar._elide()
        self.assertEqual(bar.text(), long_name)
        self.assertIn("…", bar._text.text())
        self.assertEqual(bar._text.toolTip(), long_name)


class FileCreatedNoticeTests(unittest.TestCase):
    """파일 생성 뒤 확인 창 없이 알림 띠에 「엑셀 열기」·「폴더 열기」를 보인다."""

    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from tests.test_order_load_sheets import load_app_module
        cls.module = load_app_module()

    def window(self):
        from PySide6.QtWidgets import QMainWindow
        from tests.test_order_load_sheets import main_window_members
        from shared.ui_notice import NoticeBar
        names = ("setup_status_bar", "notify", "show_excel_created_message",
                 "_open_created_file", "_open_created_file_location")
        members = main_window_members(self.module)
        window = type("FakeWindow", (QMainWindow,), {n: members[n] for n in names})()
        window.setup_status_bar()
        window.notice_bar = NoticeBar(window)
        window.opened = []
        window.open_file_with_default_app = lambda path: window.opened.append(("file", path)) or True
        window.open_file_location = lambda path: window.opened.append(("folder", path)) or False
        return window

    def test_created_file_uses_notice_not_dialog(self):
        from PySide6.QtWidgets import QMessageBox
        window = self.window()
        fail = lambda *a, **k: self.fail("확인 창을 띄우면 안 된다")  # noqa: E731
        with patch.object(QMessageBox, "exec", fail), patch.object(QMessageBox, "information", fail):
            window.show_excel_created_message("C:/output/일괄발송_쿠팡_1.xlsx", "일괄 발송 파일이 생성되었습니다.")
        bar = window.notice_bar
        self.assertEqual(bar.text(), "일괄 발송 파일이 생성되었습니다: 일괄발송_쿠팡_1.xlsx")
        self.assertEqual(bar.action_labels(), ["엑셀 열기", "폴더 열기"])
        self.assertIn("일괄 발송 파일이 생성되었습니다", window.status_label.text())
        bar._buttons[0].click()
        self.assertEqual(window.opened, [("file", "C:/output/일괄발송_쿠팡_1.xlsx")])
        with patch.object(QMessageBox, "warning") as warning:
            bar._buttons[1].click()  # 폴더 열기 실패는 경고 창으로 알린다(그대로)
        warning.assert_called_once()

    def test_last_notice_stays_after_banner_and_temporary_status_expire(self):
        window = self.window()
        window.notify("ok", "파일 생성 완료")
        window.notice_bar._timer.timeout.emit()
        self.assertTrue(window.notice_bar.isHidden())
        window._status_clear_timer.timeout.emit()
        self.assertEqual(window.status_label.text(), "✓ 파일 생성 완료")

        window.statusBar().showMessage("메뉴 안내")
        self.assertEqual(window.status_label.text(), "메뉴 안내")
        window._status_clear_timer.timeout.emit()
        self.assertEqual(window.status_label.text(), "✓ 파일 생성 완료")

    def test_pdf_label(self):
        window = self.window()
        window.show_excel_created_message("C:/output/견적서.pdf", "견적서 엑셀 및 PDF 파일이 생성되었습니다.", "PDF 열기")
        self.assertEqual(window.notice_bar.action_labels(), ["PDF 열기", "폴더 열기"])

    def test_no_completion_dialogs_left_in_source(self):
        # 파일 생성 「완료」 확인 창 네 곳(우체국 접수 엑셀, 일괄 발송, 분류 결과, 퀵 엑셀)을 모두 바꿨다.
        # 메인 창 코드는 easy-fulfill.py와 ui/main/의 믹스인 모듈에 나뉘어 있다.
        for path in [ROOT / "easy-fulfill.py", *sorted((ROOT / "ui" / "main").glob("*.py"))]:
            with self.subTest(path=path.name):
                self.assertNotIn('setWindowTitle("완료")', path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
