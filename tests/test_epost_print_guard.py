"""실제 인쇄(Enter 전송)가 인쇄 창이 전면임을 확인한 뒤에만 실행되는지 검증한다."""

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from shipping import epost_desktop_windows as desktop
from shipping import epost_oz_print_dialog as dialog

VIEWER = 100


def foreground(handle, title, pid, image="ozviewer.exe"):
    """전면 창 정보를 흉내 낸다(뷰어 PID는 10)."""
    pids = {VIEWER: 10, handle: pid}
    images = {10: "ozviewer.exe", pid: image}
    return patch.multiple(
        desktop,
        _foreground_window=lambda: handle,
        _window_title=lambda h: title,
        _window_process_id=lambda h: pids.get(h, 0),
        _process_image_name=lambda p: images.get(p, ""),
    )


class ForegroundCheckTests(unittest.TestCase):
    def test_accepts_print_dialog_of_viewer_or_driver_host(self):
        with foreground(200, "인쇄", 10):
            self.assertTrue(desktop.is_print_dialog_foreground(VIEWER))
        with foreground(201, "Printing Preferences", 10):
            self.assertTrue(desktop.is_print_dialog_foreground(VIEWER))  # 뷰어 프로세스의 모달 창
        with foreground(202, "프린터 설정", 77, image="splwow64.exe"):
            self.assertTrue(desktop.is_print_dialog_foreground(VIEWER))  # 32비트 앱의 드라이버 창

    def test_rejects_viewer_itself_and_unrelated_windows(self):
        with foreground(VIEWER, "오즈 리포트 뷰어", 10):
            self.assertFalse(desktop.is_print_dialog_foreground(VIEWER))
        with foreground(300, "메모장", 55, image="notepad.exe"):
            self.assertFalse(desktop.is_print_dialog_foreground(VIEWER))
        with foreground(0, "", 0):
            self.assertFalse(desktop.is_print_dialog_foreground(VIEWER))


class ExecutePrintTests(unittest.TestCase):
    def run_execute(self, ready):
        store = Mock()
        candidates = [SimpleNamespace(regi_no="6012345678901")]
        viewer = SimpleNamespace(handle=VIEWER)
        with patch.object(dialog, "wait_for_print_dialog_foreground", return_value=ready) as wait, \
                patch.object(dialog, "send_enter_to_print_dialog", return_value=ready) as enter:
            outcome = None
            try:
                dialog.execute_print_when_dialog_ready(viewer, candidates, store)
                outcome = "sent"
            except RuntimeError as error:
                outcome = str(error)
        return outcome, store, wait, enter

    def test_sends_enter_and_records_only_when_dialog_is_in_front(self):
        outcome, store, wait, enter = self.run_execute(True)
        self.assertEqual(outcome, "sent")
        enter.assert_called_once_with(VIEWER)
        store.mark_windows_print_requested.assert_called_once_with(["6012345678901"])

    def test_no_enter_and_no_record_when_dialog_not_confirmed(self):
        outcome, store, wait, enter = self.run_execute(False)
        self.assertIn("보내지 않았습니다", outcome)
        self.assertIn("이력도 남기지 않았습니다", outcome)
        enter.assert_not_called()
        store.mark_windows_print_requested.assert_not_called()

    def test_send_rechecks_foreground_right_before_key(self):
        with patch.object(desktop, "is_print_dialog_foreground", return_value=False), \
                patch.object(desktop, "_press_enter") as press:
            self.assertFalse(desktop.send_enter_to_print_dialog(VIEWER))
        press.assert_not_called()
        with patch.object(desktop, "is_print_dialog_foreground", return_value=True), \
                patch.object(desktop, "_press_enter") as press:
            self.assertTrue(desktop.send_enter_to_print_dialog(VIEWER))
        press.assert_called_once()


if __name__ == "__main__":
    unittest.main()
