"""OZ Viewer 창 제목 필터의 순수 로직 검증."""

import unittest

from epost_oz_print_dialog import (
    close_context_after_print,
    PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS,
    PRINT_DIALOG_OPEN_ATTEMPTS,
    OZ_TOOLBAR_DIAGNOSTIC_PATH,
)
from epost_desktop_windows import (
    OZ_PRINT_TOOLBAR_CLIENT_X,
    OZ_PRINT_TOOLBAR_CLIENT_Y,
    PRINT_DIALOG_TITLE,
    DesktopWindow,
    IDOK,
    matching_visible_windows,
    MFC_ID_FILE_PRINT,
    OZ_PRINT_TOOLBAR_BUTTON_INDEX,
    wait_for_oz_toolbar_diagnostics,
    click_oz_viewer_print_toolbar_icon,
    enable_per_monitor_dpi_awareness,
    send_enter_to_foreground_window,
    oz_print_command_id,
    oz_toolbar_handle_from_diagnostics,
)


class EpostDesktopWindowsTests(unittest.TestCase):
    def test_closed_browser_context_does_not_override_completed_print_result(self):
        class AlreadyClosedContext:
            def close(self):
                raise RuntimeError("Target page, context or browser has been closed")

        close_context_after_print(AlreadyClosedContext())

    def test_only_oz_viewer_title_is_selected(self):
        windows = [
            DesktopWindow(handle=1, title="Easy Fulfill - 주문 처리 시스템"),
            DesktopWindow(handle=2, title="오즈 리포트 뷰어"),
            DesktopWindow(handle=3, title="인쇄"),
        ]
        self.assertEqual(
            matching_visible_windows(windows, "오즈 리포트 뷰어"),
            [DesktopWindow(handle=2, title="오즈 리포트 뷰어")],
        )

    def test_print_dialog_title_is_exact(self):
        self.assertEqual(PRINT_DIALOG_TITLE, "인쇄")
        self.assertEqual(IDOK, 1)


    def test_oz_toolbar_and_printer_point_are_identified(self):
        self.assertEqual((OZ_PRINT_TOOLBAR_CLIENT_X, OZ_PRINT_TOOLBAR_CLIENT_Y), (49, 20))
        self.assertEqual(oz_toolbar_handle_from_diagnostics([{
            "handle": 123, "className": "Afx:ToolBar:1", "top": 53,
        }]), 123)
        self.assertEqual(oz_toolbar_handle_from_diagnostics([{
            "handle": 123, "className": "Button", "top": 53,
        }]), 0)
        self.assertEqual(oz_print_command_id([
            {"index": 0, "commandId": 0xE103},
            {"index": 1, "commandId": MFC_ID_FILE_PRINT},
        ]), MFC_ID_FILE_PRINT)
        self.assertEqual(OZ_PRINT_TOOLBAR_BUTTON_INDEX, 1)
        self.assertEqual(oz_print_command_id([
            {"index": 0, "commandId": 32796},
            {"index": 1, "commandId": 32832},
            {"index": 2, "commandId": 0},
        ]), 32832)
        self.assertEqual(oz_print_command_id([
            {"index": 0, "commandId": 32796},
            {"index": 1, "commandId": 0},
        ]), 0)
        self.assertEqual(oz_print_command_id([]), 0)

    def test_print_dialog_retry_is_bounded(self):
        self.assertEqual(PRINT_DIALOG_OPEN_ATTEMPTS, 1)
        self.assertGreaterEqual(PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS, 8)
        self.assertLessEqual(PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS, 30)

    def test_toolbar_diagnostic_stays_in_local_output(self):
        self.assertEqual(OZ_TOOLBAR_DIAGNOSTIC_PATH.parent.name, "output")

    def test_toolbar_wait_is_available(self):
        self.assertEqual(wait_for_oz_toolbar_diagnostics.__name__, "wait_for_oz_toolbar_diagnostics")
        self.assertEqual(click_oz_viewer_print_toolbar_icon.__name__, "click_oz_viewer_print_toolbar_icon")
        self.assertEqual(enable_per_monitor_dpi_awareness.__name__, "enable_per_monitor_dpi_awareness")
        self.assertEqual(send_enter_to_foreground_window.__name__, "send_enter_to_foreground_window")


if __name__ == "__main__":
    unittest.main()
