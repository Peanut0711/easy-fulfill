"""OZ Viewer 창 제목 필터의 순수 로직 검증."""

import unittest

from epost_oz_print_dialog import (
    PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS,
    PRINT_DIALOG_OPEN_ATTEMPTS,
    OZ_TOOLBAR_DIAGNOSTIC_PATH,
)
from epost_desktop_windows import (
    OZ_PRINT_TOOLBAR_CLIENT_X,
    OZ_PRINT_TOOLBAR_CLIENT_Y,
    PRINT_DIALOG_TITLE,
    DesktopWindow,
    matching_visible_windows,
    MFC_ID_FILE_PRINT,
    OZ_PRINT_TOOLBAR_BUTTON_INDEX,
    oz_print_command_id,
    oz_toolbar_handle_from_diagnostics,
)


class EpostDesktopWindowsTests(unittest.TestCase):
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
        self.assertEqual(PRINT_DIALOG_OPEN_ATTEMPTS, 4)
        self.assertLessEqual(PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS, 3)

    def test_toolbar_diagnostic_stays_in_local_output(self):
        self.assertEqual(OZ_TOOLBAR_DIAGNOSTIC_PATH.parent.name, "output")


if __name__ == "__main__":
    unittest.main()
