"""포털 팝업 인쇄 버튼의 안전한 후보 선택 검증."""

import unittest

from epost_portal_oz_viewer import (
    POPUP_PRINT_BUTTON_DIAGNOSTIC_PATH,
    preferred_popup_print_button_id,
)
from epost_portal_print_popup import preferred_outer_control_id


class EpostPortalOzViewerTests(unittest.TestCase):
    def test_popup_print_diagnostic_stays_in_local_output(self):
        self.assertEqual(POPUP_PRINT_BUTTON_DIAGNOSTIC_PATH.parent.name, "output")

    def test_outer_popup_print_component_is_preferred(self):
        self.assertEqual(
            preferred_outer_control_id(["popup_btnPrint", "popup_btnPrintTextBoxElement"]),
            "popup_btnPrint",
        )

    def test_only_popup_header_print_button_is_accepted(self):
        self.assertEqual(
            preferred_popup_print_button_id([
                "mainframe_work_btnDataPrint",
                "mainframe_popup_sbfabc_form_btnPrint2",
                "mainframe_popup_sbfabc_form_Div03_Div04_btnPrint",
            ]),
            "mainframe_popup_sbfabc_form_btnPrint2",
        )
        self.assertEqual(
            preferred_popup_print_button_id([
                "mainframe_popup_sbfabc_form_btnPrint2",
                "mainframe_popup_sbfdef_form_btnPrint2",
            ]),
            "",
        )


if __name__ == "__main__":
    unittest.main()
