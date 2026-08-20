"""우체국 운송장출력 팝업 전 단계의 안전한 컨트롤 선택 검증."""

import unittest

from epost_portal_print_popup import preferred_outer_control_id, print_popup_is_open


class _VisibleText:
    def __init__(self, visible):
        self.visible = visible

    def is_visible(self):
        return self.visible


class _PrintControls:
    def __init__(self, visible):
        self.visible = visible

    def count(self):
        return len(self.visible)

    def nth(self, index):
        return _VisibleText(self.visible[index])


class _Body:
    def __init__(self, text):
        self.text = text

    def inner_text(self, timeout):
        return self.text


class _Page:
    def __init__(self, text, print_visible):
        self.text = text
        self.print_visible = print_visible

    def locator(self, _selector):
        return _Body(self.text)

    def get_by_text(self, _text, exact):
        return _PrintControls(self.print_visible)


class EpostPortalPrintPopupTests(unittest.TestCase):
    def test_outer_component_is_preferred_over_text_child(self):
        self.assertEqual(
            preferred_outer_control_id(["btnPrint", "btnPrintTextBoxElement"]),
            "btnPrint",
        )

    def test_distinct_buttons_remain_ambiguous(self):
        self.assertEqual(preferred_outer_control_id(["btnPrint", "btnReset"]), "")

    def test_popup_requires_title_and_visible_print_button(self):
        self.assertTrue(print_popup_is_open(_Page("운송장출력(팝업)", [True])))
        self.assertFalse(print_popup_is_open(_Page("운송장출력(팝업)", [False])))


if __name__ == "__main__":
    unittest.main()
