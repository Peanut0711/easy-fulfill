"""우체국 전용 Chromium 로그인 상태 판정의 의존성 없는 검증."""

import unittest

from epost_portal_session import page_has_logged_in_state


class _Body:
    def __init__(self, body_text: str):
        self.body_text = body_text

    def inner_text(self, timeout):
        return self.body_text


class _Page:
    def __init__(self, body_text: str):
        self.body_text = body_text

    def locator(self, selector):
        self.selector = selector
        return _Body(self.body_text)


class EpostPortalSessionTests(unittest.TestCase):
    def test_logged_out_page_is_not_accepted(self):
        self.assertFalse(page_has_logged_in_state(_Page("로그인")))

    def test_logout_text_marks_connected_session(self):
        self.assertTrue(page_has_logged_in_state(_Page("HiGenies 로그아웃")))


if __name__ == "__main__":
    unittest.main()
