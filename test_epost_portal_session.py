"""우체국 전용 Chromium 로그인 상태 판정의 의존성 없는 검증."""

import unittest

from epost_portal_session import (
    CONFIG_KEY_PORTAL_MEMBER_ID,
    CONFIG_KEY_PORTAL_PASSWORD,
    PortalCredentialError,
    page_has_logged_in_state,
    portal_credentials_from_settings,
    wait_for_logged_in_state,
)


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

    def test_logged_in_state_wait_accepts_ready_page(self):
        self.assertTrue(wait_for_logged_in_state(_Page("HiGenies 로그아웃"), 0))

    def test_portal_credentials_require_both_setting_values(self):
        credentials = portal_credentials_from_settings({
            CONFIG_KEY_PORTAL_MEMBER_ID: "portal-user",
            CONFIG_KEY_PORTAL_PASSWORD: "portal-password",
        })
        self.assertEqual(credentials.member_id, "portal-user")

        with self.assertRaises(PortalCredentialError) as error:
            portal_credentials_from_settings({CONFIG_KEY_PORTAL_MEMBER_ID: "portal-user"})
        self.assertIn(CONFIG_KEY_PORTAL_PASSWORD, str(error.exception))


if __name__ == "__main__":
    unittest.main()
