"""우체국 포털 로그인 확인이 세션 만료 시 12초를 기다리지 않고, 유효한 세션은 그대로 쓰는지 검증한다."""

import time
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from playwright.sync_api import sync_playwright

from shipping import epost_portal_session as session

USER = session.LOGIN_MEMBER_ID_SELECTOR.lstrip("#")
PASSWORD = session.LOGIN_PASSWORD_SELECTOR.lstrip("#")
BUTTON = session.LOGIN_BUTTON_SELECTOR.lstrip("#")

# mode=expired: 로그인 화면만 표시, mode=valid: 1.5초 뒤 로그아웃 표식 표시,
# mode=flash: 로그인 칸이 0.5초 보였다가 세션 인식으로 로그아웃 표식으로 바뀜.
MOCK_PORTAL = f"""<!doctype html><meta charset="utf-8"><body><div id="root"></div>
<script>
const mode = new URLSearchParams(location.search).get('mode');
const root = document.getElementById('root');
function loginForm() {{
  root.innerHTML = '<input id="{USER}"><input id="{PASSWORD}" type="password"><button id="{BUTTON}">로그인</button>';
  document.getElementById('{BUTTON}').onclick = () => {{ root.innerHTML = '<a>로그아웃</a>'; }};
}}
if (mode === 'expired') setTimeout(loginForm, 300);
if (mode === 'valid') setTimeout(() => {{ root.innerHTML = '<a>로그아웃</a>'; }}, 1500);
if (mode === 'flash') {{ loginForm(); setTimeout(() => {{ root.innerHTML = '<a>로그아웃</a>'; }}, 500); }}
</script>"""


class PortalLoginWaitTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.browser = cls.playwright.chromium.launch(headless=True)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def login(self, mode):
        context = self.browser.new_context()
        self.addCleanup(context.close)
        page = context.new_page()
        page.route("https://portal.test/**", lambda route: route.fulfill(
            content_type="text/html; charset=utf-8", body=MOCK_PORTAL))
        page.goto(f"https://portal.test/?mode={mode}")
        credentials = SimpleNamespace(member_id="user", password="pw")
        with patch.object(session, "load_portal_credentials", return_value=credentials):
            started = time.monotonic()
            result = session.ensure_portal_login(page, 30)
        return result, time.monotonic() - started

    def test_expired_session_logs_in_without_waiting_twelve_seconds(self):
        result, elapsed = self.login("expired")
        self.assertEqual(result, "config_credentials")
        self.assertLess(elapsed, 5)  # 수정 전에는 로그아웃 표식을 12초 기다린 뒤 로그인

    def test_valid_session_rendered_late_is_reused(self):
        result, elapsed = self.login("valid")
        self.assertEqual(result, "saved_session")
        self.assertLess(elapsed, 5)

    def test_login_form_flashing_before_session_is_not_used(self):
        result, _elapsed = self.login("flash")
        self.assertEqual(result, "saved_session")


if __name__ == "__main__":
    unittest.main()
