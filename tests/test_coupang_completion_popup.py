"""실제 쿠팡 접속 없이 Chromium에서 완료 팝업 클릭과 종료 흐름을 검증한다."""

import contextlib
import io
import unittest
from unittest.mock import patch
from urllib.parse import quote

from playwright.sync_api import Error, sync_playwright

import coupang_detail_replace as detail


POPUP = '''<div><p>수정요청이 완료되었습니다.</p>
<button id="close" onclick="this.parentElement.remove()"><span>닫기</span></button>
<button id="list">상품목록</button></div>'''


class CompletionPopupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.browser = cls.playwright.chromium.launch(headless=True)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def setUp(self):
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.page.set_content(POPUP)
        self.requested = detail.watch_completion_popup_close(self.page)

    def tearDown(self):
        self.context.close()

    def wait_signal(self):
        for _ in range(20):
            if self.requested.is_set():
                return
            self.page.wait_for_timeout(50)
        self.fail('완료 팝업 닫기 신호를 받지 못했습니다.')

    def test_close_click_ends_context_and_emits_gui_marker(self):
        extra_tab = self.context.new_page()
        self.page.locator('#close span').click()
        self.wait_signal()
        self.assertEqual(self.page.locator('#close').count(), 0)
        output = io.StringIO()
        # 위에서 실제 브라우저 클릭으로 받은 신호를 대기 루프에 전달한다.
        with patch.object(detail, 'watch_completion_popup_close', return_value=self.requested), \
                patch.object(detail.sys, 'stdin', io.StringIO('')), contextlib.redirect_stdout(output):
            detail.wait_for_wing_close_or_finish(self.page, self.context, Error)
        self.assertTrue(self.page.is_closed())
        self.assertTrue(extra_tab.is_closed())
        self.assertIn(detail.WING_CLOSED_MARKER, output.getvalue())

    def test_product_list_does_not_request_close(self):
        self.page.locator('#list').click()
        self.page.wait_for_timeout(100)
        self.assertFalse(self.requested.is_set())

    def test_other_close_buttons_and_errors_are_ignored(self):
        for content in (
            '<div>수정요청에 실패했습니다.<button id="other">닫기</button></div>',
            POPUP + '<div><button id="other">닫기</button></div>',
            '<div><p hidden>수정요청이 완료되었습니다.</p><p>오류</p><button id="other">닫기</button></div>',
        ):
            with self.subTest(content=content):
                self.page.set_content(content)
                self.page.locator('#other').click()
                self.page.wait_for_timeout(100)
                self.assertFalse(self.requested.is_set())

    def test_script_click_is_ignored(self):
        self.page.locator('#close').evaluate('(button) => button.click()')
        self.page.wait_for_timeout(100)
        self.assertFalse(self.requested.is_set())

    def test_keyboard_activation_is_detected(self):
        self.page.locator('#close').focus()
        self.page.keyboard.press('Enter')
        self.wait_signal()

    def test_watcher_survives_navigation_and_spacing(self):
        content = POPUP.replace('수정요청이 완료되었습니다.', '수정 요청이\n완료되었습니다.')
        self.page.goto('data:text/html;charset=utf-8,' + quote(content))
        self.page.locator('#close').click()
        self.wait_signal()


if __name__ == '__main__':
    unittest.main()
