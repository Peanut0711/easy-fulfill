"""로컬 Chromium 화면으로 WING 탭 실패 진단을 검증한다."""

import json
import tempfile
import unittest
from collections import deque
from pathlib import Path
from unittest.mock import Mock, patch

from playwright.sync_api import Error, TimeoutError, sync_playwright

import coupang_detail_replace as detail


class DetailDiagnosticTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.browser = cls.playwright.chromium.launch(headless=True)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def test_missing_tab_records_page_popup_http_error_and_screenshot(self):
        with tempfile.TemporaryDirectory() as directory:
            context = self.browser.new_context()
            self.addCleanup(context.close)
            page = context.new_page()
            events = detail.start_detail_diagnostics(page)
            def serve(route):
                if route.request.url.endswith('/api?token=hidden'):
                    route.fulfill(status=503, body='Unavailable')
                else:
                    route.fulfill(content_type='text/html; charset=utf-8', body='''<title>상품 수정</title>
                        <input id="tab-content-level-0" type="radio" checked>
                        <label for="tab-content-level-0">기본 상세</label>
                        <div role="dialog">상세 설정을 확인하세요</div>
                        <script>fetch('/api?token=hidden'); console.error('token=secret');</script>''')
            page.route('https://wing.test/**', serve)
            page.goto('https://wing.test/modify?token=private')
            page.wait_for_function("document.readyState === 'complete'")
            # 로컬 페이지에서 실제 클릭 실패를 발생시키되 15초 대기는 생략한다.
            class FastPage:
                def __getattr__(self, name):
                    return getattr(page, name)
                def locator(self, selector):
                    target = page.locator(selector)
                    return Mock(click=lambda **kwargs: target.click(timeout=50))
            with patch.object(detail, 'ROOT', Path(directory)):
                with self.assertRaisesRegex(RuntimeError, 'HTML 작성 탭을 선택하지 못했습니다'):
                    detail.select_html_detail_tab(FastPage(), '123', Error, events)
            files = list((Path(directory) / 'output' / 'wing-diagnostics').glob('*.json'))
            self.assertEqual(len(files), 1)
            raw = files[0].read_text(encoding='utf-8')
            record = json.loads(raw)
            self.assertEqual(record['url'], 'https://wing.test/modify')
            self.assertEqual(record['vendorInventoryId'], '123')
            self.assertEqual(record['stage'], 'HTML 작성 탭 선택')
            self.assertEqual(record['dom']['tabs'][0]['for'], 'tab-content-level-0')
            self.assertTrue(record['dom']['tabs'][0]['checked'])
            self.assertIn('상세 설정을 확인하세요', record['dom']['dialogs'])
            self.assertTrue(any(event.get('status') == 503 for event in record['events']))
            self.assertTrue(Path(record['screenshot']).exists())
            for secret in ('hidden', 'private', 'secret'):
                self.assertNotIn(secret, raw)

    def test_closed_page_still_writes_diagnostic_json(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(detail, 'ROOT', Path(directory)):
            page = self.browser.new_page()
            page.close()
            path = detail.save_detail_diagnostics(page, '123', 'HTML 작성 탭 선택', deque())
            self.assertIsNotNone(path)
            record = json.loads(path.read_text(encoding='utf-8'))
            self.assertTrue(record['captureErrors'])
            self.assertNotIn('screenshot', record)

    def test_storage_failure_preserves_original_click_error(self):
        page = Mock(url='https://wing.test/modify?token=hidden')
        failure = TimeoutError('missing tab')
        page.locator.return_value.click.side_effect = failure
        page.evaluate.side_effect = Error('closed')
        with patch.object(Path, 'mkdir', side_effect=OSError('disk error')):
            with self.assertRaisesRegex(RuntimeError, '진단 파일: 저장 실패') as raised:
                detail.select_html_detail_tab(page, '123', Error)
        self.assertIs(raised.exception.__cause__, failure)

    def test_basic_tab_failure_also_captures_diagnostics(self):
        page = Mock()
        page.locator.return_value.click.side_effect = TimeoutError('missing basic tab')
        with patch.object(detail, 'save_detail_diagnostics') as capture:
            with self.assertRaisesRegex(RuntimeError, '상세설명 탭을 찾지 못했습니다'):
                detail.select_basic_detail_level(page, '123', Error, [])
        capture.assert_called_once_with(page, '123', '기본 상세 선택', [])

    def test_successful_html_tab_click_does_not_capture_or_fill(self):
        page = Mock()
        with patch.object(detail, 'save_detail_diagnostics') as capture:
            detail.select_html_detail_tab(page, '123', Error)
        capture.assert_not_called()
        page.locator.assert_called_once_with('label[for="tab-content-2"]')
        page.locator.return_value.click.assert_called_once_with(timeout=15_000)
        page.locator.return_value.fill.assert_not_called()


if __name__ == '__main__':
    unittest.main()
