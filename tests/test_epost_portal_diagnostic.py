"""우체국 포털 읽기 전용 진단의 개인정보 제외·화면 판정 검증."""

import unittest

from epost_portal_diagnostic import looks_like_print_page, sanitized_page_path


class _Locator:
    def __init__(self, page, selector):
        self.page = page
        self.selector = selector

    def inner_text(self, timeout):
        return self.page.body_text

    def evaluate_all(self, script):
        return self.page.controls


class _Page:
    def __init__(self, body_text, controls):
        self.body_text = body_text
        self.controls = controls

    def locator(self, selector):
        return _Locator(self, selector)


class EpostPortalDiagnosticTests(unittest.TestCase):
    def test_query_string_is_not_saved(self):
        self.assertEqual(
            sanitized_page_path("https://biz.epost.go.kr/ui/index.jsp?token=secret"),
            "https://biz.epost.go.kr/ui/index.jsp",
        )

    def test_print_page_requires_heading_search_conditions_and_inputs(self):
        controls = [{"text": "", "tag": "input"} for _ in range(4)]
        page = _Page("운송장출력 검색일자 발송지 출력대상", controls)
        self.assertTrue(looks_like_print_page(page))

        self.assertFalse(looks_like_print_page(_Page("운송장출력 검색일자", controls)))


if __name__ == "__main__":
    unittest.main()
