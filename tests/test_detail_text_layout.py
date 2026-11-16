"""원문 목록/문단 보존과 외부 CSS 환경에서의 쿠팡 본문 표시 회귀 검사."""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from playwright.sync_api import sync_playwright

import coupang_cdn_upload as cdn
import naver_detail_preview as detail
from detail_text_style import portable_text_html


SOURCE = (
    '<div class="se-component se-text"><div class="se-module-text">'
    '<p>상품 소개</p><p><span class="se-fs-fs19">\u200b</span></p>'
    '<p><span class="se-fs-fs19"><b>기본 정보</b></span></p>'
    '<ul class="se-text-list-type-bullet-disc"><li>'
    '<ul class="se-text-list-type-bullet-circle">'
    '<li><p><b>모델명</b>: EMT8150B</p></li>'
    '<li><p><b>센서 타입</b>: CMOS</p></li>'
    '<li><p>전류:</p><ul class="se-text-list-type-bullet-square">'
    '<li><p>평균 200mA @5V</p></li><li><p>최대 300mA @5V</p></li>'
    '</ul></li></ul></li></ul>'
    '<p class="se-text-paragraph-align-right"><span style="font-size:15px;color:#123456">자료</span></p>'
    '</div></div>'
)


def render(source):
    parser = detail.TreeParser()
    parser.feed(source)
    return detail.render_text_component(parser.root)


def tree(source):
    parser = detail.TreeParser()
    parser.feed(source)
    return parser.root


class DetailTextLayoutTests(unittest.TestCase):
    def test_nested_lists_preserve_every_item_and_order_without_duplicates(self):
        original, converted = tree(SOURCE), tree(render(SOURCE))
        for tag in ('ul', 'li', 'p'):
            before = [detail.node_text(n) for n in detail.walk(original) if n.tag == tag]
            after = [detail.node_text(n) for n in detail.walk(converted) if n.tag == tag]
            self.assertEqual(after, before)
        leaves = [detail.node_text(n) for n in detail.walk(converted) if n.tag == 'li'
                  and not any(c.tag in {'ul', 'ol'} for c in detail.walk(n))]
        self.assertEqual(leaves, ['모델명: EMT8150B', '센서 타입: CMOS', '평균 200mA @5V', '최대 300mA @5V'])

    def test_blank_paragraphs_linebreaks_and_emoji_sequences_survive(self):
        result = render('<div><p>A<br>B</p><p></p><p><span>\u200b</span></p>'
                        '<p><br></p><p>👩\u200d💻 &amp; C</p></div>')
        self.assertEqual(result.count('<br>'), 4)
        self.assertIn('👩\u200d💻 &amp; C', result)
        self.assertNotIn('\u200b', result)

    def test_font_size_does_not_turn_unrelated_first_paragraph_into_heading(self):
        result = render('<div><p>앞 문단</p><p><span class="se-fs-fs24">제목</span></p></div>')
        self.assertNotIn('<h1', result)
        self.assertIn('font-size:24px', result)
        self.assertEqual(result.count('<p>'), 2)

    def test_final_html_removes_icons_but_preserves_specs_links_and_layout(self):
        source = ('<p>🧩 상품 설명</p><p>⚙️ 주요 특징</p><p>📐 주요 사양</p>'
                  '<p>🔹 기본 정보 🔌 인터페이스 🔧 활용 예시 📦 구성품 ⚠️ 주의사항</p>'
                  '<p>👩\u200d💻 개발 • 1280 × 720 ≥ 5mil ± 1°C Ω 1D / 2D</p>'
                  '<ul><li>목록</li></ul><p>&#x1F9E9;&#xFE0F;상품 설명</p>'
                  '<a href="https://example.com/🧩">자료</a>')
        result = cdn.render_paste_html('<main>' + source + '</main>')
        paragraphs = [detail.node_text(n) for n in detail.walk(tree(result)) if n.tag == 'p']
        self.assertEqual(paragraphs, ['상품 설명', '주요 특징', '주요 사양',
            '기본 정보 인터페이스 활용 예시 구성품 주의사항',
            '개발 • 1280 × 720 ≥ 5mil ± 1°C Ω 1D / 2D', '상품 설명'])
        self.assertIn('href="https://example.com/🧩"', result)
        self.assertEqual(portable_text_html(result), result)
        edited = '<p style="color:red">🧩 제목</p><div style="height:24px">&nbsp;</div>'
        self.assertEqual(portable_text_html(edited, style_text=False), edited.replace('🧩 ', ''))

    def test_ordered_list_numbering_and_safe_source_styles(self):
        result = render('<div><ol start="3" reversed type="A"><li value="7"><p>일곱</p></li></ol>'
                        '<p class="se-text-paragraph-align-center" style="line-height:1.5;position:fixed">'
                        '<span class="se-fs-fs19" style="font-size:15px;color:#123456">색상</span>'
                        '<a href="javascript:bad()">문구</a></p></div>')
        self.assertIn('start="3" reversed', result)
        self.assertIn('value="7"', result)
        self.assertIn('list-style-type:upper-alpha', result)
        self.assertIn('text-align:center;line-height:1.5', result)
        self.assertIn('font-size:15px;color:#123456', result)
        self.assertNotIn('position', result)
        self.assertNotIn('javascript:', result)

    def test_table_cell_paragraphs_and_nested_lists_survive(self):
        root = tree('<div><table><tr><td><p>첫 줄</p><p>둘째 줄</p>'
                    '<ul><li>항목 A</li><li>항목 B</li></ul></td></tr></table></div>')
        result = detail.render_table_component(root)
        self.assertIn('<p>첫 줄</p><p>둘째 줄</p>', result)
        self.assertIn('<ul><li>항목 A</li><li>항목 B</li></ul>', result)

    def test_styling_is_idempotent_and_preserves_nontext_assets_and_attributes(self):
        source = ('<p data-note="style=wrong &gt;" style="color:red">A &amp; B</p>'
                  '<img src="x" style="width:30px"><iframe src="https://example.com"></iframe>')
        result = portable_text_html(source)
        self.assertEqual(portable_text_html(result), result)
        self.assertIn('<img src="x" style="width:30px"><iframe src="https://example.com"></iframe>', result)
        self.assertIn('data-note="style=wrong &gt;"', result)
        self.assertIn('color:red', result)

    def test_pc_mobile_and_host_css_reset_keep_structure_and_styles(self):
        # 쿠팡의 실제 CSS 복제가 아닌, 외부 사이트의 일반 초기화 CSS에 대한 내성 검사.
        reset = '<style>*{margin:0;padding:0;font:normal 12px/1 Arial;color:black;text-align:left}' \
                'ul,ol,li{list-style:none}p{display:inline}strong{font-weight:normal}</style>'
        with tempfile.TemporaryDirectory() as directory, patch.object(detail, 'OUTPUT_ROOT', Path(directory)):
            preview, _, _ = detail.build_preview('test', {'originProduct': {'detailContent': SOURCE}})
            full = preview.read_text(encoding='utf-8')
            paste = cdn.render_paste_html(full)
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            try:
                page = browser.new_page()
                for width in (900, 390):
                    page.set_viewport_size({'width': width, 'height': 900})
                    for host in ('', reset):
                        with self.subTest(width=width, reset=bool(host)):
                            page.set_content(host + paste)
                            paragraphs = page.locator('p').evaluate_all('els => els.map(e => ({text:e.textContent, y:e.getBoundingClientRect().y, height:e.getBoundingClientRect().height, display:getComputedStyle(e).display}))')
                            ys = [p['y'] for p in paragraphs]
                            self.assertEqual(ys, sorted(set(ys)))
                            self.assertTrue(all(p['height'] > 0 and p['display'] == 'block' for p in paragraphs))
                            self.assertEqual(page.locator('ul').evaluate_all('els=>els.map(e=>getComputedStyle(e).listStyleType)'), ['disc', 'circle', 'square'])
                            self.assertEqual(page.locator('li').first.evaluate('e=>getComputedStyle(e).listStyleType'), 'none')
                            self.assertEqual(page.locator('strong').first.evaluate('e=>getComputedStyle(e).fontWeight'), '700')
                            self.assertEqual(page.locator('span').last.evaluate('e=>getComputedStyle(e).fontSize'), '15px')
                            self.assertEqual(page.locator('p').last.evaluate('e=>getComputedStyle(e).textAlign'), 'right')
                            self.assertLessEqual(page.evaluate('document.documentElement.scrollWidth'), width)
            finally:
                browser.close()


if __name__ == '__main__':
    unittest.main()
