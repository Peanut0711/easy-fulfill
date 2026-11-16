"""묶음 비율·행 경계·CDN 변환과 PC/모바일 브라우저 배치를 검증한다."""

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from playwright.sync_api import sync_playwright

import coupang_cdn_upload as cdn
import naver_detail_preview as detail
from detail_image_layout import normalized_widths, render_image_group


def strip(widths, columns=2):
    return (f'<div class="se-imageStrip-container se-imageStrip-col-{columns}">' + ''.join(
        f'<div class="se-module se-module-image" style="width:{width}%;"><img src="https://example.com/{i}.jpg"></div>'
        for i, width in enumerate(widths)) + '</div>')


class ImageLayoutTests(unittest.TestCase):
    def test_original_widths_and_multiple_rows_survive_cdn_conversion(self):
        source = '<div class="se-component se-imageStrip">' + strip([84, 16]) + strip([46, 31, 23], 3) + '</div>'
        # 같은 정사각형 크기를 반환해 원본 지정 비율이 이미지 크기보다 우선하는지 확인한다.
        metadata = {"width": 100, "height": 100, "format": "PNG", "bytes": 1}
        with tempfile.TemporaryDirectory() as directory, patch.object(detail, 'OUTPUT_ROOT', Path(directory)), \
                patch.object(detail, 'download_image', return_value=metadata):
            preview, _, report = detail.build_preview('test', {'originProduct': {'name': '묶음', 'detailContent': source}})
            mapping = [{'file': item['file'], 'cdnUrl': f'https://cdn.example.com/{item["file"]}'} for item in report['images']]
            cdn_preview, _, paste = cdn.write_cdn_html(preview.parent, mapping)
            self.assertEqual([g['imageIndices'] for g in report['imageGroups']], [[1, 2], [3, 4, 5]])
            self.assertEqual([g['widthPercentages'] for g in report['imageGroups']], [[84, 16], [46, 31, 23]])
            for file in (preview, cdn_preview, paste):
                text = file.read_text(encoding='utf-8')
                self.assertEqual(text.count('data-ef-image-group='), 2)
                self.assertEqual(text.count('<img '), 5)
                self.assertIn('width:83.160000%', text)
                self.assertNotIn('padding:10px 12px', text.split('data-ef-image-group=', 1)[1])
            self.assertNotIn('src="images/', paste.read_text(encoding='utf-8'))

    def test_missing_invalid_widths_use_image_aspect_ratios(self):
        images = [{'width': 650, 'height': 230}, {'width': 233, 'height': 429}]
        for widths in (None, [None, 16], [float('nan'), 16], [-1, 16]):
            values = normalized_widths(images, widths)
            self.assertAlmostEqual(values[0], 83.87979785826015)
            self.assertAlmostEqual(sum(values), 100)

    def test_legacy_output_uses_report_sizes_and_preserves_regular_tables(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            (folder / 'coupang-preview.html').write_text(
                '<main><section class="image-block grid-2"><img src="images/a.jpg"><img src="images/b.jpg"></section>'
                '<section class="table-block"><table><tr><td>사양</td></tr></table></section></main>', encoding='utf-8')
            records = [{'file': 'a.jpg', 'width': 650, 'height': 230}, {'file': 'b.jpg', 'width': 233, 'height': 429}]
            (folder / 'report.json').write_text(json.dumps({'images': records}), encoding='utf-8')
            _, _, paste = cdn.write_cdn_html(folder, [{'file': x['file'], 'cdnUrl': 'https://cdn.example.com/' + x['file']} for x in records])
            text = paste.read_text(encoding='utf-8')
            self.assertIn('width:83.041000%', text)
            self.assertEqual(text.count('padding:10px 12px'), 1)
            self.assertNotIn('grid-2', text)

    def test_group_stays_in_one_row_on_pc_and_mobile(self):
        from urllib.parse import quote
        images = []
        for width, height in ((650, 230), (233, 429)):
            svg = f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}"><rect width="100%" height="100%" fill="gray"/></svg>'
            images.append({'src': 'data:image/svg+xml,' + quote(svg), 'width': width, 'height': height})
        source = cdn.render_paste_html('<main>' + render_image_group(images) + '</main>')
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            try:
                page = browser.new_page()
                for viewport in (900, 390):
                    page.set_viewport_size({'width': viewport, 'height': 700})
                    page.set_content(source, wait_until='load')
                    boxes = page.locator('img').evaluate_all('(images) => images.map(img => {const r=img.getBoundingClientRect(); return {x:r.x,y:r.y,w:r.width,h:r.height};})')
                    self.assertAlmostEqual(boxes[0]['y'], boxes[1]['y'], delta=1)
                    self.assertAlmostEqual(boxes[0]['h'], boxes[1]['h'], delta=1)
                    self.assertGreater(boxes[1]['x'], boxes[0]['x'] + boxes[0]['w'])
                    self.assertLessEqual(boxes[1]['x'] + boxes[1]['w'], viewport)
                    self.assertAlmostEqual(boxes[0]['w'] / boxes[1]['w'], (650 / 230) / (233 / 429), delta=.01)
            finally:
                browser.close()


if __name__ == '__main__':
    unittest.main()
