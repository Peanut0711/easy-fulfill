"""YouTube 추출부터 쿠팡 붙여넣기 HTML까지 외부 접속 없이 검증한다."""

import html
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import coupang_cdn_upload as cdn
import naver_detail_preview as detail


FIRST = "5RydDnx4VE4"
SECOND = "NaL3tsLzK0o"


def oembed(video_id, **overrides):
    data = {
        "inputUrl": f"https://www.youtube.com/watch?v={video_id}",
        "title": '제품 "시연" <영상>',
        "html": f'<iframe src=\\"https://www.youtube.com/embed/{video_id}?feature=oembed\\"></iframe>',
    }
    data.update(overrides)
    raw = html.escape(json.dumps({"type": "v2_oembed", "data": data}, ensure_ascii=False), quote=True)
    return ('<div class="se-component se-oembed se-l-default"><div></div>'
            f'<script type="text/data" data-module="{raw}" data-module-v2="{raw}"></script></div>')


class DetailVideoTests(unittest.TestCase):
    def test_supported_urls_and_untrusted_hosts(self):
        for url in (
            f"https://www.youtube.com/watch?v={FIRST}&feature=share",
            f"https://youtu.be/{FIRST}?si=abc",
            f"https://www.youtube.com/shorts/{FIRST}",
            f"https://m.youtube.com/live/{FIRST}",
            f"//www.youtube-nocookie.com/embed/{FIRST}",
        ):
            with self.subTest(url=url):
                self.assertEqual(detail.youtube_video_id(url), FIRST)
        for url in (
            f"https://youtube.com.evil.test/watch?v={FIRST}",
            f"https://www.youtube.com@evil.test/embed/{FIRST}",
            f"https://evil.test@www.youtube.com/embed/{FIRST}",
            f"javascript://www.youtube.com/embed/{FIRST}",
            "https://[bad", f"https://youtube.com:bad/embed/{FIRST}",
            f"https://www.youtube.com/embed/{FIRST}/extra", None, 123,
            "https://www.youtube.com/watch?v=short",
        ):
            with self.subTest(url=url):
                self.assertIsNone(detail.youtube_video_id(url))

    def build(self, directory, source):
        with patch.object(detail, "OUTPUT_ROOT", Path(directory)):
            return detail.build_preview("test", {"originProduct": {"name": "영상 상품", "detailContent": source}})

    def test_actual_oembed_shape_order_deduplication_and_cdn_output(self):
        source = ('<div class="se-component se-text"><p>상품 소개</p></div>'
                  + oembed(FIRST) + oembed(SECOND)
                  + '<div class="se-component se-image"><img src="https://example.com/board.jpg"></div>'
                  + '<div class="se-component se-text"><p>제품 사양</p></div>')
        metadata = {"width": 100, "height": 100, "bytes": 3, "format": "PNG"}
        with tempfile.TemporaryDirectory() as directory, patch.object(detail, "download_image", return_value=metadata):
            preview_path, report_path, report = self.build(directory, source)
            preview = preview_path.read_text(encoding="utf-8")
            mapping = [{"file": "image-01.jpg", "cdnUrl": "https://cdn.example.com/board.jpg"}]
            cdn_path, _, paste_path = cdn.write_cdn_html(preview_path.parent, mapping)
            for path in (preview_path, cdn_path, paste_path):
                result = path.read_text(encoding="utf-8")
                self.assertEqual(result.count("<iframe "), 2)
                self.assertLess(result.index("상품 소개"), result.index(FIRST))
                self.assertLess(result.index(FIRST), result.index(SECOND))
                self.assertLess(result.index(SECOND), result.index("<img "))
                self.assertLess(result.index("<img "), result.index("제품 사양"))
                self.assertIn('title="제품 &quot;시연&quot; &lt;영상&gt;"', result)
                self.assertNotIn("<script", result)
            self.assertIn("https://cdn.example.com/board.jpg", paste_path.read_text(encoding="utf-8"))
            self.assertEqual(json.loads(report_path.read_text(encoding="utf-8")), report)
        self.assertEqual(report["videoCount"], 2)
        self.assertEqual(report["imageCount"], 1)
        self.assertEqual(report["skippedVideoComponentCount"], 0)
        self.assertEqual([v["componentIndex"] for v in report["videos"]], [2, 3])
        self.assertNotIn("feature=oembed", preview)

    def test_embedded_html_fallback_and_direct_iframe(self):
        source = (oembed(FIRST, inputUrl="", title="")
                  + '<div class="se-component se-oembed">'
                  f'<iframe src="https://www.youtube.com/embed/{SECOND}" onload="bad()"></iframe></div>')
        with tempfile.TemporaryDirectory() as directory:
            path, _, report = self.build(directory, source)
            result = path.read_text(encoding="utf-8")
        self.assertEqual(report["videoCount"], 2)
        self.assertNotIn("onload", result)
        self.assertIn('title="제품 시연 영상"', result)

    def test_bad_and_unsupported_video_components_are_reported(self):
        source = ('<div class="se-component se-oembed"><script data-module="{bad}"></script></div>'
                  + oembed(FIRST, inputUrl="https://evil.test/abc", html='<iframe src="https://evil.test"></iframe>')
                  + '<div class="se-component se-video"></div>'
                  + oembed(FIRST, inputUrl=f"https://youtu.be/{FIRST}")
                  + oembed(FIRST))  # 다른 위치에 의도적으로 다시 삽입한 영상은 유지한다.
        with tempfile.TemporaryDirectory() as directory:
            path, _, report = self.build(directory, source)
            result = path.read_text(encoding="utf-8")
        self.assertEqual(report["videoCount"], 2)
        self.assertEqual(report["skippedVideoComponentCount"], 3)
        self.assertEqual(len(report["warnings"]), 4)
        self.assertNotIn("evil.test", result)


if __name__ == "__main__":
    unittest.main()
