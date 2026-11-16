"""네이버 상세 이미지를 병렬로 미리 받아도 순서·파일·오류가 순차 다운로드와 같은지 검증한다."""

import io
import tempfile
import threading
import time
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from product_detail import naver_detail_preview as detail

DELAY = 0.2


def png(width):
    buffer = io.BytesIO()
    Image.new("RGB", (width, 10), (width % 255, 0, 0)).save(buffer, "PNG")
    return buffer.getvalue()


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        time.sleep(DELAY)
        if self.path.startswith("/broken"):
            self.send_response(500)
            self.end_headers()
            return
        width = int(self.path.strip("/").split(".")[0])
        body = png(width)
        self.send_response(200)
        self.send_header("Content-Type", "image/png")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass


class PrefetchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        cls.base = f"http://127.0.0.1:{cls.server.server_address[1]}"
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def build(self, urls):
        source = '<div class="se-component se-text"><p>소개</p></div>' + "".join(
            f'<div class="se-component se-image"><img src="{url}"></div>' for url in urls)
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        with patch.object(detail, "OUTPUT_ROOT", Path(directory.name)):
            started = time.monotonic()
            result = detail.build_preview("p1", {"originProduct": {"name": "상품", "detailContent": source}})
            return result, time.monotonic() - started, Path(directory.name) / "p1" / "images"

    def test_parallel_download_keeps_order_and_is_faster(self):
        widths = list(range(20, 32))
        (_preview, _report_path, report), elapsed, image_dir = self.build([f"{self.base}/{w}.png" for w in widths])
        self.assertLess(elapsed, DELAY * len(widths) * 0.6)  # 순차(12×0.2초=2.4초)보다 빠름
        for index, width in enumerate(widths, 1):
            with Image.open(image_dir / f"image-{index:02d}.jpg") as image:
                self.assertEqual(image.width, width)
        self.assertEqual(sorted(p.name for p in image_dir.iterdir()),
                         [f"image-{index:02d}.jpg" for index in range(1, 13)])  # 미리 받기 임시 파일 없음
        self.assertEqual(report["imageCount"] if "imageCount" in report else len(widths), len(widths))

    def test_failed_image_raises_like_sequential_download(self):
        with self.assertRaises(Exception) as raised:
            self.build([f"{self.base}/20.png", f"{self.base}/broken.png"])
        self.assertIn("500", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
