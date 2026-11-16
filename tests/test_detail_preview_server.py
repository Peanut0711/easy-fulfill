import tempfile
import unittest
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from detail_preview_server import DetailPreviewServer


class PreviewServerTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name) / 'detail.html'
        self.path.write_text('<p>저장된 본문</p>', encoding='utf-8')
        self.server = DetailPreviewServer(self.path)

    def tearDown(self):
        self.server.close()
        self.directory.cleanup()

    def read(self, url):
        with urlopen(url, timeout=2) as response:
            return response.read(), response.headers

    def test_live_saved_file_and_unsaved_editor_revision(self):
        self.assertIn('저장된 본문', self.read(self.server.url)[0].decode())
        self.path.write_text('<p>수정된 저장본</p>', encoding='utf-8')
        self.assertIn('수정된 저장본', self.read(self.server.url)[0].decode())
        first = self.server.update('<p>편집 중 😀</p>')
        second = self.server.update('<p>최종 편집</p>')
        self.assertNotEqual(first, second)
        body, headers = self.read(second)
        self.assertEqual(body.decode(), '<p>최종 편집</p>')
        self.assertEqual(headers['Referrer-Policy'], 'strict-origin-when-cross-origin')
        self.assertEqual(headers['Cache-Control'], 'no-store')
        self.assertEqual(self.path.read_text(encoding='utf-8'), '<p>수정된 저장본</p>')

    def test_only_scoped_images_are_served(self):
        folder = self.path.parent / 'images'
        folder.mkdir()
        (folder / '한글 사진.png').write_bytes(b'image')
        base = self.server.url.rsplit('/', 1)[0] + '/'
        from urllib.parse import quote
        self.assertEqual(self.read(base + quote('images/한글 사진.png'))[0], b'image')
        for suffix in ('detail.html', '../detail.html', '%2e%2e/detail.html', 'images/', 'missing.png'):
            with self.subTest(suffix=suffix), self.assertRaises(HTTPError) as error:
                self.read(base + suffix)
            self.assertEqual(error.exception.code, 404)
        outside = self.path.parent.parent / (self.path.parent.name + '.png')
        try:
            outside.write_bytes(b'outside')
            with self.assertRaises(HTTPError):
                self.read(base + '%2e%2e/' + outside.name)
        finally:
            outside.unlink(missing_ok=True)
        with self.assertRaises(HTTPError) as error:
            self.read(Request(self.server.url, headers={'Host': 'example.com'}))
        self.assertEqual(error.exception.code, 403)

    def test_close_releases_listener_and_is_idempotent(self):
        self.server.close()
        self.server.close()
        self.assertFalse(self.server._thread.is_alive())
        with self.assertRaises(URLError):
            self.read(self.server.url)


if __name__ == '__main__':
    unittest.main()
