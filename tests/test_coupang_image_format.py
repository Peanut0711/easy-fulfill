"""실제 이미지 바이트로 파일 형식·MIME 일치와 변환·캐시 보존을 검증한다."""

import io
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from PIL import Image, features

import coupang_cdn_upload as cdn


class ImageFormatTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        (self.root / 'images').mkdir()
        # 다운로드 저장 방식처럼 확장자를 실제 형식과 무관하게 고정한다.
        self.image_path = self.root / 'images' / 'image-03.jpg'
        self.report = {'images': [{'index': 3, 'file': self.image_path.name, 'source': 'https://source/third'}]}
        self.context = Mock()
        self.context.request.post.side_effect = self.validate_upload

    def validate_upload(self, url, **kwargs):
        payload = kwargs['multipart']['multipartFile']
        with Image.open(io.BytesIO(payload['buffer'])) as uploaded:
            uploaded.load()
            expected = {'JPEG': ('.jpg', 'image/jpeg'), 'PNG': ('.png', 'image/png'),
                        'WEBP': ('.webp', 'image/webp'), 'BMP': ('.bmp', 'image/bmp'),
                        'GIF': ('.gif', 'image/gif')}.get(uploaded.format)
        valid = expected is not None and (Path(payload['name']).suffix, payload['mimeType']) == expected
        return SimpleNamespace(ok=True, status=200, status_text='OK', url=url,
                               headers={'content-type': 'application/json'}, dispose=Mock(),
                               text=lambda: json.dumps({'success': valid,
                                                        'message': 'test/third.jpg' if valid else '파일 형식이 맞지 않습니다.'}))

    def upload(self):
        return cdn.upload_images(self.context, self.root, self.report)

    def payload(self):
        return self.context.request.post.call_args.kwargs['multipart']['multipartFile']

    def test_supported_formats_keep_bytes_and_send_matching_extension_and_mime(self):
        for image_format in ('JPEG', 'PNG', 'GIF', 'WEBP', 'BMP'):
            with self.subTest(image_format=image_format):
                (self.root / 'coupang-cdn-progress.json').unlink(missing_ok=True)
                Image.new('RGB', (400, 300), 'red').save(self.image_path, image_format)
                original = self.image_path.read_bytes()
                self.context.request.post.reset_mock()
                self.assertEqual(self.upload()[0]['cdnUrl'], cdn.CDN_BASE + 'test/third.jpg')
                self.assertEqual(self.payload()['buffer'], original)
                self.assertEqual(self.image_path.read_bytes(), original)
                self.context.request.post.assert_called_once()

    def test_transparent_png_keeps_alpha_without_recompression(self):
        Image.new('RGBA', (400, 300), (255, 0, 0, 0)).save(self.image_path, 'PNG')
        original = self.image_path.read_bytes()
        self.upload()
        self.assertEqual(self.payload()['buffer'], original)
        self.assertEqual(self.payload()['name'], 'image-03.png')
        with Image.open(io.BytesIO(self.payload()['buffer'])) as uploaded:
            self.assertEqual(uploaded.getpixel((0, 0))[3], 0)

    def test_animated_gif_keeps_frames(self):
        first = Image.new('RGB', (400, 300), 'red')
        second = Image.new('RGB', (400, 300), 'blue')
        first.save(self.image_path, 'GIF', save_all=True, append_images=[second], duration=100, loop=0)
        original = self.image_path.read_bytes()
        self.upload()
        self.assertEqual(self.payload()['buffer'], original)
        with Image.open(io.BytesIO(self.payload()['buffer'])) as uploaded:
            self.assertEqual(uploaded.n_frames, 2)

    def test_tiff_converts_to_rgb_jpeg_without_changing_source(self):
        Image.new('RGBA', (400, 300), (255, 0, 0, 0)).save(self.image_path, 'TIFF')
        original = self.image_path.read_bytes()
        self.upload()
        self.assertEqual(self.image_path.read_bytes(), original)
        self.assertEqual((self.payload()['name'], self.payload()['mimeType']), ('image-03.jpg', 'image/jpeg'))
        with Image.open(io.BytesIO(self.payload()['buffer'])) as uploaded:
            self.assertEqual((uploaded.format, uploaded.mode, uploaded.size), ('JPEG', 'RGB', (400, 300)))
            self.assertEqual(uploaded.getpixel((0, 0)), (255, 255, 255))

    @unittest.skipUnless(features.check('avif'), '현재 Pillow가 AVIF를 지원하지 않습니다.')
    def test_avif_is_converted_even_when_dimensions_are_valid(self):
        Image.new('RGB', (400, 300), 'blue').save(self.image_path, 'AVIF')
        original = self.image_path.read_bytes()
        self.upload()
        self.assertEqual(self.image_path.read_bytes(), original)
        with Image.open(io.BytesIO(self.payload()['buffer'])) as uploaded:
            self.assertEqual((uploaded.format, uploaded.mode), ('JPEG', 'RGB'))

    def test_cmyk_jpeg_is_normalized_to_rgb(self):
        Image.new('CMYK', (400, 300)).save(self.image_path, 'JPEG')
        self.upload()
        with Image.open(io.BytesIO(self.payload()['buffer'])) as uploaded:
            self.assertEqual(uploaded.mode, 'RGB')

    def test_rotation_and_resizing_use_converted_dimensions(self):
        exif = Image.Exif()
        exif[274] = 6
        Image.new('RGB', (400, 200)).save(self.image_path, 'JPEG', exif=exif)
        self.upload()
        with Image.open(io.BytesIO(self.payload()['buffer'])) as uploaded:
            self.assertEqual(uploaded.size, (300, 600))
            self.assertNotIn(274, uploaded.getexif())

    def test_small_transparent_png_is_resized_to_white_rgb_jpeg(self):
        Image.new('RGBA', (100, 200), (0, 0, 0, 0)).save(self.image_path, 'PNG')
        self.upload()
        with Image.open(io.BytesIO(self.payload()['buffer'])) as uploaded:
            self.assertEqual((uploaded.format, uploaded.size), ('JPEG', (300, 600)))
            self.assertEqual(uploaded.getpixel((0, 0)), (255, 255, 255))

    def test_previously_uploaded_images_are_reused_when_third_image_is_converted(self):
        saved = [{'index': index, 'file': f'image-{index:02d}.jpg', 'source': f'https://source/{index}',
                  'cdnUrl': cdn.CDN_BASE + f'cached/{index}.jpg'} for index in (1, 2)]
        self.report['images'] = saved + self.report['images']
        (self.root / 'coupang-cdn-progress.json').write_text(json.dumps(saved), encoding='utf-8')
        Image.new('RGB', (400, 300)).save(self.image_path, 'TIFF')
        mapping = self.upload()
        self.assertEqual(mapping[:2], saved)
        self.context.request.post.assert_called_once()
        self.assertEqual([item['index'] for item in mapping], [1, 2, 3])
        self.assertEqual(len(json.loads((self.root / 'coupang-cdn-progress.json').read_text())), 3)

    def test_corrupt_image_is_rejected_before_network_request(self):
        self.image_path.write_bytes(b'not an image')
        with self.assertRaisesRegex(RuntimeError, '이미지 3 파일을 읽거나 변환하지 못했습니다'):
            self.upload()
        self.context.request.post.assert_not_called()

    def test_server_format_rejection_reports_actual_metadata_without_retry(self):
        Image.new('RGB', (400, 300)).save(self.image_path, 'PNG')
        response = self.validate_upload(cdn.UPLOAD_URL, multipart={'multipartFile': {
            'name': 'bad.jpg', 'mimeType': 'image/jpeg', 'buffer': self.image_path.read_bytes()}})
        self.context.request.post.side_effect = None
        self.context.request.post.return_value = response
        with patch.object(cdn.time, 'sleep') as sleep:
            with self.assertRaisesRegex(RuntimeError, r'HTTP 200.*image-03.png.*PNG/RGB.*image/png.*400x300'):
                self.upload()
        sleep.assert_not_called()
        self.context.request.post.assert_called_once()
        self.assertFalse((self.root / 'coupang-cdn-progress.json').exists())


if __name__ == '__main__':
    unittest.main()
