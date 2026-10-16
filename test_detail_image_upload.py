"""PC 사진 업로드와 프로세스 오류 처리를 외부 요청 없이 검증한다."""

import io
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PIL import Image
from PySide6.QtWidgets import QApplication, QDialog
from PySide6.QtCore import QProcess, QTimer
import coupang_cdn_upload as cdn
import detail_image_upload as ui


class LocalImageTests(unittest.TestCase):
    def test_upload_converts_alpha_to_jpeg_and_leaves_existing_files_untouched(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            image_path = root / '사진.png'
            Image.new('RGBA', (400, 300), (255, 0, 0, 0)).save(image_path)
            original = image_path.read_bytes()
            html = root / 'coupang-paste.html'
            html.write_text('<p>기존 편집본</p>', encoding='utf-8')
            context = Mock()
            response = context.request.post.return_value
            response.ok = True
            response.text.return_value = json.dumps({'success': True, 'message': 'test/photo.jpg'})
            with patch.object(cdn, 'sync_playwright'), \
                    patch.object(cdn, 'launch_coupang_upload_context', return_value=(context, Mock())), \
                    patch.object(cdn, 'write_cdn_html') as render:
                result = cdn.upload_local_image(image_path)
            self.assertEqual(result['cdnUrl'], 'https://image.coupangcdn.com/image/test/photo.jpg')
            payload = context.request.post.call_args.kwargs['multipart']['multipartFile']
            self.assertEqual(payload['mimeType'], 'image/jpeg')
            with Image.open(io.BytesIO(payload['buffer'])) as uploaded:
                self.assertEqual(uploaded.format, 'JPEG')
                self.assertEqual(uploaded.size, (400, 300))
                self.assertEqual(uploaded.getpixel((0, 0)), (255, 255, 255))
            self.assertEqual(image_path.read_bytes(), original)
            self.assertEqual(html.read_text(encoding='utf-8'), '<p>기존 편집본</p>')
            self.assertEqual(sorted(p.name for p in root.iterdir()), ['coupang-paste.html', '사진.png'])
            render.assert_not_called()
            context.close.assert_called_once()

    def test_invalid_file_does_not_open_login(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'bad.jpg'
            path.write_text('not an image')
            with patch.object(cdn, 'sync_playwright') as browser:
                with self.assertRaises(Exception):
                    cdn.upload_local_image(path)
            browser.assert_not_called()

    def test_upload_failure_closes_browser(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'photo.jpg'
            Image.new('RGB', (400, 300)).save(path)
            context = Mock()
            with patch.object(cdn, 'sync_playwright'), \
                    patch.object(cdn, 'launch_coupang_upload_context', return_value=(context, Mock())), \
                    patch.object(cdn, 'upload_images', side_effect=RuntimeError('업로드 실패')):
                with self.assertRaisesRegex(RuntimeError, '업로드 실패'):
                    cdn.upload_local_image(path)
            context.close.assert_called_once()

    def test_phone_rotation_and_small_image_resize_before_upload(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'phone.jpg'
            exif = Image.Exif()
            exif[274] = 6
            Image.new('RGB', (200, 100)).save(path, exif=exif)
            original = path.read_bytes()
            context = Mock()
            def inspect_image(_context, output_dir, report):
                item = report['images'][0]
                with Image.open(output_dir / 'images' / item['file']) as image:
                    self.assertEqual(image.size, (300, 600))
                    self.assertNotIn(274, image.getexif())
                self.assertEqual((item['width'], item['height']), (300, 600))
                return [{'cdnUrl': 'https://image.coupangcdn.com/image/photo.jpg'}]
            with patch.object(cdn, 'sync_playwright'), \
                    patch.object(cdn, 'launch_coupang_upload_context', return_value=(context, Mock())), \
                    patch.object(cdn, 'upload_images', side_effect=inspect_image):
                cdn.upload_local_image(path)
            self.assertEqual(path.read_bytes(), original)

    def test_cli_local_upload_writes_only_result_without_product_lookup(self):
        with tempfile.TemporaryDirectory() as directory:
            result = Path(directory) / 'result.json'
            with patch.object(sys, 'argv', ['cdn', '--upload', '--local-image', '사진.png', '--result-json', str(result)]), \
                    patch.object(cdn, 'upload_local_image', return_value={'cdnUrl': 'https://image.coupangcdn.com/image/a.jpg'}) as upload, \
                    patch.object(cdn, 'load_report') as load:
                cdn.main()
            upload.assert_called_once_with(Path('사진.png'))
            load.assert_not_called()
            self.assertEqual(json.loads(result.read_text(encoding='utf-8'))['cdnUrl'],
                             'https://image.coupangcdn.com/image/a.jpg')


class UploadDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_process_result_and_cleanup(self):
        class SuccessfulDialog:
            def __init__(self, parent, image_path, result_path):
                self.error = None
                result_path.write_text(json.dumps({'cdnUrl': 'https://image.coupangcdn.com/image/photo.jpg'}))
            def exec(self):
                return QDialog.Accepted
            def deleteLater(self):
                pass
        with patch.object(ui, 'ImageUploadDialog', side_effect=SuccessfulDialog) as dialog:
            self.assertEqual(ui.upload_image_for_editor(None, Path('사진.png')),
                             'https://image.coupangcdn.com/image/photo.jpg')
        self.assertFalse(dialog.call_args.args[2].exists())

    def test_failed_process_and_start_errors_return_from_dialog(self):
        for failed in (False, True):
            with self.subTest(failed=failed), patch.object(QProcess, 'start'):
                dialog = ui.ImageUploadDialog(None, Path('사진.png'), Path('result.json'))
            if failed:
                QTimer.singleShot(0, lambda: dialog._process_error(QProcess.FailedToStart))
            else:
                QTimer.singleShot(0, lambda: dialog._finished(1, QProcess.NormalExit))
            self.assertEqual(dialog.exec(), QDialog.Rejected)
            self.assertTrue(dialog.error)
            dialog.deleteLater()

    def test_url_validation_rejects_local_and_other_hosts(self):
        for url in (None, 'images/photo.jpg', 'http://image.coupangcdn.com/image/p.jpg',
                    'https://image.coupangcdn.com.evil.example/image/p.jpg'):
            with self.subTest(url=url), self.assertRaises(ValueError):
                ui.validated_cdn_url(url)

    def test_real_child_process_success_keeps_dialog_responsive(self):
        original_start = QProcess.start
        observed = []
        def start(process, executable, arguments):
            self.assertEqual(executable, sys.executable)
            self.assertEqual(arguments[1:4], ['--upload', '--local-image', '사진.png'])
            result_path = arguments[-1]
            def try_escape():
                process.parent().reject()
                observed.append(process.parent().isVisible())
            QTimer.singleShot(10, try_escape)
            code = ('import pathlib,sys,time; time.sleep(0.15); '
                    'pathlib.Path(sys.argv[1]).write_text('
                    '\'{"cdnUrl":"https://image.coupangcdn.com/image/photo.jpg"}\')')
            original_start(process, executable, ['-c', code, result_path])
        with patch.object(QProcess, 'start', start):
            self.assertEqual(ui.upload_image_for_editor(None, Path('사진.png')),
                             'https://image.coupangcdn.com/image/photo.jpg')
        self.assertEqual(observed, [True])


if __name__ == '__main__':
    unittest.main()
