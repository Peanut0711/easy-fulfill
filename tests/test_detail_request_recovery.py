"""외부 접속 없이 업로드 재시도·보안 응답·이미지 404 복구를 검증한다."""

import io
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, call, patch

import requests
from PIL import Image

import coupang_cdn_upload as cdn
import naver_detail_preview as preview


def upload_response(status=200, body=None, content_type="application/json", **headers):
    if body is None:
        body = json.dumps({"success": True, "message": "test/image.jpg"})
    return SimpleNamespace(status=status, status_text="test", ok=200 <= status < 300,
                           url=cdn.UPLOAD_URL, text=Mock(return_value=body), dispose=Mock(),
                           headers={"content-type": content_type, **headers})


class UploadRecoveryTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        image_dir = self.root / "images"
        image_dir.mkdir()
        Image.new("RGB", (400, 300)).save(image_dir / "one.jpg")
        Image.new("RGB", (400, 300)).save(image_dir / "two.jpg")
        self.report = {"images": [{"index": 1, "file": "one.jpg", "source": "https://source/one.jpg"}]}
        self.context = Mock()
        self.sleep = patch.object(cdn.time, "sleep").start()
        self.addCleanup(patch.stopall)

    def upload(self):
        return cdn.upload_images(self.context, self.root, self.report)

    def test_transient_statuses_retry_same_image_then_save_progress(self):
        for status in (500, 502, 503, 504):
            with self.subTest(status=status):
                (self.root / "coupang-cdn-progress.json").unlink(missing_ok=True)
                failed = upload_response(status, "<html>Temporarily Unavailable</html>", "text/html")
                success = upload_response()
                self.context.request.post.reset_mock(side_effect=True)
                self.context.request.post.side_effect = [failed, success]
                self.sleep.reset_mock()
                mapping = self.upload()
                self.assertEqual(mapping[0]["cdnUrl"], cdn.CDN_BASE + "test/image.jpg")
                self.assertEqual(self.context.request.post.call_args_list[0], self.context.request.post.call_args_list[1])
                self.assertEqual(self.context.request.post.call_count, 2)
                self.sleep.assert_called_once_with(3)
                failed.dispose.assert_called_once()
                success.dispose.assert_called_once()
                self.assertEqual(len(json.loads((self.root / "coupang-cdn-progress.json").read_text())), 1)

    def test_exhaustion_preserves_completed_images_and_next_run_resumes(self):
        self.report["images"].append({"index": 2, "file": "two.jpg", "source": "https://source/two.jpg"})
        self.context.request.post.side_effect = [upload_response(), *[upload_response(503) for _ in range(3)]]
        with self.assertRaisesRegex(RuntimeError, "총 3회"):
            self.upload()
        self.assertEqual(self.context.request.post.call_count, 4)
        self.assertEqual(self.sleep.call_args_list, [call(3), call(8)])
        progress = json.loads((self.root / "coupang-cdn-progress.json").read_text())
        self.assertEqual([item["index"] for item in progress], [1])
        self.context.request.post.reset_mock(side_effect=True)
        self.context.request.post.return_value = upload_response()
        self.assertEqual([item["index"] for item in self.upload()], [1, 2])
        self.context.request.post.assert_called_once()

    def test_security_html_never_retries_or_writes_progress(self):
        for status, marker in ((503, '<div id="sec-overlay"></div>'),
                               (200, '<div class="g-recaptcha"></div>'),
                               (403, "Verify you are human")):
            with self.subTest(status=status):
                response = upload_response(status, f"<html>{marker}</html>", "text/html")
                self.context.request.post.reset_mock()
                self.context.request.post.return_value = response
                with self.assertRaisesRegex(RuntimeError, "보안 확인 필요.*쿠팡 로그인 연결/갱신"):
                    self.upload()
                self.context.request.post.assert_called_once()
                self.sleep.assert_not_called()
                response.dispose.assert_called_once()
                self.assertFalse((self.root / "coupang-cdn-progress.json").exists())

    def test_denied_and_invalid_responses_do_not_retry(self):
        for status, body, expected in ((403, "<html>Forbidden</html>", "접근 제한"),
                                       (429, "<html>Too many requests</html>", "접근 제한"),
                                       (401, "{}", "로그인 확인 필요"),
                                       (200, "<html>Unknown</html>", "JSON이 아닙니다"),
                                       (400, '{"success":false}', "업로드 실패")):
            with self.subTest(status=status, body=body):
                self.context.request.post.reset_mock()
                self.context.request.post.return_value = upload_response(status, body, "text/html")
                with self.assertRaisesRegex(RuntimeError, expected):
                    self.upload()
                self.sleep.assert_not_called()
                self.context.request.post.assert_called_once()

    def test_network_exception_is_not_repeated(self):
        self.context.request.post.side_effect = TimeoutError("timeout")
        with self.assertRaises(TimeoutError):
            self.upload()
        self.context.request.post.assert_called_once()
        self.sleep.assert_not_called()

    def test_retry_after_seconds_and_http_date_are_respected(self):
        for value in ("12", "Thu, 01 Jan 1970 00:00:12 GMT"):
            with self.subTest(value=value), patch.object(cdn.time, "time", return_value=0):
                send = Mock(side_effect=[upload_response(503, **{"retry-after": value}), upload_response()])
                self.sleep.reset_mock()
                result = cdn._request_upload_with_retry(send, label="업로드")
                self.sleep.assert_called_once_with(12)
                result.dispose()
        send = Mock(return_value=upload_response(503, **{"retry-after": "120"}))
        self.sleep.reset_mock()
        with self.assertRaisesRegex(RuntimeError, "120초"):
            cdn._request_upload_with_retry(send, label="업로드")
        self.sleep.assert_not_called()
        send.assert_called_once()

    def test_login_probe_retries_and_security_message_is_not_replaced_by_network_error(self):
        page = Mock()
        page.context.request.get.side_effect = [upload_response(503), upload_response()]
        self.assertTrue(cdn._has_active_upload_session(page))
        self.sleep.assert_called_once_with(3)
        page.context.request.get.side_effect = None
        page.context.request.get.return_value = upload_response(503, '<html id="sec-overlay"></html>', "text/html")
        with self.assertRaisesRegex(RuntimeError, "보안 확인 필요"):
            cdn._has_active_upload_session(page)

    def test_json_filename_with_captcha_is_not_a_security_page(self):
        self.context.request.post.return_value = upload_response(body='{"success":true,"message":"captcha.jpg"}')
        self.assertEqual(self.upload()[0]["cdnUrl"], cdn.CDN_BASE + "captcha.jpg")
        self.sleep.assert_not_called()


class DownloadRecoveryTests(unittest.TestCase):
    def setUp(self):
        buffer = io.BytesIO()
        Image.new("RGB", (10, 12)).save(buffer, format="PNG")
        self.image_bytes = buffer.getvalue()

    def response(self, status):
        response = Mock(status_code=status, content=self.image_bytes)
        if status >= 400:
            response.raise_for_status.side_effect = requests.HTTPError(f"HTTP {status}")
        return response

    def test_404_then_success_writes_complete_image(self):
        failed, success = self.response(404), self.response(200)
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(preview.requests, "get", side_effect=[failed, success]) as get, \
                patch.object(preview.time, "sleep") as sleep:
            path = Path(directory) / "image.jpg"
            metadata = preview.download_image("https://source/image", path)
            self.assertEqual(path.read_bytes(), self.image_bytes)
            self.assertEqual((metadata["width"], metadata["height"]), (10, 12))
            self.assertEqual(get.call_args_list, [call("https://source/image", timeout=30)] * 2)
            sleep.assert_called_once_with(3)
            failed.close.assert_called_once()
            success.close.assert_called_once()

    def test_persistent_404_preserves_existing_file_and_stops(self):
        responses = [self.response(404) for _ in range(3)]
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(preview.requests, "get", side_effect=responses) as get, \
                patch.object(preview.time, "sleep") as sleep:
            path = Path(directory) / "image.jpg"
            path.write_bytes(b"existing")
            with self.assertRaisesRegex(RuntimeError, "404가 총 3회"):
                preview.download_image("https://source/image", path)
            self.assertEqual(path.read_bytes(), b"existing")
            self.assertEqual(get.call_count, 3)
            self.assertEqual(sleep.call_args_list, [call(3), call(8)])
            for response in responses:
                response.close.assert_called_once()

    def test_download_403_is_not_retried(self):
        with patch.object(preview.requests, "get", return_value=self.response(403)) as get, \
                patch.object(preview.time, "sleep") as sleep:
            with self.assertRaises(requests.HTTPError):
                preview.download_image("https://source/image", Path("unused.jpg"))
            get.assert_called_once()
            sleep.assert_not_called()

    def test_product_not_found_is_not_retried(self):
        response = self.response(404)
        response.json.return_value = {"code": "NOT_FOUND", "message": "존재하지 않는 상품입니다."}
        with patch.object(preview, "load_config", return_value={"naver_client_id": "test", "naver_client_secret": "test"}), \
                patch.object(preview.naver_commerce, "get_access_token", return_value="test"), \
                patch.object(preview.requests, "get", return_value=response) as get, \
                patch.object(preview.time, "sleep") as sleep:
            with self.assertRaisesRegex(RuntimeError, "NOT_FOUND"):
                preview.fetch_product("15656231487")
            get.assert_called_once()
            sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
