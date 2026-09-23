"""실제 쿠팡 요청 없이 CDN 재사용 및 로그인 판정 경로를 검증한다."""

import json
import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import coupang_cdn_upload as cdn
import coupang_session
import naver_to_coupang_html as batch


def response(status=400, url=cdn.UPLOAD_URL, content_type="application/json"):
    return SimpleNamespace(status=status, url=url, headers={"content-type": content_type}, dispose=Mock())


class SessionTests(unittest.TestCase):
    def test_response_classification(self):
        for status, url, content_type, expected in [
            (400, cdn.UPLOAD_URL, "application/json", "active"),
            (200, cdn.UPLOAD_URL, "application/json", "active"),
            (500, cdn.UPLOAD_URL, "application/json", "server_error"),
            (503, "https://xauth.coupang.com/login", "text/html", "server_error"),
            (403, cdn.UPLOAD_URL, "application/json", "blocked"),
            (403, "https://xauth.coupang.com/login", "text/html", "blocked"),
            (429, cdn.UPLOAD_URL, "text/html", "blocked"),
            (401, cdn.UPLOAD_URL, "application/json", "login_required"),
            (200, "https://xauth.coupang.com/login", "text/html", "login_required"),
            (200, "https://wing.coupang.com/sso/login", "text/html", "login_required"),
            (200, cdn.UPLOAD_URL, "text/html", "unknown"),
            (404, cdn.UPLOAD_URL, "application/json", "unknown"),
            (200, "https://other.example/upload", "application/json", "unknown"),
        ]:
            with self.subTest(status=status, url=url):
                self.assertEqual(cdn._upload_session_status(response(status, url, content_type)), expected)

    def test_blocked_request_does_not_prompt_or_navigate(self):
        page = Mock()
        result = response(403, "https://xauth.coupang.com/login?state=private-value", "text/html")
        page.context.request.get.return_value = result
        with self.assertRaisesRegex(RuntimeError, "접근 차단") as raised:
            cdn.wait_for_login(page)
        self.assertNotIn("private-value", str(raised.exception))
        page.goto.assert_not_called()
        result.dispose.assert_called_once()

    def test_network_error_does_not_prompt(self):
        page = Mock()
        page.context.request.get.side_effect = TimeoutError("network")
        with self.assertRaisesRegex(RuntimeError, "통신 오류"):
            cdn.wait_for_login(page)
        page.goto.assert_not_called()

    def test_sso_can_return_without_manual_login(self):
        page = Mock(url=cdn.WING_HOME)
        with patch.object(cdn, "_has_active_upload_session", side_effect=[False, True]), patch("builtins.print") as log:
            cdn.wait_for_login(page)
        page.goto.assert_called_once()
        self.assertFalse(any("로그인해 주세요" in str(call) for call in log.call_args_list))

    def test_context_refresh_fallback_and_cleanup(self):
        for active, failure in [(True, False), (False, False), (False, True)]:
            with self.subTest(active=active, failure=failure), ExitStack() as stack:
                hidden, visible = Mock(), Mock()
                hidden.pages, visible.pages = [Mock()], [Mock()]
                launch = stack.enter_context(patch.object(cdn, "launch_coupang_context", side_effect=[hidden, visible]))
                stack.enter_context(patch.object(cdn, "restore_coupang_session"))
                stack.enter_context(patch.object(cdn, "_has_active_upload_session", return_value=active))
                wait = stack.enter_context(patch.object(cdn, "wait_for_login", side_effect=RuntimeError("blocked") if failure else None))
                save = stack.enter_context(patch.object(cdn, "save_coupang_session"))
                if failure:
                    with self.assertRaisesRegex(RuntimeError, "blocked"):
                        cdn.launch_coupang_upload_context(Mock())
                    visible.close.assert_called_once()
                    save.assert_not_called()
                else:
                    context, page = cdn.launch_coupang_upload_context(Mock())
                    self.assertIs(context, hidden if active else visible)
                    save.assert_called_once_with(context, page)
                    context.close.assert_not_called()
                self.assertEqual(launch.call_count, 1 if active else 2)
                self.assertEqual(wait.call_count, 0 if active else 1)
                self.assertEqual(hidden.close.call_count, 0 if active else 1)

    def test_hidden_blocked_closes_without_visible_retry(self):
        context = Mock(pages=[Mock()])
        with patch.object(cdn, "launch_coupang_context", return_value=context) as launch, \
             patch.object(cdn, "restore_coupang_session"), \
             patch.object(cdn, "_has_active_upload_session", side_effect=RuntimeError("blocked")):
            with self.assertRaisesRegex(RuntimeError, "blocked"):
                cdn.launch_coupang_upload_context(Mock())
        launch.assert_called_once()
        context.close.assert_called_once()

    def test_snapshot_preserves_current_cookies_and_skips_expired(self):
        cookie = {"name": "sid", "value": "old", "domain": "wing.coupang.com", "path": "/", "expires": -1}
        snapshot = {"cookies": [cookie, {**cookie, "name": "expired", "expires": 1}, {**cookie, "name": "missing"}],
                    "origins": [{"origin": cdn.WING_HOME.rstrip("/"), "localStorage": [{"name": "key", "value": "old"}]}]}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.bin"
            path.write_bytes(json.dumps(snapshot).encode())
            context = Mock()
            context.cookies.return_value = [{**cookie, "value": "new"}]
            with patch.object(cdn, "AUTH_STATE_PATH", path), patch.object(cdn, "_unprotect_for_current_windows_user", side_effect=lambda data: data):
                self.assertTrue(cdn.restore_coupang_session(context))
            restored = context.add_cookies.call_args.args[0]
            self.assertEqual([item["name"] for item in restored], ["missing"])
            self.assertNotIn("expires", restored[0])
            self.assertIn("getItem(key) === null", context.add_init_script.call_args.args[0])

    def test_bad_snapshot_is_not_deleted(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.bin"
            path.write_bytes(b"existing")
            with patch.object(cdn, "AUTH_STATE_PATH", path), patch.object(cdn, "_unprotect_for_current_windows_user", side_effect=ValueError()):
                self.assertFalse(cdn.restore_coupang_session(Mock()))
            self.assertEqual(path.read_bytes(), b"existing")

    def test_hidden_save_keeps_previous_session_storage(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.bin"
            storage = {cdn.WING_HOME.rstrip("/"): {"key": "value"}}
            path.write_text(json.dumps({"easyFulfillSessionStorage": storage}), encoding="utf-8")
            context = Mock()
            context.storage_state.return_value = {"cookies": [{"name": "new-cookie"}], "origins": []}
            with patch.object(cdn, "AUTH_STATE_PATH", path), \
                 patch.object(cdn, "_unprotect_for_current_windows_user", side_effect=lambda data: data), \
                 patch.object(cdn, "_protect_for_current_windows_user", side_effect=lambda data: data):
                cdn.save_coupang_session(context, Mock(url="about:blank"))
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(saved["easyFulfillSessionStorage"], storage)
            self.assertEqual(saved["cookies"], [{"name": "new-cookie"}])

    def test_explicit_login_restores_before_check(self):
        context = Mock(pages=[Mock()])
        order = Mock()
        with patch("sys.argv", ["coupang_session.py", "--login"]), \
             patch("playwright.sync_api.sync_playwright"), \
             patch.object(cdn, "launch_coupang_context", return_value=context), \
             patch.object(cdn, "restore_coupang_session", order.restore), \
             patch.object(cdn, "wait_for_login", order.wait), \
             patch.object(cdn, "save_coupang_session", order.save), \
             patch.object(coupang_session, "seller_name", return_value=""):
            coupang_session.main()
        self.assertEqual([call[0] for call in order.mock_calls], ["restore", "wait", "save"])
        context.close.assert_called_once()


class CacheTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.image = {"index": 1, "file": "image.jpg", "source": "https://source.example/image.jpg"}
        self.saved = {**self.image, "cdnUrl": cdn.CDN_BASE + "test.jpg"}

    def product(self, number, cached=True, images=True):
        directory = self.root / number
        directory.mkdir()
        report = {"images": [self.image] if images else []}
        (directory / "report.json").write_text(json.dumps(report), encoding="utf-8")
        html = '<main><section class="text-block"><p>상품 설명</p></section>'
        if images:
            html += '<img src="images/image.jpg">'
        (directory / "coupang-preview.html").write_text(html + "</main>", encoding="utf-8")
        if cached:
            (directory / "coupang-cdn-progress.json").write_text(json.dumps([self.saved]), encoding="utf-8")
        return number, report

    def test_full_cache_and_no_images_do_not_start_playwright(self):
        prepared = [self.product("1"), self.product("2", images=False)]
        results = [{"productNo": number} for number, _ in prepared]
        browser = Mock(side_effect=AssertionError("browser must not start"))
        with patch.object(batch, "OUTPUT_ROOT", self.root), patch.object(batch, "load_coupang_upload_modules", return_value=(browser, cdn)):
            batch.upload(prepared, results)
        browser.assert_not_called()
        self.assertTrue(all(item["status"] == "completed" for item in results))
        html = (self.root / "1" / "coupang-paste.html").read_text(encoding="utf-8")
        self.assertIn(self.saved["cdnUrl"], html)
        self.assertNotIn('src="images/', html)

    def test_cache_key_requires_same_source_and_file(self):
        number, report = self.product("1")
        self.assertFalse(cdn.needs_image_upload(self.root / number, report))
        for key in ["source", "file"]:
            with self.subTest(key=key):
                changed = {"images": [{**self.image, key: "changed"}]}
                self.assertTrue(cdn.needs_image_upload(self.root / number, changed))

    def test_failed_connection_still_renders_cached_products_once(self):
        prepared = [self.product("1", cached=False), self.product("2"), self.product("3", cached=False)]
        results = [{"productNo": number} for number, _ in prepared]
        browser = Mock()
        browser.return_value.__enter__ = Mock(return_value=Mock())
        browser.return_value.__exit__ = Mock(return_value=False)
        with patch.object(batch, "OUTPUT_ROOT", self.root), \
             patch.object(batch, "load_coupang_upload_modules", return_value=(browser, cdn)), \
             patch.object(cdn, "launch_coupang_upload_context", side_effect=RuntimeError("blocked")) as launch:
            batch.upload(prepared, results)
        self.assertEqual([item["status"] for item in results], ["upload_failed", "completed", "upload_failed"])
        launch.assert_called_once()
        browser.assert_called_once()

    def test_cli_full_cache_does_not_start_playwright(self):
        self.product("1")
        with patch.object(cdn, "OUTPUT_ROOT", self.root), patch("sys.argv", ["coupang_cdn_upload.py", "1", "--upload"]), \
             patch.object(cdn, "sync_playwright", side_effect=AssertionError("browser must not start")) as browser:
            cdn.main()
        browser.assert_not_called()
        self.assertTrue((self.root / "1" / "coupang-paste.html").exists())

    def test_partial_cache_posts_only_new_image_and_preserves_order(self):
        number, report = self.product("1")
        second = {"index": 2, "file": "new.jpg", "source": "https://source.example/new.jpg"}
        report["images"].append(second)
        upload_file = self.root / "new.jpg"
        upload_file.write_bytes(b"mock-image")
        context = Mock()
        context.request.post.return_value = SimpleNamespace(ok=True)
        self.assertTrue(cdn.needs_image_upload(self.root / number, report))
        with patch.object(cdn, "prepare_image_for_upload", return_value=upload_file), \
             patch.object(cdn, "_read_upload_response", return_value={"success": True, "message": "new.jpg"}):
            mapping = cdn.upload_images(context, self.root / number, report)
        context.request.post.assert_called_once()
        self.assertEqual([item["cdnUrl"] for item in mapping], [self.saved["cdnUrl"], cdn.CDN_BASE + "new.jpg"])
        self.assertFalse(cdn.needs_image_upload(self.root / number, report))

    def test_batch_new_images_share_one_context_and_close(self):
        prepared = [self.product("1", cached=False), self.product("2", cached=False)]
        results = [{"productNo": number} for number, _ in prepared]
        browser = Mock()
        browser.return_value.__enter__ = Mock(return_value=Mock())
        browser.return_value.__exit__ = Mock(return_value=False)
        context = Mock()
        with patch.object(batch, "OUTPUT_ROOT", self.root), \
             patch.object(batch, "load_coupang_upload_modules", return_value=(browser, cdn)), \
             patch.object(cdn, "launch_coupang_upload_context", return_value=(context, Mock())) as launch, \
             patch.object(cdn, "upload_images", return_value=[self.saved]) as upload:
            batch.upload(prepared, results)
        launch.assert_called_once()
        context.close.assert_called_once()
        self.assertEqual(upload.call_count, 2)
        self.assertTrue(all(item["status"] == "completed" for item in results))


if __name__ == "__main__":
    unittest.main()
