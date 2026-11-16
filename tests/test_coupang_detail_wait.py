"""모의 WING 수정 화면으로 상세설명 HTML 탭 전환 대기(팝업·늦은 편집기·타임아웃)를 검증한다."""

import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.parse import parse_qs, urlparse

from playwright.sync_api import Error, sync_playwright

from product_detail import coupang_detail_replace as detail


# 쿠팡 WING 상품 수정 화면의 상세설명 영역을 흉내 낸다. 실제 WING처럼 탭 라디오는 선택 후에도
# checked가 false로 남고, 비활성 입력칸과 피드백 입력칸이 함께 있다.
# popup1/popup2: 탭 클릭 시 SweetAlert 확인 팝업 표시, editor_delay: HTML 편집기 표시 지연(ms),
# no_editor: 편집기를 끝내 표시하지 않음, basic_already: 이미 기본 상세라 HTML 작성 탭이 처음부터 보임.
MOCK_WING_PAGE = r"""<!doctype html><meta charset="utf-8"><title>상품 수정</title>
<style>.panel{display:none}.sweet-alert{display:none}</style>
<input type="radio" name="tab-content-level" id="tab-content-level-0"><label for="tab-content-level-0" id="basic">기본 등록</label>
<input type="radio" name="tab-content-level" id="tab-content-level-1"><label for="tab-content-level-1">옵션별 등록</label>
<div class="panel" id="basic-panel">
  <input type="radio" name="tab-content" id="tab-content-0"><label for="tab-content-0">이미지 업로드</label>
  <input type="radio" name="tab-content" id="tab-content-2"><label for="tab-content-2" id="html-tab">HTML 작성</label>
  <textarea class="sc-common-input" disabled></textarea>
  <div id="editor-host"></div>
</div>
<textarea placeholder="의견을 적어주세요." style="display:none"></textarea>
<textarea placeholder="현재 페이지에 대한 피드백을 적어주세요."></textarea>
<div class="sweet-alert"><p>변경하시겠습니까?</p><button class="confirm">확인</button></div>
<script>
const params = new URLSearchParams(location.search);
const flag = name => params.get(name) === '1';
const alertBox = document.querySelector('.sweet-alert');
let pending = null;
function openAlert(onConfirm) {
  pending = onConfirm;
  alertBox.style.display = 'block';
  alertBox.classList.add('showSweetAlert');
  alertBox.classList.remove('hideSweetAlert');
  setTimeout(() => alertBox.classList.add('visible'), Number(params.get('alert_delay') || 300));
}
alertBox.querySelector('button.confirm').addEventListener('click', () => {
  alertBox.classList.remove('showSweetAlert', 'visible');
  alertBox.classList.add('hideSweetAlert');
  setTimeout(() => { alertBox.style.display = 'none'; }, 100);
  const action = pending; pending = null; action && action();
});
if (flag('basic_already')) document.getElementById('basic-panel').style.display = 'block';
function chooseBasic() {
  setTimeout(() => { document.getElementById('basic-panel').style.display = 'block'; }, 150);
}
function chooseHtml() {
  if (flag('no_editor')) return;
  setTimeout(() => {
    const editor = document.createElement('textarea');
    editor.id = 'html-editor';
    document.getElementById('editor-host').appendChild(editor);
  }, Number(params.get('editor_delay') || 200));
}
document.getElementById('basic').addEventListener('click', event => {
  event.preventDefault();
  flag('popup1') ? openAlert(chooseBasic) : chooseBasic();
});
document.getElementById('html-tab').addEventListener('click', event => {
  event.preventDefault();
  flag('popup2') ? openAlert(chooseHtml) : chooseHtml();
});
</script>"""


def open_mock_page(browser, query):
    context = browser.new_context()
    page = context.new_page()
    page.route("https://wing.test/**", lambda route: route.fulfill(
        content_type="text/html; charset=utf-8", body=MOCK_WING_PAGE))
    page.goto(f"https://wing.test/modify?{query}")
    return context, page


class DetailWaitTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.browser = cls.playwright.chromium.launch(headless=True)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def fill(self, query):
        context, page = open_mock_page(self.browser, query)
        self.addCleanup(context.close)
        started = time.monotonic()
        textarea = detail.fill_detail_html(page, "123", "<p>상세</p>", Error, [])
        elapsed = time.monotonic() - started
        self.assertEqual(textarea.get_attribute("id"), "html-editor")
        self.assertEqual(page.input_value("#html-editor"), "<p>상세</p>")
        for placeholder in detail.FEEDBACK_PLACEHOLDERS:
            self.assertEqual(page.input_value(f'textarea[placeholder="{placeholder}"]'), "")
        self.assertFalse(page.locator(".sweet-alert.showSweetAlert").count())
        return page, elapsed

    def test_without_popup_proceeds_as_soon_as_editor_is_ready(self):
        _page, elapsed = self.fill("")
        # 기존 경로는 팝업 대기 3초 + 고정 1초를 두 번 기다려 8초 이상 걸렸다.
        self.assertLess(elapsed, 3.0)

    def test_confirms_popups_on_both_tabs(self):
        self.fill("popup1=1&popup2=1")

    def test_already_basic_product_confirms_html_mode_popup(self):
        # 실제 WING(등록상품 15120489904)에서 관찰한 흐름: 기본 등록은 변화 없음, HTML 작성만 확인 팝업.
        _page, elapsed = self.fill("basic_already=1&popup2=1&alert_delay=550")
        self.assertLess(elapsed, 3.0)

    def test_waits_for_late_editor(self):
        _page, elapsed = self.fill("editor_delay=2500")
        self.assertGreaterEqual(elapsed, 2.4)

    def test_step_returns_number_of_confirmed_popups(self):
        context, page = open_mock_page(self.browser, "popup1=1")
        self.addCleanup(context.close)
        detail.select_basic_detail_level(page, "123", Error, [])
        self.assertEqual(detail.wait_for_detail_step(page, "level", "123", Error, []), 1)

    def test_timeout_saves_diagnostics_with_stage(self):
        context, page = open_mock_page(self.browser, "no_editor=1&token=hidden")
        self.addCleanup(context.close)
        detail.select_basic_detail_level(page, "123", Error, [])
        detail.wait_for_detail_step(page, "level", "123", Error, [])
        detail.select_html_detail_tab(page, "123", Error, [])
        with tempfile.TemporaryDirectory() as directory, patch.object(detail, "ROOT", Path(directory)):
            started = time.monotonic()
            with self.assertRaisesRegex(RuntimeError, "HTML 편집기 준비 상태를 0.6초 안에") as raised:
                detail.wait_for_detail_step(page, "editor", "123", Error, [], timeout=600)
            self.assertLess(time.monotonic() - started, 5)
            self.assertIsInstance(raised.exception.__cause__, Error)
            files = list((Path(directory) / "output" / "wing-diagnostics").glob("*.json"))
            self.assertEqual(len(files), 1)
            raw = files[0].read_text(encoding="utf-8")
            record = json.loads(raw)
            self.assertEqual(record["stage"], "HTML 편집기 준비")
            self.assertEqual(parse_qs(urlparse(record["url"]).query), {})
            self.assertNotIn("hidden", raw)
            self.assertIn(str(files[0]), str(raised.exception))


if __name__ == "__main__":
    unittest.main()
