"""「네이버 문의」 탭 답변 등록(고객에게 보이는 쓰기)을 검증한다.

- API: 상품문의 PUT commentContent·고객문의 POST answerComment, 빈 응답도 성공, 잘못된 입력 거절,
  템플릿 목록, 답변 직전 다시 찾기
- 작업자: 이미 답변됨·못 찾음이면 보내지 않음, 401이면 새 토큰으로 처음부터 한 번 더, 보내는 중 연결
  끊김은 「알 수 없음」, 서버 거절은 확실한 실패
- 화면: 미답변만 답변 칸, 문의별 초안, 확인 창에서 취소하면 안 보냄, 성공·실패 처리, 템플릿 넣기,
  새로고침 뒤 답변됨이면 칸 숨김
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from integrations import naver_commerce  # noqa: E402

QNA = {"questionId": 101, "createDate": "2026-11-20T13:22:11.0+09:00", "productName": "유기농 쌀 10kg",
       "productId": 555, "maskedWriterId": "abc***", "question": "언제 발송되나요?", "answered": False}
CUSTOMER = {"inquiryNo": 202, "inquiryRegistrationDateTime": "2026-11-21T09:05:00.0+09:00",
            "productName": "현미 5kg", "orderId": "", "customerName": "홍*동", "title": "주소 변경",
            "inquiryContent": "주소를 바꾸고 싶어요", "answered": False}


class FakeResponse:
    def __init__(self, status=200, body=b""):
        self.status_code = status
        self.content = body
        self.text = body.decode("utf-8", "replace")

    def json(self):
        return json.loads(self.content.decode("utf-8"))


class AnswerApiTests(unittest.TestCase):
    def send(self, kind, inquiry_id, text, response=None):
        calls = []

        def fake_request(method, url, headers=None, json=None, timeout=None):
            calls.append((method, url, json, headers.get("Authorization")))
            return response or FakeResponse(200, b"")

        with patch.object(naver_commerce.requests, "request", fake_request):
            result = naver_commerce.post_inquiry_answer("tok", kind, inquiry_id, text)
        return result, calls

    def test_product_qna_uses_put_comment_content(self):
        result, calls = self.send("상품문의", "101", "  오늘 발송됩니다.  ")
        self.assertEqual(result, {})  # 빈 응답도 성공
        self.assertEqual(calls, [("PUT", naver_commerce.API_BASE + "/v1/contents/qnas/101",
                                  {"commentContent": "오늘 발송됩니다."}, "Bearer tok")])

    def test_customer_inquiry_uses_post_answer_comment(self):
        result, calls = self.send("고객문의", 202, "변경해 드렸습니다.", FakeResponse(200, b'{"data": 1}'))
        self.assertEqual(result, {"data": 1})
        self.assertEqual(calls[0][:3], ("POST", naver_commerce.API_BASE + "/v1/pay-merchant/inquiries/202/answer",
                                        {"answerComment": "변경해 드렸습니다."}))

    def test_success_with_non_json_body_is_success(self):
        result, _ = self.send("상품문의", "1", "네", FakeResponse(200, b"OK"))
        self.assertEqual(result, {})

    def test_rejects_bad_input_without_calling(self):
        for kind, inquiry_id, text in (("상품문의", "1", "  "), ("상품문의", "1/../2", "네"),
                                       ("상품문의", "", "네"), ("쿠팡", "1", "네")):
            with self.subTest(kind=kind, inquiry_id=inquiry_id, text=text):
                with self.assertRaises(ValueError):
                    self.send(kind, inquiry_id, text)

    def test_server_rejection_raises_with_message(self):
        with self.assertRaises(RuntimeError) as ctx:
            self.send("고객문의", "202", "네", FakeResponse(400, '{"code": "ERR", "message": "필수값 누락"}'.encode()))
        self.assertIn("HTTP 400", str(ctx.exception))
        self.assertIn("필수값 누락", str(ctx.exception))

    def test_templates(self):
        listed = [{"subject": "배송", "content": "오늘 발송됩니다.", "questionType": "DELIVERY"},
                  {"subject": "빈 것", "content": "  "}]
        with patch.object(naver_commerce, "_get_json", return_value=listed):
            self.assertEqual(naver_commerce.fetch_qna_templates("t"),
                             [{"subject": "배송", "content": "오늘 발송됩니다.", "question_type": "DELIVERY"}])
        with patch.object(naver_commerce, "_get_json", return_value={"contents": [{"content": "네"}]}):
            self.assertEqual(naver_commerce.fetch_qna_templates("t")[0]["subject"], "(제목 없음)")

    def test_find_inquiry(self):
        seen = []

        def qnas(token, start, end, answered=None):
            seen.append((start, answered))
            return [QNA, dict(QNA, questionId=102)]

        with patch.object(naver_commerce, "fetch_product_qnas", qnas):
            row = naver_commerce.find_inquiry("t", "상품문의", "102", "2026-11-20 13:22")
        self.assertEqual(row["id"], "102")
        self.assertEqual(seen, [(datetime(2026, 11, 19, 13, 22), None)])  # 등록 하루 전부터, 답변 여부 상관없이
        with patch.object(naver_commerce, "fetch_customer_inquiries", return_value=[CUSTOMER]):
            self.assertIsNone(naver_commerce.find_inquiry("t", "고객문의", "999", "2026-11-21 09:05"))


def run_scenario():
    import importlib.util
    import threading
    import time

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_answer", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    inquiries = sys.modules["ui.main.inquiries"]
    app = QApplication.instance() or QApplication([])
    result = {}

    # ── 작업자 ──
    calls = []
    state = {"current": None, "post_error": None, "reject_token": None}

    def fake_find(token, kind, inquiry_id, created):
        calls.append(["find", token])
        if state["reject_token"] == token:
            raise RuntimeError("HTTP 401: 토큰 만료")
        return state["current"]

    def fake_post(token, kind, inquiry_id, text):
        calls.append(["post", token, kind, inquiry_id, text])
        if state["post_error"]:
            raise state["post_error"]
        return {}

    def fake_new_token():
        module._remember_inquiry_token("new")
        return "new"

    worker = {}
    with patch.object(naver_commerce, "find_inquiry", fake_find), \
            patch.object(naver_commerce, "post_inquiry_answer", fake_post), \
            patch.object(module, "_naver_inquiry_token_from_config", fake_new_token):
        unanswered = naver_commerce.normalize_customer_inquiry(CUSTOMER)
        module._remember_inquiry_token("tok")
        state["current"] = dict(unanswered, answered=True)
        worker["already"] = [module.run_naver_inquiry_answer_worker("고객문의", "202", "2026-11-21 09:05", "네"),
                             list(calls)]
        del calls[:]
        state["current"] = None
        worker["missing"] = [module.run_naver_inquiry_answer_worker("고객문의", "202", "", "네")["ok"], list(calls)]
        del calls[:]
        state["current"] = unanswered
        worker["ok"] = [module.run_naver_inquiry_answer_worker("고객문의", "202", "", "네"), list(calls)]
        del calls[:]
        module._remember_inquiry_token("old")
        state["reject_token"] = "old"
        worker["retry"] = [module.run_naver_inquiry_answer_worker("고객문의", "202", "", "네")["ok"], list(calls)]
        state["reject_token"] = None
        del calls[:]
        state["post_error"] = TimeoutError("Read timed out")
        worker["timeout"] = module.run_naver_inquiry_answer_worker("고객문의", "202", "", "네")
        state["post_error"] = RuntimeError("HTTP 400: ERR 필수값 누락")
        worker["rejected"] = module.run_naver_inquiry_answer_worker("고객문의", "202", "", "네")
    result["worker"] = worker

    # ── 화면 ──
    rows = [naver_commerce.normalize_customer_inquiry(CUSTOMER), naver_commerce.normalize_product_qna(QNA),
            naver_commerce.normalize_product_qna(dict(QNA, questionId=103, answered=True, answer="네"))]
    list_responses = []
    list_calls = []
    answer_calls = []
    answer_results = []
    answer_gate = threading.Event()
    answer_gate.set()
    confirms = []
    confirm_answer = {"value": False}

    def fake_list(days, unanswered_only=True):
        list_calls.append(days)
        return list_responses.pop(0) if list_responses else {
            "ok": True, "rows": rows, "errors": [], "sample_keys": {}, "fetched_at": "10:00"}

    def fake_answer(kind, inquiry_id, created, text):
        answer_calls.append([kind, inquiry_id, created, text])
        answer_gate.wait(10)
        return answer_results.pop(0)

    def fake_confirm(parent, title, text, action_text, destructive=False, detail="", cancel_text="취소"):
        confirms.append({"title": title, "text": text, "action": action_text, "destructive": destructive,
                         "detail": detail})
        return confirm_answer["value"]

    def wait(predicate):
        deadline = time.time() + 10
        while time.time() < deadline:
            app.processEvents()
            if predicate():
                app.processEvents()
                return
            time.sleep(0.01)
        raise AssertionError("기다리는 상태가 오지 않음")

    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(module, "run_inquiry_list_worker", fake_list), \
            patch.object(module, "run_naver_inquiry_answer_worker", fake_answer), \
            patch.object(module, "run_naver_qna_templates_worker", lambda: {"ok": True, "templates": [
                {"subject": "배송 안내", "content": "오늘 발송됩니다.", "question_type": ""},
                {"subject": "재고", "content": "재고 있습니다.", "question_type": ""}]}), \
            patch.object(module, "run_naver_inquiry_order_worker",
                         lambda order_id, ids=(): {"ok": True, "order_id": order_id, "items": []}), \
            patch.object(module, "run_naver_inquiry_count_worker",
                         lambda: {"ok": True, "count": 0, "partial": False, "checked_at": "10:00"}), \
            patch.object(inquiries, "confirm_action", fake_confirm):
        window = module.MainWindow()
        window.show()
        app.processEvents()
        window.nav_rail_actions["tab_inquiries"].trigger()
        wait(lambda: window._inquiry_thread is None and window.tableWidget_inquiries.rowCount() == 3)
        table = window.tableWidget_inquiries
        card = window.card_inquiry_answer
        edit = window.plainTextEdit_inquiry_answer
        button = window.pushButton_inquiry_answer_send
        ui = {}
        table.selectRow(2)  # 답변된 상품문의
        app.processEvents()
        ui["answered_hidden"] = not card.isVisible()
        table.selectRow(1)  # 미답변 상품문의
        wait(lambda: window._inquiry_templates is not None)
        ui["unanswered_visible"] = card.isVisible()
        ui["public_note"] = [window.label_inquiry_answer_status.text(), window.label_inquiry_answer_status.toolTip()]
        ui["empty_disabled"] = not button.isEnabled()
        ui["templates"] = [window.comboBox_inquiry_template.itemText(i)
                           for i in range(window.comboBox_inquiry_template.count())]
        edit.setPlainText("안녕하세요. ")
        edit.moveCursor(edit.textCursor().MoveOperation.End)
        window.comboBox_inquiry_template.activated.emit(1)
        app.processEvents()
        ui["after_template"] = [edit.toPlainText(), window.comboBox_inquiry_template.currentIndex()]
        ui["typed_enabled"] = button.isEnabled()
        # 문의마다 초안을 따로 기억한다
        table.selectRow(0)
        app.processEvents()
        ui["other_draft"] = edit.toPlainText()
        ui["customer_note"] = [window.label_inquiry_answer_status.text(), window.label_inquiry_answer_status.toolTip()]
        edit.setPlainText("고객 답변 초안")
        table.selectRow(1)
        app.processEvents()
        ui["restored"] = edit.toPlainText()
        # 확인 창에서 취소 → 보내지 않는다
        button.click()
        app.processEvents()
        ui["cancel"] = [len(answer_calls), window._inquiry_answer_thread is None]
        ui["confirm"] = confirms[-1]
        # 보내면: 보내는 중 버튼 잠김 → 성공 → 초안 지움·목록 새로고침
        confirm_answer["value"] = True
        answer_gate.clear()
        answer_results.append({"ok": True, "kind": "상품문의", "id": "101"})
        lists_before = len(list_calls)
        answered_rows = [rows[0], dict(rows[1], answered=True, answer="안녕하세요. 오늘 발송됩니다."), rows[2]]
        list_responses.append({"ok": True, "rows": answered_rows, "errors": [], "sample_keys": {},
                               "fetched_at": "10:05"})
        button.click()
        app.processEvents()
        ui["sending"] = [button.isEnabled(), window.label_inquiry_answer_status.text()]
        answer_gate.set()
        wait(lambda: window._inquiry_answer_thread is None and window._inquiry_thread is None)
        ui["sent_call"] = answer_calls[-1]
        ui["after_success"] = {
            "draft": window._inquiry_drafts.get(("상품문의", "101")),
            "refreshed": len(list_calls) > lists_before,
            "hidden_after_refresh": not card.isVisible(),
            "notice": window.status_label.text(),
        }
        # 실패(연결 끊김)면 초안을 지키고 새로고침해 확인하라고 알린다
        table.selectRow(0)
        app.processEvents()
        answer_results.append({"ok": False, "kind": "고객문의", "id": "202", "error": "Read timed out",
                               "uncertain": True})
        button.click()
        wait(lambda: window._inquiry_answer_thread is None and window._inquiry_thread is None)
        ui["after_failure"] = {
            "draft": edit.toPlainText(),
            "status": window.label_inquiry_answer_status.text(),
            "tooltip": window.label_inquiry_answer_status.toolTip(),
            "visible": card.isVisible(),
        }
    result["ui"] = ui
    print("RESULT " + json.dumps(result, ensure_ascii=False, default=str), flush=True)
    os._exit(0)


_cached = None


def scenario():
    global _cached
    if _cached is None:
        with tempfile.TemporaryDirectory() as directory:
            env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
            completed = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), "--scenario"],
                cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=180,
            )
        line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
        if line is None:
            raise AssertionError(completed.stdout[-3000:] + completed.stderr[-3000:])
        _cached = json.loads(line[len("RESULT "):])
    return _cached


class AnswerWorkerTests(unittest.TestCase):
    def test_already_answered_is_not_sent(self):
        res, calls = scenario()["worker"]["already"]
        self.assertFalse(res["ok"])
        self.assertTrue(res["already_answered"])
        self.assertEqual(calls, [["find", "tok"]])

    def test_missing_is_not_sent(self):
        ok, calls = scenario()["worker"]["missing"]
        self.assertFalse(ok)
        self.assertEqual(calls, [["find", "tok"]])

    def test_sends_after_recheck(self):
        res, calls = scenario()["worker"]["ok"]
        self.assertEqual(res, {"kind": "고객문의", "id": "202", "ok": True})
        self.assertEqual(calls, [["find", "tok"], ["post", "tok", "고객문의", "202", "네"]])

    def test_expired_token_retries_from_recheck(self):
        ok, calls = scenario()["worker"]["retry"]
        self.assertTrue(ok)
        self.assertEqual(calls, [["find", "old"], ["find", "new"], ["post", "new", "고객문의", "202", "네"]])

    def test_timeout_is_uncertain_but_rejection_is_not(self):
        w = scenario()["worker"]
        self.assertFalse(w["timeout"]["ok"])
        self.assertTrue(w["timeout"]["uncertain"])
        self.assertFalse(w["rejected"]["ok"])
        self.assertFalse(w["rejected"]["uncertain"])
        self.assertIn("필수값 누락", w["rejected"]["error"])


class AnswerPanelTests(unittest.TestCase):
    def test_panel_only_for_unanswered(self):
        ui = scenario()["ui"]
        self.assertTrue(ui["answered_hidden"])
        self.assertTrue(ui["unanswered_visible"])
        self.assertEqual(ui["public_note"], ["상품 페이지에 공개", "상품문의 답변은 상품 페이지에 공개되어 다른 고객도 봅니다."])
        self.assertEqual(ui["customer_note"], ["고객에게만 보임", "고객문의 답변은 문의한 고객에게만 보입니다."])
        self.assertTrue(ui["empty_disabled"])

    def test_templates_and_drafts(self):
        ui = scenario()["ui"]
        self.assertEqual(ui["templates"], ["템플릿 넣기…", "배송 안내", "재고"])
        self.assertEqual(ui["after_template"], ["안녕하세요. 오늘 발송됩니다.", 0])
        self.assertTrue(ui["typed_enabled"])
        self.assertEqual(ui["other_draft"], "")
        self.assertEqual(ui["restored"], "안녕하세요. 오늘 발송됩니다.")

    def test_confirm_cancel_sends_nothing(self):
        ui = scenario()["ui"]
        self.assertEqual(ui["cancel"], [0, True])
        confirm = ui["confirm"]
        self.assertEqual(confirm["action"], "답변 등록")
        self.assertTrue(confirm["destructive"])  # 기본 버튼은 「취소」
        self.assertIn("상품 페이지에 공개", confirm["detail"])
        self.assertIn("안녕하세요. 오늘 발송됩니다.", confirm["detail"])
        self.assertIn("abc***", confirm["text"])

    def test_success(self):
        ui = scenario()["ui"]
        self.assertEqual(ui["sending"], [False, "보내는 중…"])
        self.assertEqual(ui["sent_call"], ["상품문의", "101", "2026-11-20 13:22", "안녕하세요. 오늘 발송됩니다."])
        after = ui["after_success"]
        self.assertIsNone(after["draft"])
        self.assertTrue(after["refreshed"])
        self.assertTrue(after["hidden_after_refresh"])  # 새로고침 뒤 답변됨 → 칸 숨김
        self.assertIn("답변을 등록했습니다", after["notice"])

    def test_uncertain_failure_keeps_draft(self):
        after = scenario()["ui"]["after_failure"]
        self.assertEqual(after["draft"], "고객 답변 초안")
        self.assertIn("등록됐는지 알 수 없습니다", after["status"])
        self.assertIn("새로고침", after["tooltip"])
        self.assertTrue(after["visible"])


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
