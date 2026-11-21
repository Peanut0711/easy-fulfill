"""「네이버 문의」 탭(조회 전용)을 검증한다.

- 상품문의·고객문의 응답 → 탭 행 변환(필드 후보, 답변 여부, 날짜 줄이기), 최근 등록 순 정렬,
  한 종류 실패 시 다른 종류는 보여 줌, 상세 HTML 이스케이프
- 탭에 처음 들어가면 한 번 자동 조회, 목록·상세·상태 문구, 필터 변경 시 다시 조회,
  조회 중 필터 변경은 끝난 뒤 한 번 더, 실패·빈 결과 표시
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

QNA = {
    "questionId": 101,
    "createDate": "2026-11-20T13:22:11.0+09:00",
    "productName": "유기농 쌀 10kg",
    "productId": 555,
    "maskedWriterId": "abc***",
    "question": "언제 발송되나요?",
    "answered": False,
}
CUSTOMER = {
    "inquiryNo": 202,
    "inquiryRegistrationDateTime": "2026-11-21T09:05:00.0+09:00",
    "productName": "현미 5kg",
    "orderId": "2026112100001",
    "customerName": "홍*동",
    "category": "배송",
    "title": "주소 변경",
    "inquiryContent": "주소를 바꾸고 싶어요\n아파트 동 호수",
    "answered": True,
    "answerContent": "변경해 드렸습니다.",
}


class NormalizeTests(unittest.TestCase):
    def test_product_qna(self):
        row = naver_commerce.normalize_product_qna(QNA)
        self.assertEqual(row["kind"], "상품문의")
        self.assertEqual(row["id"], "101")
        self.assertEqual(row["created"], "2026-11-20 13:22")
        self.assertEqual(row["writer"], "abc***")
        self.assertEqual(row["product"], "유기농 쌀 10kg")
        self.assertEqual(row["product_no"], "555")
        self.assertEqual(row["content"], "언제 발송되나요?")
        self.assertIs(row["answered"], False)

    def test_customer_inquiry(self):
        row = naver_commerce.normalize_customer_inquiry(CUSTOMER)
        self.assertEqual(row["kind"], "고객문의")
        self.assertEqual(row["id"], "202")
        self.assertEqual(row["created"], "2026-11-21 09:05")
        self.assertEqual(row["order_id"], "2026112100001")
        self.assertEqual(row["title"], "주소 변경")
        self.assertIs(row["answered"], True)
        self.assertEqual(row["answer"], "변경해 드렸습니다.")

    def test_answered_unknown_unless_answer_present(self):
        bare = {"questionId": 1, "question": "q"}
        self.assertIsNone(naver_commerce.normalize_product_qna(bare)["answered"])
        with_answer = dict(bare, answer="네")
        self.assertIs(naver_commerce.normalize_product_qna(with_answer)["answered"], True)
        self.assertIs(naver_commerce.normalize_product_qna(dict(bare, answered="false"))["answered"], False)

    def test_product_qna_answer_time_comes_from_answers(self):
        # 실응답: answer(문자열)와 answers([{answer, createDate}]), 답변 시각은 answers에만 있다
        row = naver_commerce.normalize_product_qna(dict(
            QNA, answered=True, answer="오늘 발송됩니다.",
            answers=[{"answer": "오늘 발송됩니다.", "createDate": "2026-11-20T15:01:02.0+09:00"}]))
        self.assertIs(row["answered"], True)
        self.assertEqual(row["answer"], "오늘 발송됩니다.")
        self.assertEqual(row["answered_at"], "2026-11-20 15:01")
        only_list = naver_commerce.normalize_product_qna(
            {"questionId": 3, "answers": [{"answer": "네", "createDate": "2026-11-20T15:01:02.0+09:00"}]})
        self.assertEqual((only_list["answer"], only_list["answered"]), ("네", True))

    def test_customer_inquiry_answer_time(self):
        row = naver_commerce.normalize_customer_inquiry(
            dict(CUSTOMER, answerRegistrationDateTime="2026-11-21T10:00:00.0+09:00"))
        self.assertEqual(row["answered_at"], "2026-11-21 10:00")

    def test_short_datetime(self):
        self.assertEqual(naver_commerce.short_datetime("2026-11-20 13:22:11"), "2026-11-20 13:22")
        self.assertEqual(naver_commerce.short_datetime("어제"), "어제")
        self.assertEqual(naver_commerce.short_datetime(None), "")


class FetchListTests(unittest.TestCase):
    def call(self, qnas, customers, unanswered_only=True):
        calls = []

        def fake(result):
            def fn(token, from_dt, to_dt, answered=None):
                calls.append(answered)
                if isinstance(result, Exception):
                    raise result
                return result
            return fn

        with patch.object(naver_commerce, "fetch_product_qnas", fake(qnas)), \
                patch.object(naver_commerce, "fetch_customer_inquiries", fake(customers)):
            res = naver_commerce.fetch_inquiry_list("t", datetime(2026, 11, 1), datetime(2026, 11, 22),
                                                    unanswered_only=unanswered_only)
        return res, calls

    def test_sorted_newest_first_and_skips_rows_without_id(self):
        res, calls = self.call([QNA, {"question": "번호 없음"}], [CUSTOMER], unanswered_only=False)
        self.assertEqual([r["id"] for r in res["rows"]], ["202", "101"])
        self.assertEqual(calls, [None, None])
        self.assertEqual(res["errors"], [])
        self.assertIn("questionId", res["sample_keys"]["상품문의"])

    def test_unanswered_only_filters_and_defaults_to_unanswered(self):
        res, calls = self.call([{"questionId": 7, "question": "q"}], [], unanswered_only=True)
        self.assertEqual(calls, [False, False])
        self.assertIs(res["rows"][0]["answered"], False)

    def test_one_kind_failing_keeps_the_other(self):
        res, _ = self.call(RuntimeError("HTTP 403: 권한 없음"), [CUSTOMER])
        self.assertEqual([r["id"] for r in res["rows"]], ["202"])
        self.assertEqual(res["errors"], ["상품문의 조회 실패: HTTP 403: 권한 없음"])


class DetailHtmlTests(unittest.TestCase):
    def test_escapes_and_shows_answer(self):
        from ui.main.inquiries import inquiry_detail_html, inquiry_summary
        row = naver_commerce.normalize_customer_inquiry(dict(CUSTOMER, inquiryContent="<b>굵게</b>\n둘째 줄"))
        text = inquiry_detail_html(row)
        self.assertIn("&lt;b&gt;굵게&lt;/b&gt;<br>둘째 줄", text)
        self.assertIn("주문번호: 2026112100001", text)
        self.assertIn("변경해 드렸습니다.", text)
        self.assertEqual(inquiry_summary(row), "주소 변경")
        unanswered = naver_commerce.normalize_product_qna(QNA)
        self.assertIn("미답변", inquiry_detail_html(unanswered))
        self.assertEqual(inquiry_summary(unanswered), "언제 발송되나요?")


def run_scenario():
    import importlib.util
    import threading
    import time

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_inquiries", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])
    result = {"calls": []}
    rows = [naver_commerce.normalize_customer_inquiry(CUSTOMER), naver_commerce.normalize_product_qna(QNA)]
    responses = []
    gate = threading.Event()
    gate.set()

    def fake_worker(days, unanswered_only=True):
        result["calls"].append([days, unanswered_only])
        gate.wait(10)
        return responses.pop(0)

    def wait_idle(window):
        deadline = time.time() + 10
        while time.time() < deadline:
            app.processEvents()
            if window._inquiry_thread is None and not window._inquiry_refresh_pending:
                app.processEvents()
                return
            time.sleep(0.01)
        raise AssertionError("문의 조회가 끝나지 않음")

    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(module, "run_naver_inquiry_list_worker", fake_worker):
        window = module.MainWindow()
        window.show()
        app.processEvents()
        table = window.tableWidget_inquiries
        result["before"] = window.stackedWidget_inquiry_list.currentWidget() is window.emptyState_inquiry

        responses.append({"ok": True, "rows": rows, "errors": [], "sample_keys": {}, "fetched_at": "11:42"})
        window.nav_rail_actions["tab_inquiries"].trigger()
        wait_idle(window)
        result["first"] = {
            "calls": list(result["calls"]),
            "rows": table.rowCount(),
            "cells": [table.item(0, c).text() for c in range(table.columnCount())],
            "status": window.label_inquiry_status.text(),
            "shows_table": window.stackedWidget_inquiry_list.currentWidget() is table,
            "detail_empty": window.textBrowser_inquiry_detail.toPlainText() == "",
        }
        table.selectRow(1)
        app.processEvents()
        result["detail"] = window.textBrowser_inquiry_detail.toPlainText()

        # 다시 들어와도 자동 조회는 처음 한 번뿐
        window.nav_rail_actions["tab"].trigger()
        window.nav_rail_actions["tab_inquiries"].trigger()
        wait_idle(window)
        result["reenter_calls"] = len(result["calls"])

        # 필터 변경 → 다시 조회, 같은 문의 선택 유지, 일부 실패는 주의 표시
        responses.append({"ok": True, "rows": rows, "errors": ["상품문의 조회 실패: HTTP 403"],
                          "sample_keys": {}, "fetched_at": "11:50"})
        window.checkBox_inquiry_unanswered.setChecked(False)
        wait_idle(window)
        result["filter"] = {
            "call": result["calls"][-1],
            "status": window.label_inquiry_status.text(),
            "tooltip": window.label_inquiry_status.toolTip(),
            "selected": table.currentRow(),
            "detail_kept": "언제 발송되나요?" in window.textBrowser_inquiry_detail.toPlainText(),
        }

        # 조회 중 기간을 바꾸면 끝난 뒤 바뀐 기간으로 한 번 더 부른다
        gate.clear()
        responses.append({"ok": True, "rows": [], "errors": [], "sample_keys": {}, "fetched_at": "12:00"})
        responses.append({"ok": True, "rows": [], "errors": [], "sample_keys": {}, "fetched_at": "12:01"})
        window.refresh_inquiries()
        app.processEvents()
        result["button_while_loading"] = window.pushButton_inquiry_refresh.isEnabled()
        window.comboBox_inquiry_period.setCurrentIndex(0)
        gate.set()
        wait_idle(window)
        result["pending"] = {
            "calls": result["calls"][-2:],
            "empty_shown": window.stackedWidget_inquiry_list.currentWidget() is window.emptyState_inquiry,
            "empty_text": window.emptyState_inquiry.label.text(),
            "button": window.pushButton_inquiry_refresh.isEnabled(),
        }

        responses.append({"ok": False, "error": "네이버 client_id/secret 이 설정되지 않았습니다."})
        window.refresh_inquiries()
        wait_idle(window)
        result["error_status"] = window.label_inquiry_status.text()
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
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


class InquiriesTabTests(unittest.TestCase):
    def test_first_entry_loads_once(self):
        result = scenario()
        self.assertTrue(result["before"])
        first = result["first"]
        self.assertEqual(first["calls"], [[30, True]])
        self.assertEqual(first["rows"], 2)
        self.assertEqual(first["cells"], ["고객문의", "2026-11-21 09:05", "홍*동", "현미 5kg", "주소 변경"])
        self.assertEqual(first["status"], "✓ 미답변 2건 · 11:42 조회")
        self.assertTrue(first["shows_table"])
        self.assertTrue(first["detail_empty"])
        self.assertEqual(result["reenter_calls"], 1)

    def test_selecting_row_shows_detail(self):
        detail = scenario()["detail"]
        self.assertIn("[상품문의]", detail)
        self.assertIn("언제 발송되나요?", detail)
        self.assertIn("미답변", detail)

    def test_filter_change_reloads_and_keeps_selection(self):
        f = scenario()["filter"]
        self.assertEqual(f["call"], [30, False])
        self.assertEqual(f["status"], "! 문의 2건 · 11:50 조회 (일부 실패)")
        self.assertEqual(f["tooltip"], "상품문의 조회 실패: HTTP 403")
        self.assertEqual(f["selected"], 1)
        self.assertTrue(f["detail_kept"])

    def test_change_during_load_runs_again_after(self):
        result = scenario()
        self.assertFalse(result["button_while_loading"])
        pending = result["pending"]
        self.assertEqual(pending["calls"], [[30, False], [7, False]])
        self.assertTrue(pending["empty_shown"])
        self.assertEqual(pending["empty_text"], "이 기간에 들어온 문의가 없습니다.")
        self.assertTrue(pending["button"])

    def test_failure_is_shown(self):
        self.assertEqual(scenario()["error_status"],
                         "✗ 불러오지 못함: 네이버 client_id/secret 이 설정되지 않았습니다.")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
