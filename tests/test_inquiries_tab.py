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


def order_detail(poid, status="PURCHASE_DECIDED", **extra):
    detail = {
        "order": {"orderId": "2026112100001", "orderDate": "2026-11-20T10:00:00.0+09:00"},
        "productOrder": {
            "productOrderId": poid, "productOrderStatus": status, "placeOrderStatus": "OK",
            "productName": "현미 5kg", "productOption": "5kg", "quantity": 2,
            "expectedDeliveryMethod": "DELIVERY", "shippingDueDate": "2026-11-22T23:59:59.0+09:00",
            "shippingMemo": "문 앞",
            "shippingAddress": {"name": "홍길동", "zipCode": "06000", "baseAddress": "서울 강남구 테헤란로 1",
                                "detailedAddress": "101호", "tel1": "010-0000-0000"},
        },
        "delivery": {"deliveryCompany": "EPOST", "trackingNumber": "6012345678901",
                     "deliveryStatus": "DELIVERY_COMPLETION", "deliveryMethod": "DELIVERY",
                     "sendDate": "2026-11-20T15:00:00.0+09:00", "deliveredDate": "2026-11-21T12:30:00.0+09:00",
                     "isWrongTrackingNumber": False},
    }
    for key, value in extra.items():
        if key == "po":
            detail["productOrder"].update(value)
        else:
            detail[key] = value
    return detail


class OrderSummaryTests(unittest.TestCase):
    def test_delivered_order(self):
        s = naver_commerce.summarize_product_order(order_detail("P1"), ["P1"])
        self.assertTrue(s["is_inquiry_target"])
        self.assertEqual(s["status"], "구매확정")
        self.assertEqual((s["company"], s["tracking"], s["delivery_status"]),
                         ("우체국택배", "6012345678901", "배송완료"))
        self.assertEqual((s["send_date"], s["delivered_date"]), ("2026-11-20 15:00", "2026-11-21 12:30"))
        self.assertEqual(s["address"], "서울 강남구 테헤란로 1 101호")
        self.assertEqual((s["receiver"], s["zip"], s["memo"]), ("홍길동", "06000", "문 앞"))
        self.assertEqual(s["order_date"], "2026-11-20 10:00")
        self.assertEqual(s["delivery_method"], "")
        self.assertNotIn("010-0000-0000", json.dumps(s, ensure_ascii=False))  # 연락처는 넣지 않는다

    def test_not_yet_shipped_and_claims(self):
        waiting = naver_commerce.summarize_product_order(
            order_detail("P2", status="PAYED", delivery={}, po={"placeOrderStatus": "NOT_YET"}))
        self.assertEqual(waiting["status"], "결제완료(발주 미확인)")
        self.assertFalse(waiting["is_inquiry_target"])
        self.assertEqual(waiting["shipping_due"], "2026-11-22 23:59")
        exchange = naver_commerce.summarize_product_order(order_detail(
            "P3", status="EXCHANGED", po={"claimType": "EXCHANGE", "claimStatus": "EXCHANGE_DONE"},
            exchange={"exchangeReason": "BROKEN_AND_BAD", "collectDeliveryCompany": "EPOST",
                      "collectTrackingNumber": "6000", "reDeliveryCompany": "CJGLS",
                      "reDeliveryTrackingNumber": "7000"}))
        self.assertEqual(exchange["claim"], "교환 · 교환 완료 · 사유: 상품 파손·불량")
        self.assertEqual(exchange["claim_tracking"], "수거 우체국택배 6000 / 재배송 CJ대한통운 7000")
        unknown = naver_commerce.summarize_product_order(order_detail(
            "P4", status="NEW_STATUS", po={"claimType": "CANCEL", "claimStatus": "CANCEL_DONE"},
            cancel={"cancelReason": "NEW_REASON"}, delivery={"deliveryCompany": "NEWCO", "trackingNumber": "1"}))
        self.assertEqual(unknown["status"], "NEW_STATUS")
        self.assertEqual(unknown["claim"], "취소 · 취소 완료 · 사유: NEW_REASON")
        self.assertEqual(unknown["company"], "NEWCO")
        visit = naver_commerce.summarize_product_order(order_detail(
            "P5", po={"expectedDeliveryMethod": "VISIT_RECEIPT"}, delivery={"deliveryMethod": "VISIT_RECEIPT"}))
        self.assertEqual(visit["delivery_method"], "방문수령")

    def test_split_ids(self):
        self.assertEqual(naver_commerce._split_ids("2026A"), ["2026A"])
        self.assertEqual(naver_commerce._split_ids("A, B,"), ["A", "B"])
        self.assertEqual(naver_commerce._split_ids(["A", 1]), ["A", "1"])
        self.assertEqual(naver_commerce._split_ids(None), [])
        row = naver_commerce.normalize_customer_inquiry(dict(CUSTOMER, productOrderIdList="P1"))
        self.assertEqual(row["product_order_ids"], ["P1"])

    def test_fetch_puts_inquiry_target_first(self):
        asked = []

        def details(token, ids):
            asked.append(list(ids))
            return [order_detail(i) for i in ids]

        with patch.object(naver_commerce, "fetch_product_order_ids_of_order", return_value=["P0", "P1"]), \
                patch.object(naver_commerce, "fetch_product_order_details", side_effect=details):
            rows = naver_commerce.fetch_inquiry_order_summary("t", "O1", ["P1"])
        self.assertEqual(asked, [["P0", "P1"]])
        self.assertEqual([r["product_order_id"] for r in rows], ["P1", "P0"])

    def test_fetch_falls_back_to_inquiry_ids(self):
        with patch.object(naver_commerce, "fetch_product_order_ids_of_order", side_effect=RuntimeError("HTTP 500")), \
                patch.object(naver_commerce, "fetch_product_order_details",
                             side_effect=lambda t, ids: [order_detail(i) for i in ids]):
            rows = naver_commerce.fetch_inquiry_order_summary("t", "O1", ["P1"])
            self.assertEqual([r["product_order_id"] for r in rows], ["P1"])
            with self.assertRaises(RuntimeError):
                naver_commerce.fetch_inquiry_order_summary("t", "O1", [])


class BadgeTextTests(unittest.TestCase):
    def test_badge_text(self):
        from ui.main.chrome import nav_badge_text
        self.assertEqual([nav_badge_text(v) for v in (None, 0, -1, "x", 1, 99, 100, "7")],
                         ["", "", "", "", "1", "99", "99+", "7"])


class OrderHtmlTests(unittest.TestCase):
    def test_states(self):
        from ui.main.inquiries import inquiry_detail_html, inquiry_order_html
        self.assertEqual(inquiry_order_html(None), "")
        self.assertIn("불러오는 중", inquiry_order_html({"loading": True}))
        self.assertIn("불러오지 못했습니다: HTTP 500 &lt;x&gt;",
                      inquiry_order_html({"ok": False, "error": "HTTP 500 <x>"}))
        self.assertIn("찾지 못했습니다", inquiry_order_html({"ok": True, "items": []}))
        row = naver_commerce.normalize_customer_inquiry(CUSTOMER)
        self.assertNotIn("주문·배송", inquiry_detail_html(row))

    def test_items(self):
        from ui.main.inquiries import inquiry_order_html
        target = naver_commerce.summarize_product_order(
            order_detail("P1", delivery=dict(order_detail("P1")["delivery"], isWrongTrackingNumber=True)), ["P1"])
        other = naver_commerce.summarize_product_order(order_detail(
            "P2", status="PAYED", delivery={}, po={"productName": "<쌀>", "productOption": ""}), ["P1"])
        text = inquiry_order_html({"ok": True, "items": [target, other]})
        self.assertIn("(주문 2026-11-20 10:00)", text)
        self.assertIn("<b>[문의 상품]</b> 현미 5kg / 5kg × 2", text)
        self.assertIn("배송: 우체국택배 6012345678901 (배송완료) · 발송 2026-11-20 15:00 · 도착 2026-11-21 12:30", text)
        self.assertIn("송장번호 오류", text)
        self.assertIn("&lt;쌀&gt; × 2", text)
        self.assertIn("아직 발송 전 · 발송기한 2026-11-22 23:59", text)
        self.assertEqual(text.count("수령인: 홍길동"), 1)  # 같은 주소면 한 번만
        self.assertIn("주소: (06000) 서울 강남구 테헤란로 1 101호", text)
        self.assertIn("배송메모: 문 앞", text)
        single = inquiry_order_html({"ok": True, "items": [target]})
        self.assertNotIn("[문의 상품]", single)  # 상품이 하나면 표시하지 않는다
        sent = naver_commerce.summarize_product_order(order_detail(
            "P3", status="DELIVERED", delivery={"sendDate": "2026-11-20T15:00:00.0+09:00"}))
        sent_text = inquiry_order_html({"ok": True, "items": [sent]})
        self.assertIn("발송 2026-11-20 15:00", sent_text)
        self.assertNotIn("아직 발송 전", sent_text)


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

    order_calls = []
    order_gate = threading.Event()
    order_gate.set()

    def fake_order_worker(order_id, product_order_ids=()):
        order_calls.append([order_id, list(product_order_ids)])
        order_gate.wait(10)
        if order_id == "FAIL":
            return {"ok": False, "order_id": order_id, "error": "HTTP 500"}
        item = naver_commerce.summarize_product_order(order_detail("P1"), product_order_ids)
        return {"ok": True, "order_id": order_id, "items": [item]}

    def wait_order(window):
        deadline = time.time() + 10
        while time.time() < deadline:
            app.processEvents()
            wanted = window._inquiry_order_wanted
            if window._inquiry_order_thread is None and (wanted is None or wanted[0] in window._inquiry_orders):
                app.processEvents()
                return
            time.sleep(0.01)
        raise AssertionError("주문 정보 조회가 끝나지 않음")

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
            patch.object(module, "run_naver_inquiry_list_worker", fake_worker), \
            patch.object(module, "run_naver_inquiry_order_worker", fake_order_worker):
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
        result["qna_order_calls"] = len(order_calls)  # 상품문의는 주문 정보를 부르지 않는다

        # 고객문의를 고르면 주문·배송 정보를 불러와 상세 아래에 붙이고, 다시 골라도 다시 부르지 않는다
        order_gate.clear()
        table.selectRow(0)
        app.processEvents()
        loading_text = window.textBrowser_inquiry_detail.toPlainText()
        order_gate.set()
        wait_order(window)
        result["order"] = {
            "loading": "주문 정보를 불러오는 중" in loading_text,
            "calls": list(order_calls),
            "text": window.textBrowser_inquiry_detail.toPlainText(),
        }
        table.selectRow(1)
        app.processEvents()
        table.selectRow(0)
        wait_order(window)
        result["order"]["calls_after_reselect"] = len(order_calls)
        table.selectRow(1)  # 아래 단계(필터 변경 시 선택 유지)는 상품문의를 고른 상태에서 본다
        app.processEvents()

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

        # 주문 정보를 불러오는 중에 다른 고객문의를 고르면 끝난 뒤 마지막으로 고른 것만 불러온다.
        # 실패한 주문은 다시 고르면 다시 부른다. 목록을 새로 불러오면 주문 정보도 다시 읽는다.
        many = [naver_commerce.normalize_customer_inquiry(dict(CUSTOMER, inquiryNo=900 + i, orderId=oid,
                                                                productOrderIdList=f"P{i}"))
                for i, oid in enumerate(("A", "B", "C", "FAIL"))]
        responses.append({"ok": True, "rows": many, "errors": [], "sample_keys": {}, "fetched_at": "12:10"})
        window.refresh_inquiries()
        wait_idle(window)
        del order_calls[:]
        order_gate.clear()
        table.selectRow(0)
        app.processEvents()
        table.selectRow(1)
        app.processEvents()
        table.selectRow(2)
        app.processEvents()
        order_gate.set()
        wait_order(window)
        queued = [c[0] for c in order_calls]
        table.selectRow(3)
        wait_order(window)
        fail_text = window.textBrowser_inquiry_detail.toPlainText()
        table.selectRow(2)
        wait_order(window)
        table.selectRow(3)
        wait_order(window)
        refetch_fail = [c[0] for c in order_calls].count("FAIL")
        responses.append({"ok": True, "rows": many, "errors": [], "sample_keys": {}, "fetched_at": "12:20"})
        window.refresh_inquiries()
        wait_idle(window)
        wait_order(window)
        result["order_queue"] = {
            "queued": queued,
            "fail_text": "주문 정보를 불러오지 못했습니다: HTTP 500" in fail_text,
            "fail_calls": refetch_fail,
            "after_refresh": [c[0] for c in order_calls][-1:],
            "selected_after_refresh": table.currentRow(),
        }

        # 측면 툴바 「문의」 배지: 주기 조회 결과로 숫자를 붙이고, 실패하면 떼고 툴팁에 알린다
        from PySide6.QtGui import QColor
        from shared import theme

        startup_names = list(type(window)._BACKGROUND_TIMER_NAMES)
        counts = []

        def fake_count():
            return counts.pop(0)

        def wait_badge():
            deadline = time.time() + 10
            while time.time() < deadline:
                app.processEvents()
                if window._inquiry_badge_thread is None:
                    app.processEvents()
                    return
                time.sleep(0.01)
            raise AssertionError("배지 조회가 끝나지 않음")

        def red_corner():
            image = window.nav_rail_actions["tab_inquiries"].icon().pixmap(48, 48).toImage()
            pixel = image.pixelColor(image.width() - 6, 6)
            danger = QColor(theme.DANGER)
            return pixel.alpha() > 200 and abs(pixel.red() - danger.red()) < 30 and pixel.green() < 100

        badge = {"timer": [window._inquiry_badge_timer.interval(), window._inquiry_badge_timer.isActive(),
                           "_inquiry_badge_timer" in startup_names],
                 "initial": list(window.nav_rail_badge("tab_inquiries")), "initial_red": red_corner()}
        with patch.object(module, "run_naver_inquiry_count_worker", fake_count):
            counts.append({"ok": True, "count": 3, "partial": False, "checked_at": "13:00"})
            window._poll_inquiry_badge()
            wait_badge()
            badge["three"] = list(window.nav_rail_badge("tab_inquiries"))
            badge["three_tip"] = window.nav_rail_actions["tab_inquiries"].toolTip()
            badge["three_red"] = red_corner()
            counts.append({"ok": True, "count": 120, "partial": True, "checked_at": "13:05"})
            window._poll_inquiry_badge()
            wait_badge()
            badge["many"] = list(window.nav_rail_badge("tab_inquiries"))
            counts.append({"ok": False, "error": "HTTP 500", "checked_at": "13:10"})
            window._poll_inquiry_badge()
            wait_badge()
            badge["failed"] = list(window.nav_rail_badge("tab_inquiries"))
            badge["failed_red"] = red_corner()
            # 탭 목록을 배지와 같은 조건(미답변만·30일)으로 빠짐없이 불러오면 배지도 맞춘다
            window.comboBox_inquiry_period.blockSignals(True)
            window.comboBox_inquiry_period.setCurrentIndex(1)
            window.comboBox_inquiry_period.blockSignals(False)
            window.checkBox_inquiry_unanswered.blockSignals(True)
            window.checkBox_inquiry_unanswered.setChecked(True)
            window.checkBox_inquiry_unanswered.blockSignals(False)
            two = many[:2]
            responses.append({"ok": True, "rows": two, "errors": [], "sample_keys": {}, "fetched_at": "13:20",
                              "days": 30, "unanswered_only": True})
            window.refresh_inquiries()
            wait_idle(window)
            badge["from_list"] = list(window.nav_rail_badge("tab_inquiries"))
            responses.append({"ok": True, "rows": many, "errors": [], "sample_keys": {}, "fetched_at": "13:25",
                              "days": 7, "unanswered_only": True})
            window.refresh_inquiries()
            wait_idle(window)
            badge["from_short_period"] = list(window.nav_rail_badge("tab_inquiries"))
            responses.append({"ok": True, "rows": many, "errors": ["상품문의 조회 실패: HTTP 500"],
                              "sample_keys": {}, "fetched_at": "13:30", "days": 30, "unanswered_only": True})
            window.refresh_inquiries()
            wait_idle(window)
            badge["from_partial_list"] = list(window.nav_rail_badge("tab_inquiries"))
            window._close_ready = True  # 종료 중에는 세지 않는다
            window._poll_inquiry_badge()
            badge["skipped_when_closing"] = window._inquiry_badge_thread is None
            window._close_ready = False
            window._stop_background_timers()
            badge["stopped"] = not window._inquiry_badge_timer.isActive()
        result["badge"] = badge

    # 작업자: 최근 토큰을 다시 쓰고, 거절되면 한 번만 새로 받아 다시 조회한다
    tokens = []
    summary_calls = []

    def fake_summary(token, order_id, product_order_ids=()):
        summary_calls.append(token)
        if token == "old":
            raise RuntimeError("HTTP 401: 토큰 만료")
        return [{"product_order_id": "P1"}]

    def fake_new_token():
        tokens.append("new")
        module._remember_inquiry_token("new")
        return "new"

    with patch.object(naver_commerce, "fetch_inquiry_order_summary", fake_summary), \
            patch.object(module, "_naver_inquiry_token_from_config", fake_new_token):
        module._remember_inquiry_token("old")
        first = module.run_naver_inquiry_order_worker("O1", ["P1"])
        second = module.run_naver_inquiry_order_worker("O1", ["P1"])
        module._forget_inquiry_token()
        third = module.run_naver_inquiry_order_worker("O1", ["P1"])
    result["token"] = {"oks": [first["ok"], second["ok"], third["ok"]],
                       "summary_calls": summary_calls, "new_tokens": len(tokens)}

    # 배지 작업자: 토큰을 다시 쓰고, 두 종류 모두 거절되면 새 토큰으로 한 번 더 센다
    list_calls = []

    def fake_list(token, from_dt, to_dt, unanswered_only=True):
        list_calls.append([token, unanswered_only, (to_dt - from_dt).days])
        if token == "old":
            return {"rows": [], "errors": ["상품문의 조회 실패: HTTP 401", "고객문의 조회 실패: HTTP 401"],
                    "sample_keys": {}}
        if token == "half":
            return {"rows": [{"id": "1"}], "errors": ["상품문의 조회 실패: HTTP 500"], "sample_keys": {}}
        return {"rows": [{"id": "1"}, {"id": "2"}], "errors": [], "sample_keys": {}}

    tokens.clear()
    with patch.object(naver_commerce, "fetch_inquiry_list", fake_list), \
            patch.object(module, "_naver_inquiry_token_from_config", fake_new_token):
        module._remember_inquiry_token("old")
        rejected = module.run_naver_inquiry_count_worker()
        reused = module.run_naver_inquiry_count_worker()
        module._remember_inquiry_token("half")
        partial = module.run_naver_inquiry_count_worker()
    result["count_worker"] = {
        "rejected": [rejected["ok"], rejected.get("count")],
        "reused": [reused["ok"], reused.get("count"), reused.get("partial")],
        "partial": [partial["ok"], partial.get("count"), partial.get("partial")],
        "calls": [c[0] for c in list_calls],
        "unanswered_days": sorted({(c[1], c[2]) for c in list_calls}),
        "new_tokens": len(tokens),
    }
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

    def test_customer_inquiry_shows_order(self):
        result = scenario()
        self.assertEqual(result["qna_order_calls"], 0)
        order = result["order"]
        self.assertTrue(order["loading"])
        self.assertEqual(order["calls"], [["2026112100001", []]])
        self.assertIn("주문·배송", order["text"])
        self.assertIn("우체국택배 6012345678901 (배송완료)", order["text"])
        self.assertIn("수령인: 홍길동", order["text"])
        self.assertEqual(order["calls_after_reselect"], 1)

    def test_order_queue_retry_and_refresh(self):
        q = scenario()["order_queue"]
        self.assertEqual(q["queued"], ["A", "C"])  # 불러오는 중에 고른 B는 건너뛴다
        self.assertTrue(q["fail_text"])
        self.assertEqual(q["fail_calls"], 2)  # 실패한 주문은 다시 고르면 다시 부른다
        self.assertEqual(q["after_refresh"], ["FAIL"])  # 새로 불러온 뒤 선택 유지 → 다시 읽음
        self.assertEqual(q["selected_after_refresh"], 3)

    def test_worker_reuses_token_and_retries_once(self):
        token = scenario()["token"]
        self.assertEqual(token["oks"], [True, True, True])
        self.assertEqual(token["summary_calls"], ["old", "new", "new", "new"])
        self.assertEqual(token["new_tokens"], 2)

    def test_badge(self):
        b = scenario()["badge"]
        self.assertEqual(b["timer"], [300000, True, True])
        self.assertEqual(b["initial"], ["", ""])
        self.assertFalse(b["initial_red"])
        self.assertEqual(b["three"], ["3", "미답변 3건 · 13:00 확인"])
        self.assertEqual(b["three_tip"], "네이버 문의 (Ctrl+7) · 미답변 3건 · 13:00 확인")
        self.assertTrue(b["three_red"])
        self.assertEqual(b["many"], ["99+", "미답변 120건 · 13:05 확인 (일부 실패)"])
        self.assertEqual(b["failed"], ["", "미답변 수 확인 실패 · 13:10"])
        self.assertFalse(b["failed_red"])
        self.assertEqual(b["from_list"], ["2", "미답변 2건 · 13:20 확인"])
        self.assertEqual(b["from_short_period"], ["2", "미답변 2건 · 13:20 확인"])  # 7일은 배지와 기간이 달라 그대로
        self.assertEqual(b["from_partial_list"], ["2", "미답변 2건 · 13:20 확인"])  # 일부 실패면 그대로
        self.assertTrue(b["skipped_when_closing"])
        self.assertTrue(b["stopped"])

    def test_count_worker(self):
        c = scenario()["count_worker"]
        self.assertEqual(c["rejected"], [True, 2])
        self.assertEqual(c["reused"], [True, 2, False])
        self.assertEqual(c["partial"], [True, 1, True])
        self.assertEqual(c["calls"], ["old", "new", "new", "half"])
        self.assertEqual(c["unanswered_days"], [[True, 30]])
        self.assertEqual(c["new_tokens"], 1)

    def test_failure_is_shown(self):
        self.assertEqual(scenario()["error_status"],
                         "✗ 불러오지 못함: 네이버 client_id/secret 이 설정되지 않았습니다.")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
