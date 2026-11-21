"""「네이버 발주」 탭(발송 전 주문 조회, 읽기 전용)과 네이버 호출 한도 처리를 검증한다.

- group_pending_orders: 발송 전(PAYED)만, 주문번호별로 묶기, 하나라도 미확인이면 신규주문, 클레임,
  합계, 발송기한 빠른 순, 엑셀 경로용 행(rows)
- 429면 기다렸다 다시(최대 3번), 하루 단위 창 사이 간격
- 상세 HTML(배지·배송메모·이스케이프), 작업자(최근 3일), 화면(첫 진입 자동 조회, 필터, 선택 유지, 실패)
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


def detail(order_id, poid, status="PAYED", place="OK", due="2026-11-24T23:59:59.0+09:00",
           paid="2026-11-21T10:00:00.0+09:00", name="상품", amount=10000, qty=1, memo="", claim=None):
    po = {"productOrderId": poid, "productOrderStatus": status, "placeOrderStatus": place,
          "productName": name, "productOption": "옵션", "quantity": qty, "totalPaymentAmount": amount,
          "shippingDueDate": due, "shippingMemo": memo, "expectedDeliveryMethod": "DELIVERY",
          "productId": "123",
          "shippingAddress": {"name": f"수령{order_id}", "tel1": "010-0000-0000", "zipCode": "06000",
                              "baseAddress": "서울", "detailedAddress": "1호"}}
    if claim:
        po.update(claimType=claim[0], claimStatus=claim[1])
    return {"order": {"orderId": order_id, "paymentDate": paid, "ordererName": "구매자"}, "productOrder": po}


SAMPLE = [
    detail("A", "A1", place="OK", due="2026-11-25T23:59:59.0+09:00", name="쌀", amount=20000, qty=2),
    detail("A", "A2", place="NOT_YET", due="2026-11-24T23:59:59.0+09:00", name="현미", amount=5000, memo="문 앞"),
    detail("B", "B1", place="OK", due="2026-11-23T23:59:59.0+09:00", name="<보리>",
           claim=("CANCEL", "CANCEL_REQUEST")),
    detail("C", "C1", status="DELIVERING"),  # 발송 끝난 건은 빼고
    detail("D", "D1", status="PURCHASE_DECIDED"),
]


class GroupTests(unittest.TestCase):
    def test_grouping(self):
        orders = naver_commerce.group_pending_orders(SAMPLE)
        self.assertEqual([o["order_id"] for o in orders], ["B", "A"])  # 발송기한 빠른 순
        b, a = orders
        self.assertEqual(a["stage"], "신규주문")  # 하나라도 미확인이면 신규주문
        self.assertEqual(b["stage"], "발주확인")
        self.assertEqual([i["stage"] for i in a["items"]], ["발주확인", "신규주문"])
        self.assertEqual(a["shipping_due"], "2026-11-24 23:59")  # 가장 이른 발송기한
        self.assertEqual((a["total_amount"], a["total_quantity"]), (25000, 3))
        self.assertEqual(a["memo"], "문 앞")
        self.assertEqual((a["receiver"], a["tel"], a["zip"], a["address"]), ("수령A", "010-0000-0000", "06000", "서울 1호"))
        self.assertEqual(a["delivery_method"], "택배,등기,소포")
        self.assertEqual(a["paid_at"], "2026-11-21 10:00")
        self.assertFalse(a["has_claim"])
        self.assertTrue(b["has_claim"])
        self.assertEqual(b["items"][0]["claim"], "취소 · 취소 요청")
        # 다음 단계에서 엑셀 처리 경로에 그대로 넘길 행
        self.assertEqual([r["_상품주문번호"] for r in a["rows"]], ["A1", "A2"])
        self.assertEqual(a["rows"][0]["주문번호"], "A")
        self.assertEqual(a["rows"][1]["배송메세지"], "문 앞")

    def test_skips_bad_items(self):
        self.assertEqual(naver_commerce.group_pending_orders([None, {}, {"productOrder": {"productOrderStatus": "PAYED"}}]), [])


class FakeResponse:
    def __init__(self, status, body=b"{}"):
        self.status_code = status
        self.content = body
        self.text = body.decode()

    def json(self):
        return json.loads(self.content.decode())


class RateLimitTests(unittest.TestCase):
    def test_retries_on_429(self):
        responses = [FakeResponse(429), FakeResponse(429), FakeResponse(200, b'{"ok": 1}')]
        sleeps = []
        with patch.object(naver_commerce.requests, "get", side_effect=lambda *a, **k: responses.pop(0)), \
                patch.object(naver_commerce, "_sleep", sleeps.append):
            self.assertEqual(naver_commerce._get_json("u", "t", {}), {"ok": 1})
        self.assertEqual(sleeps, [1.0, 2.0])

    def test_gives_up_after_three_waits(self):
        calls = []
        sleeps = []

        def get(*a, **k):
            calls.append(1)
            return FakeResponse(429, '{"code": "GW.RATE_LIMIT", "message": "잠시"}'.encode())

        with patch.object(naver_commerce.requests, "get", get), patch.object(naver_commerce, "_sleep", sleeps.append):
            with self.assertRaises(RuntimeError) as ctx:
                naver_commerce._get_json("u", "t", {})
        self.assertEqual(len(calls), 4)
        self.assertEqual(sleeps, [1.0, 2.0, 4.0])
        self.assertIn("HTTP 429", str(ctx.exception))

    def test_post_also_retries(self):
        responses = [FakeResponse(429), FakeResponse(200, b'{"data": []}')]
        with patch.object(naver_commerce.requests, "post", side_effect=lambda *a, **k: responses.pop(0)), \
                patch.object(naver_commerce, "_sleep", lambda s: None):
            self.assertEqual(naver_commerce._post_json("u", "t", {}), {"data": []})

    def test_pause_between_day_windows(self):
        events = []

        def get_json(url, token, params):
            events.append(("call", params["lastChangedFrom"][:10]))
            return {"data": {"lastChangeStatuses": [{"productOrderId": len(events)}]}}

        with patch.object(naver_commerce, "_get_json", get_json), \
                patch.object(naver_commerce, "_sleep", lambda s: events.append(("sleep", s))):
            ids = naver_commerce.fetch_changed_product_order_ids(
                "t", datetime(2026, 11, 19, 12), datetime(2026, 11, 22, 12))
        self.assertEqual(events, [("call", "2026-11-19"), ("sleep", 0.3), ("call", "2026-11-20"),
                                  ("sleep", 0.3), ("call", "2026-11-21")])
        self.assertEqual(len(ids), 3)


class DetailHtmlTests(unittest.TestCase):
    def test_detail(self):
        from shared import theme
        from ui.main.naver_orders import naver_order_detail_html, naver_order_product_summary, won
        b, a = naver_commerce.group_pending_orders(SAMPLE)
        text = naver_order_detail_html(a)
        self.assertIn(f'color:{theme.WARN}; background-color:{theme.WARN_SUBTLE}; font-weight:600;">&nbsp;신규주문', text)
        self.assertIn(">수령A</p>", text)
        self.assertIn(">주문번호</td><td>A</td>", text)
        self.assertIn("2026-11-24 23:59", text)
        self.assertIn("수량 3 · 합계 25,000원", text)
        self.assertIn("× 2 · 20,000원", text)
        self.assertIn("&nbsp;발주확인&nbsp;", text)  # 한 주문에 단계가 섞이면 상품별 배지
        self.assertIn("배송메모", text)
        self.assertIn(f'<td width="4" bgcolor="{theme.WARN}"></td><td bgcolor="{theme.WARN_SUBTLE}"', text)
        self.assertIn(">연락처</td><td>010-0000-0000</td>", text)
        b_text = naver_order_detail_html(b)
        self.assertIn("&lt;보리&gt;", b_text)
        self.assertIn(f'color:{theme.DANGER}; background-color:{theme.DANGER_SUBTLE}; font-weight:600;">&nbsp;취소 · 취소 요청', b_text)
        self.assertNotIn("배송메모", b_text)
        self.assertEqual(naver_order_product_summary(a), "쌀 외 1건")
        self.assertEqual(naver_order_product_summary(b), "<보리>")
        self.assertEqual(won(1234567), "1,234,567원")


def run_scenario():
    import importlib.util
    import threading
    import time

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_naver_orders", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])
    result = {}

    # 작업자: 최근 3일, 토큰 재사용
    seen = {}

    def changed(token, from_dt, to_dt):
        seen["token"] = token
        seen["days"] = round((to_dt - from_dt).total_seconds() / 86400, 3)
        return ["A1", "A2"]

    with patch.object(naver_commerce, "fetch_changed_product_order_ids", changed), \
            patch.object(naver_commerce, "fetch_product_order_details", lambda token, ids: SAMPLE[:2]):
        module._remember_inquiry_token("tok")
        res = module.run_naver_pending_orders_worker()
    result["worker"] = {"ok": res["ok"], "orders": [o["order_id"] for o in res["orders"]], "days": res["days"],
                        "seen": seen}

    responses = []
    calls = []
    gate = threading.Event()
    gate.set()

    def fake(days=3):
        calls.append(days)
        gate.wait(10)
        return responses.pop(0)

    def wait_idle(window):
        deadline = time.time() + 10
        while time.time() < deadline:
            app.processEvents()
            if window._naver_orders_thread is None:
                app.processEvents()
                return
            time.sleep(0.01)
        raise AssertionError("발주 조회가 끝나지 않음")

    orders = naver_commerce.group_pending_orders(SAMPLE)
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(module, "run_naver_pending_orders_worker", fake):
        window = module.MainWindow()
        window.show()
        app.processEvents()
        table = window.tableWidget_naver_orders
        ui = {"before": window.stackedWidget_naver_orders.currentWidget() is window.emptyState_naver_orders}
        responses.append({"ok": True, "orders": orders, "fetched_at": "09:30", "days": 3})
        window.nav_rail_actions["tab_naver_orders"].trigger()
        wait_idle(window)
        ui["first"] = {"calls": list(calls), "rows": table.rowCount(),
                       "cells": [[table.item(r, c).text() for c in range(table.columnCount())]
                                 for r in range(table.rowCount())],
                       "status": window.label_naver_order_status.text(),
                       "tooltip": window.label_naver_order_status.toolTip()}
        window.nav_rail_actions["tab"].trigger()
        window.nav_rail_actions["tab_naver_orders"].trigger()
        wait_idle(window)
        ui["reenter_calls"] = len(calls)
        table.selectRow(1)
        app.processEvents()
        ui["detail"] = window.textBrowser_naver_order_detail.toPlainText()
        # 필터는 다시 부르지 않고 걸러서 보인다. 고른 주문이 남아 있으면 선택 유지
        window.comboBox_naver_order_filter.setCurrentIndex(1)  # 신규주문
        app.processEvents()
        ui["filter_new"] = {"rows": table.rowCount(), "selected": table.currentRow(),
                            "detail_kept": "수령A" in window.textBrowser_naver_order_detail.toPlainText(),
                            "calls": len(calls)}
        window.comboBox_naver_order_filter.setCurrentIndex(2)  # 발주확인
        app.processEvents()
        ui["filter_confirmed"] = [table.rowCount(), table.item(0, 2).text(),
                                  window.textBrowser_naver_order_detail.toPlainText() == ""]
        window.comboBox_naver_order_filter.setCurrentIndex(0)
        app.processEvents()
        # 새로고침 중에는 버튼 잠김, 다시 눌러도 한 번만
        gate.clear()
        responses.append({"ok": True, "orders": [], "fetched_at": "09:40", "days": 3})
        window.refresh_naver_orders()
        window.refresh_naver_orders()
        app.processEvents()
        ui["loading"] = [window.pushButton_naver_order_refresh.isEnabled(), window.label_naver_order_status.text()]
        gate.set()
        wait_idle(window)
        ui["empty"] = [window.stackedWidget_naver_orders.currentWidget() is window.emptyState_naver_orders,
                       window.emptyState_naver_orders.label.text(), len(calls),
                       window.label_naver_order_status.text()]
        responses.append({"ok": False, "error": "HTTP 429: GW.RATE_LIMIT"})
        window.refresh_naver_orders()
        wait_idle(window)
        ui["error"] = window.label_naver_order_status.text()
    result["ui"] = ui
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


class NaverOrdersTabTests(unittest.TestCase):
    def test_worker(self):
        w = scenario()["worker"]
        self.assertTrue(w["ok"])
        self.assertEqual(w["orders"], ["A"])
        self.assertEqual(w["days"], 3)
        self.assertEqual(w["seen"], {"token": "tok", "days": 3.0})

    def test_first_entry_loads_once(self):
        ui = scenario()["ui"]
        self.assertTrue(ui["before"])
        first = ui["first"]
        self.assertEqual(first["calls"], [3])
        self.assertEqual(first["cells"], [["발주확인 · 클레임", "2026-11-21 10:00", "수령B", "<보리>", "10,000원"],
                                          ["신규주문", "2026-11-21 10:00", "수령A", "쌀 외 1건", "25,000원"]])
        self.assertEqual(first["status"], "! 신규 1 · 발주확인 1건 · 09:30 조회 (최근 3일) · 클레임 1건")
        self.assertIn("클레임", first["tooltip"])
        self.assertEqual(ui["reenter_calls"], 1)

    def test_detail_and_filter(self):
        ui = scenario()["ui"]
        self.assertIn("신규주문", ui["detail"])
        self.assertIn("문 앞", ui["detail"])
        self.assertEqual(ui["filter_new"], {"rows": 1, "selected": 0, "detail_kept": True, "calls": 1})
        self.assertEqual(ui["filter_confirmed"], [1, "수령B", True])

    def test_refresh_empty_and_error(self):
        ui = scenario()["ui"]
        self.assertEqual(ui["loading"], [False, "불러오는 중…"])
        empty = ui["empty"]
        self.assertTrue(empty[0])
        self.assertEqual(empty[1], "최근 3일 안에 발송 전 주문이 없습니다.")
        self.assertEqual(empty[2], 2)  # 조회 중 다시 눌러도 한 번만
        self.assertEqual(empty[3], "✓ 신규 0 · 발주확인 0건 · 09:40 조회 (최근 3일)")
        self.assertEqual(ui["error"], "✗ 불러오지 못함: HTTP 429: GW.RATE_LIMIT")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
