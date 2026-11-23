"""「문의」 탭의 쿠팡 문의 조회(읽기 전용)를 검증한다.

- 쿠팡 API: 콜센터문의 페이지(30건/쪽)를 totalPages 로 끝까지 읽음, 전체 조회는 NONE 으로 보냄
- 응답 → 탭 행 변환: 온라인문의(답변 = commentDtoList, 상품명은 등록 상품 상세에서), 콜센터문의(상태 조합 →
  답변 여부, 상담 기록, 연락처 제외, 접수 분류 정리)
- 상품명 찾기: 기억한 이름은 다시 조회하지 않음, 한 번에 찾는 수 제한, 실패는 이름 없이
- 목록: 한 종류가 실패해도 다른 종류는 보여 줌, 최근 등록 순
- 작업자: 키가 있는 스토어만 조회, 한 스토어 실패는 일부 실패, 네이버를 빠짐없이 읽었는지 표시
- 화면: 유형 칸 「쿠팡 CS문의」, 쿠팡 문의는 답변 칸·주문 조회 없음, 상담 기록·WING 안내, 키 없음 표시
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

from integrations import coupang_commerce  # noqa: E402
from integrations import naver_commerce  # noqa: E402

ONLINE = {
    "inquiryId": 3001, "productId": 9001, "sellerProductId": 7001, "sellerItemId": 1, "vendorItemId": 2,
    "content": "사이즈 문의", "inquiryAt": "2026-11-22T10:15:30+09:00", "orderIds": [111, 222],
    "commentDtoList": [
        {"inquiryCommentId": 1, "inquiryId": 3001, "content": "M 사이즈입니다.", "inquiryCommentAt": "2026-11-22T11:00:00+09:00"},
        {"inquiryCommentId": 2, "inquiryId": 3001, "content": "추가 안내", "inquiryCommentAt": "2026-11-22T12:30:00+09:00"},
    ],
}
CALLCENTER_OPEN = {
    "inquiryId": 4001, "csPartnerCounselingStatus": "requestAnswer", "inquiryStatus": "progress",
    "vendorItemId": [5], "itemName": "현미 5kg, 1개", "content": "배송이 안 와요",
    "answeredAt": "2026-11-23T09:10:00+09:00",
    "replies": [{"receptionistName": "상담원", "receptionist": None, "replyAt": "2026-11-23T09:10:00+09:00",
                 "content": "판매자님 확인 부탁드립니다", "answerId": 1, "parentAnswerId": None,
                 "partnerTransferStatus": "requestAnswer", "partnerTransferCompleteReason": "NONE",
                 "answerType": "csAgent", "needAnswer": True}],
    "inquiryAt": "2026-11-23T09:00:00+09:00", "buyerEmail": "", "buyerPhone": "010-****-1234",
    "orderId": 3100000000001, "orderDate": "2026-11-20T08:00:00+09:00",
    "receiptCategory": "배송>배송>배송",
}


def answered_callcenter():
    return dict(
        CALLCENTER_OPEN, inquiryId=4002, csPartnerCounselingStatus="answered", inquiryStatus="complete",
        inquiryAt="2026-11-21T09:00:00+09:00", receiptCategory=">>",
        replies=CALLCENTER_OPEN["replies"] + [
            {"answerType": "vendor", "needAnswer": False, "replyAt": "2026-11-23T10:00:00+09:00",
             "content": "오늘 출고했습니다."}])


class PagingTests(unittest.TestCase):
    def test_reads_every_page_by_total_pages(self):
        # 콜센터문의는 pageSize 50을 보내도 30건씩 준다 → 받은 건수가 아니라 totalPages 로 끝을 본다
        calls = []

        def fake_get(path, params, access_key, secret_key):
            calls.append(dict(params))
            count = 30 if params["pageNum"] == 1 else 5
            return {"data": {"content": [{"inquiryId": i} for i in range(count)],
                             "pagination": {"currentPage": params["pageNum"], "totalPages": 2,
                                            "totalElements": 35, "countPerPage": 30}}}

        with patch.object(coupang_commerce, "_get", fake_get):
            items = coupang_commerce.fetch_callcenter_inquiries(
                "V", "a", "s", datetime(2026, 11, 23), datetime(2026, 11, 23), answered=None)
        self.assertEqual(len(items), 35)
        self.assertEqual([c["pageNum"] for c in calls], [1, 2])
        self.assertEqual({c["partnerCounselingStatus"] for c in calls}, {"NONE"})  # 빼면 0건이 온다

    def test_falls_back_to_count_without_pagination(self):
        pages = [[{"inquiryId": i} for i in range(coupang_commerce.PAGE_SIZE)], [{"inquiryId": 99}]]
        calls = []

        def fake_get(path, params, access_key, secret_key):
            calls.append(params["pageNum"])
            return {"data": pages[params["pageNum"] - 1]}

        with patch.object(coupang_commerce, "_get", fake_get):
            items = coupang_commerce.fetch_online_inquiries(
                "V", "a", "s", datetime(2026, 11, 23), datetime(2026, 11, 23), answered=False)
        self.assertEqual(len(items), coupang_commerce.PAGE_SIZE + 1)
        self.assertEqual(calls, [1, 2])

    def test_status_parameter(self):
        seen = []

        def fake_get(path, params, access_key, secret_key):
            seen.append((path.rsplit("/", 1)[-1], params.get("partnerCounselingStatus"), params.get("answeredType")))
            return {"data": {"content": [], "pagination": {"totalPages": 0}}}

        day = datetime(2026, 11, 23)
        with patch.object(coupang_commerce, "_get", fake_get):
            coupang_commerce.fetch_callcenter_inquiries("V", "a", "s", day, day, answered=False)
            coupang_commerce.fetch_callcenter_inquiries("V", "a", "s", day, day, answered=True)
            coupang_commerce.fetch_online_inquiries("V", "a", "s", day, day, answered=False)
            coupang_commerce.fetch_online_inquiries("V", "a", "s", day, day, answered=None)
        self.assertEqual(seen, [("callCenterInquiries", "NO_ANSWER", None), ("callCenterInquiries", "ANSWER", None),
                                ("onlineInquiries", None, "NOANSWER"), ("onlineInquiries", None, "ALL")])


class NormalizeTests(unittest.TestCase):
    def test_online_inquiry(self):
        row = coupang_commerce.normalize_online_inquiry(ONLINE, {"7001": "유기농 쌀 10kg"})
        self.assertEqual(row["kind"], "쿠팡 상품문의")
        self.assertEqual(row["store"], "쿠팡")
        self.assertEqual(row["id"], "3001")
        self.assertEqual(row["created"], "2026-11-22 10:15")
        self.assertEqual(row["product"], "유기농 쌀 10kg")
        self.assertEqual(row["product_no"], "7001")
        self.assertEqual(row["order_id"], "111, 222")
        self.assertIs(row["answered"], True)
        self.assertEqual(row["answer"], "M 사이즈입니다.\n\n추가 안내")
        self.assertEqual(row["answered_at"], "2026-11-22 12:30")

    def test_online_inquiry_without_answer_or_name(self):
        row = coupang_commerce.normalize_online_inquiry(dict(ONLINE, commentDtoList=[], orderIds=None))
        self.assertIs(row["answered"], False)
        self.assertEqual((row["product"], row["order_id"], row["answer"], row["answered_at"]), ("", "", "", ""))

    def test_callcenter_open(self):
        row = coupang_commerce.normalize_callcenter_inquiry(CALLCENTER_OPEN)
        self.assertEqual(row["kind"], "쿠팡 CS문의")
        self.assertEqual(row["id"], "4001")
        self.assertIs(row["answered"], False)
        self.assertEqual(row["product"], "현미 5kg, 1개")
        self.assertEqual(row["order_id"], "3100000000001")
        self.assertEqual(row["category"], "배송")
        self.assertEqual(row["answered_at"], "")  # 미답변이면 답변 시각을 보이지 않는다
        self.assertEqual(row["replies"], [{"who": "쿠팡 상담원 · 판매자 답변 요청", "vendor": False,
                                           "at": "2026-11-23 09:10", "content": "판매자님 확인 부탁드립니다"}])
        self.assertNotIn("010-****-1234", json.dumps(row, ensure_ascii=False))  # 연락처는 넣지 않는다

    def test_callcenter_answered_and_closed(self):
        row = coupang_commerce.normalize_callcenter_inquiry(answered_callcenter())
        self.assertIs(row["answered"], True)
        self.assertEqual(row["answer"], "오늘 출고했습니다.")
        self.assertEqual(row["answered_at"], "2026-11-23 09:10")
        self.assertEqual(row["category"], "")
        self.assertEqual([r["who"] for r in row["replies"]], ["쿠팡 상담원 · 판매자 답변 요청", "판매자"])
        closed = coupang_commerce.normalize_callcenter_inquiry(dict(CALLCENTER_OPEN, inquiryStatus="complete"))
        self.assertIsNone(closed["answered"])  # 쿠팡 TRANSFER: 판매자 답변 없이 종료
        self.assertEqual(closed["status_label"], "판매자 답변 없이 종료")


class ProductNameTests(unittest.TestCase):
    def setUp(self):
        coupang_commerce._product_names.clear()

    def tearDown(self):
        coupang_commerce._product_names.clear()

    def test_remembers_limits_and_skips_failures(self):
        calls = []

        def fake_product(pid, access_key, secret_key):
            calls.append(pid)
            if pid == "bad":
                raise RuntimeError("HTTP 500")
            return {"sellerProductName": f"상품{pid}"}

        with patch.object(coupang_commerce, "fetch_seller_product", fake_product):
            first = coupang_commerce.product_names_for(["1", "2", "bad", "1", None], "a", "s")
            second = coupang_commerce.product_names_for(["1", "2", "3", "4"], "a", "s", limit=1)
        self.assertEqual(first, {"1": "상품1", "2": "상품2"})
        self.assertEqual(sorted(calls[:3]), ["1", "2", "bad"])
        self.assertEqual(calls[3:], ["3"])  # 기억한 1·2는 다시 부르지 않고, 새로 찾는 건 limit 개까지
        self.assertEqual(second, {"1": "상품1", "2": "상품2", "3": "상품3"})


class FetchListTests(unittest.TestCase):
    def setUp(self):
        coupang_commerce._product_names.clear()

    def call(self, online, callcenter, unanswered_only=True):
        answered_args = []

        def fake(result):
            def fn(vendor_id, access_key, secret_key, from_dt, to_dt, answered=None):
                answered_args.append(answered)
                if isinstance(result, Exception):
                    raise result
                return result
            return fn

        with patch.object(coupang_commerce, "fetch_online_inquiries", fake(online)), \
                patch.object(coupang_commerce, "fetch_callcenter_inquiries", fake(callcenter)), \
                patch.object(coupang_commerce, "fetch_seller_product",
                             lambda pid, a, s: {"sellerProductName": "쌀"}):
            res = coupang_commerce.fetch_inquiry_list("V", "a", "s", datetime(2026, 11, 1), datetime(2026, 11, 23),
                                                      unanswered_only=unanswered_only)
        return res, answered_args

    def test_merges_sorted_newest_first(self):
        closed = dict(CALLCENTER_OPEN, inquiryId=4003, inquiryStatus="complete", inquiryAt="2026-11-20T00:00:00+09:00")
        res, answered = self.call([ONLINE], [CALLCENTER_OPEN, closed, {"inquiryId": None}])
        self.assertEqual(answered, [False, False])
        self.assertEqual([r["id"] for r in res["rows"]], ["4001", "3001", "4003"])
        self.assertEqual(res["rows"][1]["product"], "쌀")
        self.assertIs(res["rows"][2]["answered"], False)  # 미답변만 조회에서 알 수 없음은 미답변으로
        self.assertEqual(res["errors"], [])
        self.assertIn("commentDtoList", res["sample_keys"]["쿠팡 상품문의"])

    def test_one_kind_failure_keeps_other(self):
        res, answered = self.call(RuntimeError("HTTP 403: 허용되지 않은 IP"), [CALLCENTER_OPEN], unanswered_only=False)
        self.assertEqual(answered, [None, None])
        self.assertEqual([r["id"] for r in res["rows"]], ["4001"])
        self.assertEqual(res["errors"], ["쿠팡 상품문의 조회 실패: HTTP 403: 허용되지 않은 IP"])


class DetailHtmlTests(unittest.TestCase):
    def test_coupang_rows(self):
        from ui.main.inquiries import inquiry_answerable, inquiry_detail_html, inquiry_kind_label

        open_row = coupang_commerce.normalize_callcenter_inquiry(CALLCENTER_OPEN)
        self.assertEqual(inquiry_kind_label(open_row), "쿠팡 CS문의")
        self.assertFalse(inquiry_answerable(open_row))  # 쿠팡 답변은 WING에서
        naver = naver_commerce.normalize_product_qna({"questionId": 1, "question": "q", "answered": False})
        self.assertEqual(inquiry_kind_label(naver), "네이버 상품문의")
        self.assertTrue(inquiry_answerable(naver))
        text = inquiry_detail_html(open_row)
        self.assertIn("&nbsp;쿠팡 CS문의&nbsp;", text)
        self.assertIn("상담 기록", text)
        self.assertIn("쿠팡 상담원 · 판매자 답변 요청", text)
        self.assertIn("아직 판매자가 답변하지 않았습니다.", text)
        self.assertIn("쿠팡 WING", text)
        self.assertIn("분류 배송", text)
        done = inquiry_detail_html(coupang_commerce.normalize_callcenter_inquiry(answered_callcenter()))
        self.assertIn("오늘 출고했습니다.", done)
        self.assertNotIn("쿠팡 WING", done)
        closed = inquiry_detail_html(coupang_commerce.normalize_callcenter_inquiry(
            dict(CALLCENTER_OPEN, inquiryStatus="complete")))
        self.assertIn("&nbsp;판매자 답변 없이 종료&nbsp;", closed)


def run_scenario():
    import importlib.util
    import time

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_coupang_inquiries", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])
    result = {}

    # 작업자: 키가 있는 스토어만, 한쪽 실패는 일부 실패, 네이버를 빠짐없이 읽었는지
    config = {}
    calls = {"naver": 0, "coupang": 0}

    def fake_naver_list(token, from_dt, to_dt, unanswered_only=True):
        calls["naver"] += 1
        return {"rows": [{"kind": "상품문의", "id": "1", "created": "2026-11-22 10:00"}],
                "errors": config.get("naver_errors", []), "sample_keys": {"상품문의": ["questionId"]}}

    def fake_coupang_list(vendor_id, access_key, secret_key, from_dt, to_dt, unanswered_only=True):
        calls["coupang"] += 1
        if config.get("coupang_raises"):
            raise RuntimeError("HTTP 403: 허용되지 않은 IP")
        return {"rows": [{"kind": "쿠팡 CS문의", "store": "쿠팡", "id": "9", "created": "2026-11-23 09:00"}],
                "errors": [], "sample_keys": {"쿠팡 CS문의": ["inquiryId"]}}

    def run_worker(keys, **options):
        config.clear()
        config.update(options)
        calls.update(naver=0, coupang=0)
        with patch.object(oauth, "get_authorized_gspread_client", lambda *a, **k: object()), \
                patch.object(module, "_standalone_open_config_ws", lambda gc: None), \
                patch.object(module, "_read_config_values_map", lambda ws: dict(keys)), \
                patch.object(naver_commerce, "get_access_token", lambda cid, secret: "tok"), \
                patch.object(naver_commerce, "fetch_inquiry_list", fake_naver_list), \
                patch.object(coupang_commerce, "fetch_inquiry_list", fake_coupang_list):
            res = module.run_inquiry_list_worker(0, True)
        return {"ok": res["ok"], "ids": [r["id"] for r in res.get("rows", [])], "errors": res.get("errors"),
                "error": res.get("error"), "naver_complete": res.get("naver_complete"),
                "missing": res.get("missing_keys"), "sample": sorted(res.get("sample_keys", {})),
                "calls": dict(calls)}

    naver_keys = {module.CONFIG_KEY_NAVER_CLIENT_ID: "id", module.CONFIG_KEY_NAVER_CLIENT_SECRET: "sec"}
    coupang_keys = {module.CONFIG_KEY_COUPANG_VENDOR_ID: "V", module.CONFIG_KEY_COUPANG_ACCESS_KEY: "a",
                    module.CONFIG_KEY_COUPANG_SECRET_KEY: "s"}
    result["worker"] = {
        "both": run_worker({**naver_keys, **coupang_keys}),
        "naver_only": run_worker(naver_keys),
        "coupang_only": run_worker(coupang_keys),
        "none": run_worker({}),
        "coupang_fails": run_worker({**naver_keys, **coupang_keys}, coupang_raises=True),
        "naver_partial": run_worker({**naver_keys, **coupang_keys}, naver_errors=["상품문의 조회 실패: HTTP 500"]),
    }

    # 화면: 쿠팡 문의는 답변 칸·주문 조회 없이 상담 기록을 보인다
    rows = [coupang_commerce.normalize_callcenter_inquiry(CALLCENTER_OPEN),
            coupang_commerce.normalize_online_inquiry(dict(ONLINE, commentDtoList=[])),
            naver_commerce.normalize_customer_inquiry({
                "inquiryNo": 202, "inquiryRegistrationDateTime": "2026-11-21T09:05:00.0+09:00", "orderId": "O1",
                "title": "주소 변경", "inquiryContent": "c", "answered": False})]
    order_calls = []

    def fake_list(days, unanswered_only=True):
        return {"ok": True, "rows": rows, "errors": [], "sample_keys": {}, "fetched_at": "10:00",
                "days": days, "unanswered_only": unanswered_only, "naver_complete": True, "missing_keys": ["네이버"]}

    def fake_order(order_id, product_order_ids=()):
        order_calls.append(order_id)
        return {"ok": True, "order_id": order_id, "items": []}

    def wait_idle(window):
        deadline = time.time() + 10
        while time.time() < deadline:
            app.processEvents()
            if (window._inquiry_thread is None and not window._inquiry_refresh_pending
                    and window._inquiry_order_thread is None):
                app.processEvents()
                return
            time.sleep(0.01)
        raise AssertionError("조회가 끝나지 않음")

    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(module, "run_inquiry_list_worker", fake_list), \
            patch.object(module, "run_naver_inquiry_order_worker", fake_order), \
            patch.object(module, "run_naver_qna_templates_worker", lambda: {"ok": True, "templates": []}):
        window = module.MainWindow()
        window.show()
        app.processEvents()
        window.nav_rail_actions["tab_inquiries"].trigger()
        wait_idle(window)
        table = window.tableWidget_inquiries
        card = window.card_inquiry_answer
        ui = {
            "kinds": [table.item(r, 0).text() for r in range(table.rowCount())],
            "products": [table.item(r, 3).text() for r in range(table.rowCount())],
            "status": window.label_inquiry_status.text(),
            "tab_text": window.ui.tabWidget.tabText(window.ui.tabWidget.indexOf(window.ui.tab_inquiries)),
        }
        table.selectRow(0)
        wait_idle(window)
        ui["cs_answer_visible"] = card.isVisibleTo(card.parentWidget())
        ui["cs_detail"] = window.textBrowser_inquiry_detail.toPlainText()
        ui["order_calls_after_coupang"] = list(order_calls)
        table.selectRow(2)
        wait_idle(window)
        ui["naver_answer_visible"] = card.isVisibleTo(card.parentWidget())
        ui["order_calls_after_naver"] = list(order_calls)
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


class ListWorkerTests(unittest.TestCase):
    def test_both_stores(self):
        w = scenario()["worker"]["both"]
        self.assertTrue(w["ok"])
        self.assertEqual(w["ids"], ["9", "1"])  # 최근 등록 순으로 합친다
        self.assertEqual(w["calls"], {"naver": 1, "coupang": 1})
        self.assertTrue(w["naver_complete"])
        self.assertEqual(w["missing"], [])
        self.assertEqual(w["sample"], ["네이버 상품문의", "쿠팡 CS문의"])

    def test_only_stores_with_keys(self):
        w = scenario()["worker"]
        self.assertEqual(w["naver_only"]["calls"], {"naver": 1, "coupang": 0})
        self.assertEqual(w["naver_only"]["missing"], ["쿠팡"])
        self.assertEqual(w["coupang_only"]["calls"], {"naver": 0, "coupang": 1})
        self.assertEqual(w["coupang_only"]["missing"], ["네이버"])
        self.assertFalse(w["coupang_only"]["naver_complete"])
        self.assertFalse(w["none"]["ok"])
        self.assertIn("API 키가 설정되지 않았습니다", w["none"]["error"])

    def test_partial_failures(self):
        w = scenario()["worker"]
        self.assertTrue(w["coupang_fails"]["ok"])
        self.assertEqual(w["coupang_fails"]["ids"], ["1"])
        self.assertEqual(w["coupang_fails"]["errors"], ["쿠팡 조회 실패: HTTP 403: 허용되지 않은 IP"])
        self.assertTrue(w["coupang_fails"]["naver_complete"])  # 쿠팡 실패는 배지 맞춤과 상관없다
        self.assertEqual(w["naver_partial"]["errors"], ["네이버 상품문의 조회 실패: HTTP 500"])
        self.assertFalse(w["naver_partial"]["naver_complete"])


class CoupangTabTests(unittest.TestCase):
    def test_list_and_detail(self):
        ui = scenario()["ui"]
        self.assertEqual(ui["tab_text"], "문의")
        self.assertEqual(ui["kinds"], ["쿠팡 CS문의", "쿠팡 상품문의", "네이버 고객문의"])
        self.assertEqual(ui["products"], ["현미 5kg, 1개", "상품 7001", ""])  # 이름을 못 찾으면 상품 번호
        self.assertIn("네이버 키 없음", ui["status"])
        self.assertFalse(ui["cs_answer_visible"])  # 쿠팡 답변은 WING에서
        self.assertIn("판매자 답변 요청", ui["cs_detail"])
        self.assertIn("쿠팡 WING", ui["cs_detail"])
        self.assertEqual(ui["order_calls_after_coupang"], [])  # 쿠팡 주문번호로 네이버 주문을 부르지 않는다
        self.assertTrue(ui["naver_answer_visible"])
        self.assertEqual(ui["order_calls_after_naver"], ["O1"])


if __name__ == "__main__":
    if "--scenario" in sys.argv:
        run_scenario()
    else:
        unittest.main()
