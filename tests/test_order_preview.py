"""주문 미리보기 표와 주문 불러오기·복사 동작을 검증한다.

- 표: 행 수, 펼친 상품 줄, 매칭 실패·택배 외 배송 행 색, 사은품 열
- 복사: 불러올 때 자동 복사, 메뉴 복사(Ctrl+C), 표에 초점이 있을 때 Ctrl+C 모두
  표로 바꾸기 전(main 2026.11.17.5)과 같은 마크다운이어야 한다.
"""

import json
import os
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path

ROOT = Path(os.environ.get("EF_TEST_ROOT") or Path(__file__).resolve().parents[1])
GIFT_TIERS = ((60000, "볼펜 또는 박스커터"), (100000, "에코백"))

# 표로 바꾸기 전 코드(main ea53267)로 같은 입력을 처리해 얻은 클립보드 글.
EXPECTED_MARKDOWN = {
    "naver": (
        "[ ] 1.홍길동 - 6.1만 · **🎁 사은품: 볼펜 또는 박스커터**\n"
        "▶ [A-ESP32S3]**[ 2 개 ]** - [ESP32-S3 DevKitC-1](https://smartstore.naver.com/higenis/products/111) ( 옵션 : N16R8 )\n"
        "▶ [C-USBC1M]**[ 1 개 ]** - [USB-C 케이블 1m](https://smartstore.naver.com/higenis/products/112) ( 옵션 : 없음 )\n"
        "\n"
        "[ ] 2.김철수 - 3.2만\n"
        "▶ [        ]**[ 1 개 ]** - [STM32 Nucleo-F401RE](https://smartstore.naver.com/higenis/products/999) ( 옵션 : 없음 )\n"
        "\n"
        "[ ] 3.이영희 - 12.8만 · **🎁 사은품: 에코백** **(직접전달)**\n"
        "▶ [L-TDS3P]**[ 1 개 ]** - [LilyGo T-Display S3](https://smartstore.naver.com/higenis/products/113) ( 옵션 : Pro )\n"
        "\n"
    ),
    "coupang": (
        "[ ] 1.최가람 - 6.1만 · **🎁 사은품: 볼펜 또는 박스커터**\n"
        "▶ [C-UNOR4]**[ 1 개 ]** - [아두이노 우노 R4](https://www.coupang.com/vp/products/9001) ( 옵션 : WiFi )\n"
        "▶ [C-BB830]**[ 2 개 ]** - 브레드보드 830 ( 옵션 :  )\n"
        "\n"
        "[ ] 2.정다운 - 13.5만 · **🎁 사은품: 에코백**\n"
        "▶ []**[ 1 개 ]** - 라즈베리파이 5 8GB ( 옵션 :  )\n"
        "\n"
    ),
    "gmarket": (
        "[ ] 1.한소희 - 3.0만\n"
        "▶ [PASS]**[ 3 개 ]** - ESP32 DevKit V1 ( 옵션 : 없음 )\n"
        "  (추가구성 : 핀헤더 추가)\n"
        "\n"
    ),
    "11st": (
        "[ ] 1.오세훈 - 6.6만 · **🎁 사은품: 볼펜 또는 박스커터**\n"
        "▶ [PASS]**[ 2 개 ]** - STM32 블루필 ( 옵션 : 없음 )\n"
        "\n"
    ),
}


def _naver_frame(pd):
    cols = ["주문번호", "수취인명", "수취인연락처1", "통합배송지", "구매자연락처", "배송메세지", "상품명",
            "옵션정보", "수량", "우편번호", "상품번호", "배송방법(구매자 요청)", "최종 상품별 총 주문금액"]
    std = "택배,등기,소포"
    return pd.DataFrame([
        ["2026111700001001", "홍길동", "010-1111-2222", "서울 강남구 1", "010", "", "ESP32-S3 DevKitC-1", "N16R8", 2, "06035", "111", std, 41000],
        ["2026111700001002", "홍길동", "010-1111-2222", "서울 강남구 1", "010", "", "USB-C 케이블 1m", "", 1, "06035", "112", std, 20000],
        ["2026111700002001", "김철수", "010-3333-4444", "부산 해운대구 2", "010", "", "STM32 Nucleo-F401RE", "", 1, "48000", "999", std, 32000],
        ["2026111700003001", "이영희", "010-5555-6666", "대전 유성구 3", "010", "", "LilyGo T-Display S3", "Pro", 1, "34000", "113", "직접전달", 128000],
    ], columns=cols)


def _coupang_frame(pd):
    cols = ["주문번호", "수취인이름", "수취인 주소", "수취인전화번호", "노출상품명(옵션명)", "등록옵션명",
            "구매수(수량)", "배송메세지", "우편번호", "옵션ID", "결제액"]
    return pd.DataFrame([
        ["30001", "최가람", "서울 마포구 5", "010-1212-3434", "아두이노 우노 R4", "WiFi", 1, "문 앞", "04000", 501, 52000],
        ["30001", "최가람", "서울 마포구 5", "010-1212-3434", "브레드보드 830", "", 2, "문 앞", "04000", 502, 9000],
        ["30002", "정다운", "광주 북구 6", "010-5656-7878", "라즈베리파이 5 8GB", "", 1, "", "61000", 999, 135000],
    ], columns=cols)


def _gmarket_frame(pd):
    cols = ["주문번호", "수령인명", "주소", "수령인 전화번호", "수령인 휴대폰", "상품명", "옵션", "수량",
            "배송시 요구사항", "우편번호", "판매금액", "추가구성", "배송비 금액"]
    return pd.DataFrame([
        ["40001", "한소희", "경기 수원시 7", "", "010-9090-1010", "ESP32 DevKit V1", "", 3, "", "16000", 27000, "핀헤더 추가", 3000],
    ], columns=cols)


def _11st_frame(pd):
    cols = ["주문번호", "수취인", "상품명", "옵션", "수량", "주문금액", "휴대폰번호", "전화번호", "우편번호", "주소", "배송메시지"]
    return pd.DataFrame([
        ["50001", "오세훈", "STM32 블루필", "", 2, 66000, "010-2020-3030", "", "12000", "경기 남양주시 8", ""],
    ], columns=cols)


def _tree_snapshot(window):
    from PySide6.QtCore import Qt
    view = getattr(window.ui, "treeView_order_preview", None)
    if view is None:
        return None
    model = view.model()
    rows = []
    for r in range(model.rowCount()):
        cells = [model.index(r, c).data() for c in range(model.columnCount())]
        brush = model.itemFromIndex(model.index(r, 0)).background()
        color = "" if brush.style() == Qt.BrushStyle.NoBrush else brush.color().name()
        parent = model.index(r, 0)
        children = [model.index(i, 0, parent).data() for i in range(model.rowCount(parent))]
        rows.append({"cells": cells, "bg": color, "children": children,
                     "spanned": all(view.isFirstColumnSpanned(i, parent) for i in range(model.rowCount(parent)))})
    return {"rows": rows, "page": window.ui.stackedWidget_order_preview.currentIndex(),
            "legend": window.ui.label_order_legend.isVisibleTo(window.ui.page_order_table),
            "legend_text": window.ui.label_order_legend.text()}


def _pump(app, until, timeout=10.0):
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if until():
            return True
        time.sleep(0.02)
    return False


def run_scenario(name):
    """MainWindow 시나리오(자식 프로세스, os._exit로 종료)."""
    import importlib.util
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtCore import Qt
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
    import pandas as pd

    import sheets.google_sheets_oauth as oauth
    from orders.gift_event import GIFT_EVENT_STORES, GiftEvent

    work = Path.cwd()
    spec = importlib.util.spec_from_file_location("easy_fulfill_order_preview", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    clipboard = QApplication.clipboard()
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(QMessageBox, "warning", side_effect=lambda *a, **k: None), \
            patch.object(QMessageBox, "critical", side_effect=lambda *a, **k: None), \
            patch.object(QMessageBox, "information", side_effect=lambda *a, **k: None):
        window = app_module.MainWindow()
        window._startup_order_index_sync_started = True
        window._begin_order_index_read = lambda **kwargs: kwargs["on_applied"] and kwargs["on_applied"]()
        window._order_gift_event = GiftEvent(True, tuple(GIFT_EVENT_STORES), GIFT_TIERS)
        window.show()
        clipboard.setText("이전 클립보드")

        def load(path):
            with patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(path), "")):
                window.select_excel_file()
            result["loaded"] = _pump(app, lambda: window.is_order_file_valid)

        if name == "naver":
            window.store_type = "naver"
            ok = window._build_naver_orders_from_df(
                _naver_frame(pd), {"111": "A-ESP32S3", "112": "C-USBC1M", "113": "L-TDS3P"})
            window._finish_order_processing(ok is True)
            result["loaded"] = window.is_order_file_valid
        elif name == "coupang":
            path = work / "DeliveryList(2026-11-17)_(0).xlsx"
            _coupang_frame(pd).to_excel(path, index=False)
            window._prefetched_order_mapping = {
                "store_type": "coupang", "ok": True,
                "mapping": {"501": "C-UNOR4", "502": "C-BB830"}, "coupang_vp": {"501": "9001"},
            }
            load(path)
        elif name == "flow":
            # 주문 불러오기 → 우체국 접수 엑셀 → 송장 불러오기(발송 파일) → 접수목록, 버튼 강조와 버튼 동작
            def stage_snapshot():
                window.ui.tabWidget.setCurrentIndex(0)
                app.processEvents()
                names = ("pushButton_load_order", "pushButton_copy_order", "pushButton_intake_excel",
                         "pushButton_clear_order", "pushButton_load_invoice", "pushButton_courier_receipt")
                return {
                    "order_summary": window.ui.label_order_summary.text(),
                    "ship_summary": window.ui.label_ship_summary.text(),
                    "logo_tip": window.ui.label_logo.toolTip(),
                    "logo_pixmap": not window.ui.label_logo.pixmap().isNull(),
                    "stage": window._order_ship_stage,
                    "primary": [n for n in names if getattr(window.ui, n).property("primaryAction")],
                    "enabled": {n: getattr(window.ui, n).isEnabled() for n in names},
                }
            flow = {"start": stage_snapshot()}
            path = work / "DeliveryList(2026-11-17)_(0).xlsx"
            _coupang_frame(pd).to_excel(path, index=False)
            window._prefetched_order_mapping = {
                "store_type": "coupang", "ok": True,
                "mapping": {"501": "C-UNOR4", "502": "C-BB830"}, "coupang_vp": {"501": "9001"},
            }
            load(path)
            flow["loaded"] = stage_snapshot()
            clipboard.setText("")
            window.ui.pushButton_copy_order.click()
            flow["copy_button"] = clipboard.text() == EXPECTED_MARKDOWN["coupang"]
            with patch.object(QMessageBox, "exec", lambda self: 0):
                window.ui.pushButton_intake_excel.click()
            flow["intake_files"] = sorted(p.name.split("_")[0] for p in (work / "output").glob("*.xlsx"))
            flow["intake"] = stage_snapshot()
            # 우체국 송장 엑셀(7행 헤더): 최가람은 주문과 맞고, 정다운은 전화번호가 달라 매칭되지 않는다.
            invoice = work / "1744778498617.xlsx"
            pd.DataFrame([
                ["6000000000001", "최가람", "010-1212-3434", "010-1212-3434", "서울 마포구 5"],
                ["6000000000002", "정다운", "010-0000-0000", "010-0000-0000", "광주 북구 6"],
            ], columns=["등기번호", "수취인명", "수취인 이동통신", "수취인 전화번호", "수취인상세주소"]
            ).to_excel(invoice, index=False, startrow=6)
            window._enqueue_tracking_registration = lambda records: None  # 공유 시트에 쓰지 않는다
            with patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(invoice), "")), \
                    patch.object(QMessageBox, "exec", lambda self: 0):
                window.ui.pushButton_load_invoice.click()
            flow["ship_files"] = sorted(p.name.split("_")[0] + "_" + p.name.split("_")[1]
                                        for p in (work / "output").glob("일괄발송_*.xlsx"))
            flow["shipped"] = stage_snapshot()
            calls = []
            window._tracking_controller.export_courier_receipt = lambda: calls.append("receipt")
            window.ui.pushButton_courier_receipt.click()
            flow["receipt_clicked"] = calls
            window.ui.pushButton_clear_order.click()
            flow["cleared"] = stage_snapshot()
            result["flow"] = flow
            result["loaded"] = True
        elif name == "gmarket":
            path = work / "발송관리.xlsx"
            _gmarket_frame(pd).to_excel(path, index=False)
            load(path)
        elif name == "11st":
            path = work / "20261117_logistics.xls"
            path.write_bytes(b"11st order file")
            frame = _11st_frame(pd)
            with patch.object(app_module, "run_in_background", lambda func, *a, **k: frame.copy()):
                load(path)

        if name == "flow":
            print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
            os._exit(0)
        result["auto_copied"] = clipboard.text()
        result["tree"] = _tree_snapshot(window)
        # 메뉴 복사(Ctrl+C 액션)
        clipboard.setText("")
        window.act_copy.trigger()
        result["menu_copy"] = clipboard.text()
        # 표에 초점이 있을 때 Ctrl+C
        view = getattr(window.ui, "treeView_order_preview", None)
        if view is not None:
            clipboard.setText("")
            view.setFocus()
            view.setCurrentIndex(view.model().index(0, 1))
            app.processEvents()
            QTest.keyClick(view, Qt.Key.Key_C, Qt.KeyboardModifier.ControlModifier)
            app.processEvents()
            result["tree_copy"] = clipboard.text()
            # 초기화 후에는 빈 표, 복사해도 클립보드를 지우지 않는다
            clipboard.setText("그대로")
            window.clear_list()
            window.copy_to_clipboard()
            result["after_clear"] = {"tree": _tree_snapshot(window), "clipboard": clipboard.text()}
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


def scenario(name):
    with tempfile.TemporaryDirectory() as directory:
        env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
        completed = subprocess.run(
            [sys.executable, str(Path(__file__).resolve()), "--scenario", name],
            cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=120,
        )
    line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
    if line is None:
        raise AssertionError(completed.stdout[-2000:] + completed.stderr[-2000:])
    return json.loads(line[len("RESULT "):])


class OrderPreviewModelTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        sys.path.insert(0, str(ROOT))
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def order(self, **kwargs):
        from ui.order_preview import PreviewOrder, PreviewProduct
        base = dict(number=1, recipient="홍길동", amount=61000, amount_text="6.1만", gift="",
                    delivery="택배,등기,소포", products=(PreviewProduct("A-1", 1, "보드", "없음"),),
                    check_codes=True)
        base.update(kwargs)
        return PreviewOrder(**base)

    def test_flags_and_lines(self):
        from ui.order_preview import PreviewProduct, product_line
        self.assertFalse(self.order().code_missing)
        self.assertTrue(self.order(products=(PreviewProduct("        ", 1, "보드", "없음"),)).code_missing)
        # 지마켓·11번가는 코드 매칭을 하지 않으므로 표시하지 않는다
        self.assertFalse(self.order(check_codes=False, products=(PreviewProduct("", 1, "보드", ""),)).code_missing)
        self.assertTrue(self.order(delivery="직접전달").non_parcel)
        self.assertEqual(self.order(delivery="").delivery_text, "택배")
        self.assertEqual(product_line(PreviewProduct(" ", 2, "보드", "R4"), True), "[코드 없음] 2개 · 보드 (옵션: R4)")
        self.assertEqual(product_line(PreviewProduct("PASS", 1, "보드", "없음", "핀"), False), "1개 · 보드 (추가구성: 핀)")

    def test_model_rows_children_and_colors(self):
        from ui.order_preview import (CODE_MISSING_BACKGROUND, NON_PARCEL_BACKGROUND, PreviewProduct,
                                      build_preview_model)
        model = build_preview_model([
            self.order(),
            self.order(number=2, products=(PreviewProduct("", 1, "a", ""), PreviewProduct("B", 1, "b", ""))),
            self.order(number=3, delivery="방문수령", gift="에코백"),
        ])
        self.assertEqual(model.rowCount(), 3)
        self.assertEqual(model.rowCount(model.index(1, 0)), 2)
        from PySide6.QtCore import Qt
        self.assertEqual(model.item(0, 0).background().style(), Qt.BrushStyle.NoBrush)
        self.assertEqual(model.item(1, 0).background().color().name(), CODE_MISSING_BACKGROUND)
        self.assertEqual(model.item(2, 0).background().color().name(), NON_PARCEL_BACKGROUND)
        self.assertEqual([model.index(2, c).data() for c in range(6)], ["3", "홍길동", "6.1만", "에코백", "방문수령", "1"])
        self.assertIn("매칭 실패", model.item(1, 1).toolTip())

    def test_gift_name_matches_markdown_suffix(self):
        from orders.gift_event import GiftEvent
        event = GiftEvent(True, ("naver",), GIFT_TIERS)
        self.assertEqual(event.gift_name("naver", 61000), "볼펜 또는 박스커터")
        self.assertEqual(event.gift_name("naver", 59999), "")
        self.assertEqual(event.gift_name("coupang", 200000), "")
        self.assertIn("볼펜 또는 박스커터", event.markdown_suffix("naver", 61000))


class OrderShipSummaryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        sys.path.insert(0, str(ROOT))
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def test_summaries(self):
        from ui.order_ship_panel import order_summary_html, ship_summary_html
        self.assertEqual(order_summary_html("naver", None), "")
        text = order_summary_html("naver", {"orders": 12, "gifts": 3, "code_missing": 0, "non_parcel": 1})
        self.assertIn("✓ 네이버 12건 · 사은품 3건", text)
        self.assertNotIn("코드 없음", text)
        self.assertIn("송장 매칭 11/12", ship_summary_html(batch=("naver", 12, 11, 12)))
        self.assertIn("#e65100", ship_summary_html(batch=("naver", 12, 11, 12)))  # 덜 맞으면 주황
        self.assertIn("#2e7d32", ship_summary_html(batch=("naver", 12, 12, 12)))
        self.assertIn("송장 5건", ship_summary_html(invoice_rows=5))
        self.assertEqual(ship_summary_html(), "")

    def test_store_badges(self):
        from ui.order_ship_panel import store_badge_pixmap
        image_dir = ROOT / "image"
        for store in ("naver", "coupang", "gmarket", "11st", None):
            pixmap = store_badge_pixmap(store, image_dir, (140, 40), 2.0)
            self.assertFalse(pixmap.isNull(), store)
            self.assertEqual((pixmap.width(), pixmap.height()), (280, 80))
        # 미선택 배지는 흐리게(불투명도 35%) 그린다
        image = store_badge_pixmap(None, image_dir, (140, 40)).toImage()
        alphas = {image.pixelColor(x, y).alpha() for x in range(0, 140, 3) for y in range(0, 40, 3)}
        self.assertLess(max(alphas), 120)


class OrderShipStageTests(unittest.TestCase):
    def test_next_stage(self):
        sys.path.insert(0, str(ROOT))
        from ui.order_ship_panel import next_stage
        self.assertEqual(next_stage(has_orders=False, intake_done=False, batch_done=False), "order")
        self.assertEqual(next_stage(has_orders=True, intake_done=False, batch_done=False), "intake")
        self.assertEqual(next_stage(has_orders=True, intake_done=True, batch_done=False), "invoice")
        # 발송 파일을 만들면 주문 패널은 비워지지만 다음 동작은 접수목록
        self.assertEqual(next_stage(has_orders=False, intake_done=False, batch_done=True), "receipt")

    def test_buttons_follow_work_order(self):
        flow = scenario("flow")["flow"]
        self.assertEqual(flow["start"]["primary"], ["pushButton_load_order"])
        self.assertEqual(flow["start"]["order_summary"], "")
        self.assertEqual(flow["start"]["ship_summary"], "")
        self.assertTrue(flow["start"]["logo_pixmap"])
        self.assertIn("스토어 미선택", flow["start"]["logo_tip"])
        self.assertIn("✓ 쿠팡 2건 · 사은품 2건", flow["loaded"]["order_summary"])
        self.assertIn("코드 없음 1건", flow["loaded"]["order_summary"])
        self.assertTrue(flow["loaded"]["logo_pixmap"])
        self.assertEqual(flow["loaded"]["logo_tip"], "쿠팡 주문")
        self.assertEqual(flow["ship_files"], ["일괄발송_쿠팡"])
        # 발송 파일을 만들면 ① 패널은 비워지고(기존 동작), 묶음 결과는 ② 요약에 남는다
        self.assertEqual(flow["shipped"]["order_summary"], "")
        self.assertIn("✓ 쿠팡 2건", flow["shipped"]["ship_summary"])
        self.assertIn("송장 매칭 1/2", flow["shipped"]["ship_summary"])
        self.assertEqual(flow["cleared"]["ship_summary"], "")
        self.assertFalse(flow["start"]["enabled"]["pushButton_copy_order"])
        self.assertFalse(flow["start"]["enabled"]["pushButton_intake_excel"])
        self.assertEqual(flow["loaded"]["primary"], ["pushButton_intake_excel"])
        self.assertTrue(flow["loaded"]["enabled"]["pushButton_copy_order"])
        self.assertTrue(flow["copy_button"])
        self.assertEqual(flow["intake_files"], ["하이제니스 폼"])
        self.assertEqual(flow["intake"]["primary"], ["pushButton_load_invoice"])
        self.assertEqual(flow["shipped"]["primary"], ["pushButton_courier_receipt"])
        self.assertEqual(flow["receipt_clicked"], ["receipt"])
        self.assertEqual(flow["cleared"]["primary"], ["pushButton_load_order"])
        self.assertFalse(flow["cleared"]["enabled"]["pushButton_copy_order"])


class OrderLoadAndCopyTests(unittest.TestCase):
    def check_copies(self, result, store):
        expected = EXPECTED_MARKDOWN[store]
        self.assertTrue(result["loaded"])
        self.assertEqual(result["auto_copied"], expected)
        self.assertEqual(result["menu_copy"], expected)
        self.assertEqual(result["tree_copy"], expected)
        after = result["after_clear"]
        self.assertEqual(after["tree"]["rows"], [])
        self.assertEqual(after["tree"]["page"], 0)
        self.assertEqual(after["clipboard"], "그대로")

    def test_naver_table_and_copy(self):
        from ui.order_preview import CODE_MISSING_BACKGROUND, NON_PARCEL_BACKGROUND
        result = scenario("naver")
        self.check_copies(result, "naver")
        tree = result["tree"]
        self.assertEqual(tree["page"], 1)
        self.assertEqual([r["cells"] for r in tree["rows"]], [
            ["1", "홍길동", "6.1만", "볼펜 또는 박스커터", "택배", "2"],
            ["2", "김철수", "3.2만", "", "택배", "1"],
            ["3", "이영희", "12.8만", "에코백", "직접전달", "1"],
        ])
        self.assertEqual([r["bg"] for r in tree["rows"]], ["", CODE_MISSING_BACKGROUND, NON_PARCEL_BACKGROUND])
        self.assertEqual(tree["rows"][0]["children"], [
            "[A-ESP32S3] 2개 · ESP32-S3 DevKitC-1 (옵션: N16R8)", "[C-USBC1M] 1개 · USB-C 케이블 1m"])
        self.assertEqual(tree["rows"][1]["children"], ["[코드 없음] 1개 · STM32 Nucleo-F401RE"])
        self.assertTrue(all(r["spanned"] for r in tree["rows"]))
        self.assertTrue(tree["legend"])
        self.assertIn("상품코드 매칭 실패 1건", tree["legend_text"])
        self.assertIn("택배 외 배송 1건", tree["legend_text"])

    def test_coupang_file_load_table_and_copy(self):
        result = scenario("coupang")
        self.check_copies(result, "coupang")
        rows = result["tree"]["rows"]
        self.assertEqual([r["cells"] for r in rows], [
            ["1", "최가람", "6.1만", "볼펜 또는 박스커터", "택배", "2"],
            ["2", "정다운", "13.5만", "에코백", "택배", "1"],
        ])
        self.assertEqual(rows[0]["bg"], "")
        self.assertNotEqual(rows[1]["bg"], "")
        self.assertEqual(rows[1]["children"], ["[코드 없음] 1개 · 라즈베리파이 5 8GB"])

    def test_gmarket_file_load_table_and_copy(self):
        result = scenario("gmarket")
        self.check_copies(result, "gmarket")
        rows = result["tree"]["rows"]
        self.assertEqual([r["cells"] for r in rows], [["1", "한소희", "3.0만", "", "택배", "1"]])
        self.assertEqual(rows[0]["bg"], "")
        self.assertEqual(rows[0]["children"], ["3개 · ESP32 DevKit V1 (추가구성: 핀헤더 추가)"])
        self.assertFalse(result["tree"]["legend"])

    def test_11st_file_load_table_and_copy(self):
        result = scenario("11st")
        self.check_copies(result, "11st")
        rows = result["tree"]["rows"]
        self.assertEqual([r["cells"] for r in rows], [["1", "오세훈", "6.6만", "볼펜 또는 박스커터", "택배", "1"]])
        self.assertEqual(rows[0]["children"], ["2개 · STM32 블루필"])


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--scenario":
        run_scenario(sys.argv[2])
    else:
        unittest.main()
