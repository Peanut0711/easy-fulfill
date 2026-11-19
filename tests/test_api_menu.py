"""개발 중 API 기능을 「⚙ 설정 › 프로그램」의 「개발 중 기능」 묶음에 모으고 이 PC 스위치로 켜고 끄는 것을 검증한다.

디자인 개편 검토 리스트 R7(2026-11-19 사용자 결정): API·우체국 포털 기능은 개발 중이라 실무자는 쓰지 않는다.
- 메뉴바에 API 메뉴가 없다.
- 「개발 중 기능」 묶음에 진입점 3개, 확인·진단 9개, 실제 동작 4개 버튼이 있고 같은 QAction을 부른다.
- 기본은 모두 비활성(QAction 자체가 꺼짐)이고 「개발 중 기능 켜기(이 PC)」를 체크하면 켜진다(이 PC 설정에 저장).
- 실제 동작 4개(네이버 발송처리, 우체국 실제 접수, 단건·다건 실제 인쇄)는 켠 상태에서도 확인 창에서 「아니요」를
  누르면 API·포털 호출이 일어나지 않는다. QAction을 그대로 눌러(trigger) 연결까지 확인한다.

MainWindow를 만들어야 하므로 자식 프로세스에서 실행한다(앱 종료 시 스레드 정리를 건너뛰려고 os._exit).
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

DIAGNOSTIC_ACTIONS = [
    "act_epost_test_receipt", "act_epost_portal_diagnostic", "act_epost_portal_lookup",
    "act_epost_portal_print_popup", "act_epost_portal_oz_viewer", "act_epost_portal_output_confirm",
    "act_epost_portal_reprint_popup", "act_epost_portal_reprint_oz_viewer", "act_epost_oz_print_dialog",
]
REAL_ACTIONS = ["act_dispatch", "act_epost_real_receipt", "act_epost_actual_print", "act_epost_batch_actual_print"]
ENTRY_ACTIONS = ["act_api_load", "act_epost_portal_login", "act_epost_print_targets"]
ALL_DEV_ACTIONS = ENTRY_ACTIONS + DIAGNOSTIC_ACTIONS + REAL_ACTIONS


def run_scenario():
    import importlib.util
    from types import SimpleNamespace
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from openpyxl import Workbook
    from PySide6.QtWidgets import QApplication, QFileDialog, QMenu, QMessageBox, QPushButton

    import sheets.google_sheets_oauth as oauth
    from orders.invoice import INVOICE_COLUMNS
    from shipping.post_parcel import RecipientAddressSuggestion

    work = Path.cwd()
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(list(INVOICE_COLUMNS))
    row = {"수취인명": "김철수", "우편번호": "06035", "수취인 주소": "서울 강남구 테헤란로 1",
           "수취인 전화번호": "010-1111-2222", "수취인 이동통신": "010-1111-2222", "상품명": "상품"}
    sheet.append([row.get(column, "") for column in INVOICE_COLUMNS])
    workbook.save(work / "invoice.xlsx")

    spec = importlib.util.spec_from_file_location("easy_fulfill_api_menu", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    result = {}
    questions, calls = [], []

    def answer_no(*args, **kwargs):
        questions.append(args[1] if len(args) > 1 else "")
        return QMessageBox.StandardButton.No

    silent = lambda *a, **k: QMessageBox.StandardButton.Ok  # noqa: E731
    suggestion = RecipientAddressSuggestion("06035", "서울 강남구 테헤란로 1", "-")
    from shipping.post_parcel_receipt_store import ParcelReceiptStore as RealStore

    class SpyStore(RealStore):
        def __init__(self):
            super().__init__(work / "receipts.sqlite3")

        def begin_submission(self, *args, **kwargs):
            calls.append("begin_submission")
            return super().begin_submission(*args, **kwargs)

        def list_portal_print_confirmed(self):
            return [SimpleNamespace(received_at="2026-11-18 10:00:00"),
                    SimpleNamespace(received_at="2026-11-18 11:00:00")]

    class SpyDispatchThread:
        def __init__(self, *args, **kwargs):
            calls.append("naver_dispatch")

    with patch.object(oauth, "get_authorized_gspread_client", lambda *a, **k: object()), \
            patch.object(app_module, "_standalone_open_config_ws", lambda gc: object()), \
            patch.object(app_module, "_read_config_values_map", lambda ws: {"epost_postcode_api_key": "K"}), \
            patch.object(app_module, "resolve_recipient_address", lambda address, key: [suggestion]), \
            patch.object(app_module, "submit_real_order", lambda *a, **k: calls.append("submit_real_order")), \
            patch.object(app_module, "NaverDispatchThread", SpyDispatchThread), \
            patch.object(app_module, "ParcelReceiptStore", SpyStore), \
            patch.object(app_module.QProcess, "start", lambda *a, **k: calls.append("qprocess_start")), \
            patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(work / "invoice.xlsx"), "")), \
            patch.object(QMessageBox, "warning", side_effect=silent), \
            patch.object(QMessageBox, "critical", side_effect=silent), \
            patch.object(QMessageBox, "information", side_effect=silent):
        window = app_module.MainWindow()
        app.processEvents()
        result["menus"] = [m.title() for m in window.menuBar().findChildren(QMenu)
                            if m.menuAction() in window.menuBar().actions()]
        result["real_icons_set"] = all(not getattr(window, n).icon().isNull() for n in REAL_ACTIONS)

        page = window.ui.scrollArea_settings_program
        buttons = {b.objectName(): b for b in page.findChildren(QPushButton)
                   if b.objectName().startswith(("pushButton_dev_", "pushButton_diag_", "pushButton_real_"))}
        result["buttons"] = sorted(buttons)
        result["real_roles"] = sorted({buttons["pushButton_real_" + n.removeprefix("act_")].property("role")
                                       for n in REAL_ACTIONS})
        switch = window.checkBox_dev_features

        def states():
            return {
                "checked": switch.isChecked(),
                "actions_on": sorted(n for n in ALL_DEV_ACTIONS if getattr(window, n).isEnabled()),
                "buttons_on": sorted(n for n, b in buttons.items() if b.isEnabled()),
                "saved": window.get_app_setting("dev_features_enabled", None),
            }

        triggered = []
        for name in DIAGNOSTIC_ACTIONS:
            action = getattr(window, name)
            action.triggered.disconnect()  # 실제 진단 실행(파일 대화상자·포털)은 하지 않는다
            action.triggered.connect(lambda _=False, n=name: triggered.append(n))
        result["default"] = states()
        for name in DIAGNOSTIC_ACTIONS:  # 꺼져 있으면 눌러도 아무 일 없음
            buttons["pushButton_diag_" + name.removeprefix("act_")].click()
        result["triggered_when_off"] = list(triggered)
        switch.setChecked(True)
        result["enabled"] = states()
        for name in DIAGNOSTIC_ACTIONS:
            buttons["pushButton_diag_" + name.removeprefix("act_")].click()
        result["diag_triggered"] = triggered[len(result["triggered_when_off"]):]

        # 실제 동작: 확인 창에서 No → 외부 호출 없음
        window._naver_dispatch_records = [{"orderId": "2026111800001", "invoice": "6012345678901"}]
        result["real"] = {}
        with patch.object(QMessageBox, "question", side_effect=answer_no):
            for name in REAL_ACTIONS:
                questions.clear()
                calls.clear()
                getattr(window, name).trigger()
                result["real"][name] = {
                    "questions": list(questions),
                    "calls": list(calls),
                    "process": getattr(window, "_epost_oz_print_dialog_process", None) is not None,
                }
        switch.setChecked(False)
        result["disabled_again"] = states()
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


class DevFeaturesTests(unittest.TestCase):
    def test_no_api_menu_in_menubar(self):
        self.assertFalse([title for title in scenario()["menus"] if "API" in title])
        self.assertTrue(scenario()["real_icons_set"])

    def test_dev_group_has_all_buttons(self):
        expected = sorted(["pushButton_dev_" + n.removeprefix("act_") for n in ENTRY_ACTIONS]
                          + ["pushButton_diag_" + n.removeprefix("act_") for n in DIAGNOSTIC_ACTIONS]
                          + ["pushButton_real_" + n.removeprefix("act_") for n in REAL_ACTIONS])
        self.assertEqual(scenario()["buttons"], expected)
        self.assertEqual(scenario()["real_roles"], ["danger"])

    def test_off_by_default_and_switch_enables_this_pc(self):
        result = scenario()
        self.assertEqual(result["default"], {"checked": False, "actions_on": [], "buttons_on": [], "saved": None})
        self.assertEqual(result["triggered_when_off"], [])
        self.assertTrue(result["enabled"]["checked"])
        self.assertEqual(result["enabled"]["actions_on"], sorted(ALL_DEV_ACTIONS))
        self.assertEqual(result["enabled"]["buttons_on"], result["buttons"])
        self.assertIs(result["enabled"]["saved"], True)
        self.assertEqual(result["disabled_again"]["actions_on"], [])
        self.assertIs(result["disabled_again"]["saved"], False)

    def test_diagnostic_buttons_trigger_same_actions(self):
        self.assertEqual(scenario()["diag_triggered"], DIAGNOSTIC_ACTIONS)

    def test_real_actions_do_nothing_when_confirmation_is_no(self):
        real = scenario()["real"]
        self.assertEqual(set(real), set(REAL_ACTIONS))
        for name, outcome in real.items():
            # 확인 창까지 실제로 도달했고(질문 1번), No 뒤에는 접수·발송·인쇄 요청이 없다.
            self.assertEqual(len(outcome["questions"]), 1, (name, outcome))
            self.assertEqual(outcome["calls"], [], name)
            self.assertFalse(outcome["process"], name)
        # 각 동작의 마지막 확인 창(실행 직전)까지 갔는지 제목으로 확인한다.
        self.assertEqual({name: outcome["questions"][0] for name, outcome in real.items()}, {
            "act_dispatch": "API 발송처리",
            "act_epost_real_receipt": "우체국 실제 접수 확인 확인",
            "act_epost_actual_print": "우체국 단건 실제 인쇄",
            "act_epost_batch_actual_print": "우체국 다건 실제 인쇄",
        })


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
