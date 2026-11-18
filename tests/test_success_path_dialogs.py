"""매일 흐름이 성공할 때 확인 창이 뜨지 않는지 센다(알림·상태바 계획 C-3).

흐름: 주문 불러오기(쿠팡) → 복사 → 우체국 접수 엑셀 → 송장 불러오기(일괄 발송 파일) → 단건 송장(수동 1건).
성공 경로에서 뜨는 QMessageBox(정적 함수·exec)와 QDialog.exec 수를 센다. 결과는 알림 띠와 상태바로만 보여야 한다.

``EF_TEST_ROOT``로 다른 작업 폴더(예: 바꾸기 전 커밋)를 가리키면 같은 흐름으로 전 값을 잴 수 있다.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(os.environ.get("EF_TEST_ROOT") or Path(__file__).resolve().parents[1])
HERE = Path(__file__).resolve().parent


def run_scenario():
    import importlib.util
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(HERE.parent))
    sys.path.insert(0, str(ROOT))  # 잴 작업 폴더의 모듈을 먼저 쓴다
    import pandas as pd
    from PySide6.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox

    import sheets.google_sheets_oauth as oauth
    from tests.test_order_preview import _coupang_frame, _pump

    work = Path.cwd()
    spec = importlib.util.spec_from_file_location("easy_fulfill_success_path", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    dialogs = []

    def record(kind, answer):
        def inner(*args, **kwargs):
            title = args[1] if len(args) > 1 and isinstance(args[1], str) else ""
            dialogs.append(f"{kind}:{title}")
            return answer
        return inner

    def record_exec(self):
        dialogs.append(f"exec:{self.windowTitle()}")
        return 0

    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")):
        window = app_module.MainWindow()
        window._startup_order_index_sync_started = True
        window._begin_order_index_read = lambda **kwargs: kwargs["on_applied"] and kwargs["on_applied"]()
        window._enqueue_tracking_registration = lambda records: None  # 공유 시트에 쓰지 않는다
        window.show()
        app.processEvents()
        steps = {}
        with patch.object(QMessageBox, "information", side_effect=record("information", QMessageBox.StandardButton.Ok)), \
                patch.object(QMessageBox, "warning", side_effect=record("warning", QMessageBox.StandardButton.Ok)), \
                patch.object(QMessageBox, "critical", side_effect=record("critical", QMessageBox.StandardButton.Ok)), \
                patch.object(QMessageBox, "question", side_effect=record("question", QMessageBox.StandardButton.Yes)), \
                patch.object(QMessageBox, "exec", record_exec), \
                patch.object(QDialog, "exec", record_exec):
            def step(name, action):
                before = len(dialogs)
                action()
                steps[name] = dialogs[before:]

            order = work / "DeliveryList(2026-11-17)_(0).xlsx"
            _coupang_frame(pd).to_excel(order, index=False)
            window._prefetched_order_mapping = {
                "store_type": "coupang", "ok": True,
                "mapping": {"501": "C-UNOR4", "502": "C-BB830"}, "coupang_vp": {"501": "9001"},
            }

            def load_order():
                with patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(order), "")):
                    window.select_excel_file()
                _pump(app, lambda: window.is_order_file_valid)

            step("주문 불러오기", load_order)
            step("복사", window.ui.pushButton_copy_order.click)
            step("우체국 접수 엑셀", window.ui.pushButton_intake_excel.click)
            invoice = work / "1744778498617.xlsx"
            pd.DataFrame([
                ["6000000000001", "최가람", "010-1212-3434", "010-1212-3434", "서울 마포구 5"],
            ], columns=["등기번호", "수취인명", "수취인 이동통신", "수취인 전화번호", "수취인상세주소"]
            ).to_excel(invoice, index=False, startrow=6)

            def load_invoice():
                with patch.object(QFileDialog, "getOpenFileName", lambda *a, **k: (str(invoice), "")):
                    window.ui.pushButton_load_invoice.click()
                _pump(app, lambda: bool(list((work / "output").glob("일괄발송_*.xlsx"))), timeout=10)

            step("송장 불러오기(일괄 발송 파일)", load_invoice)

            def quick_manual():
                window.ui.comboBox_store_select.setCurrentText("수동")
                window.ui.lineEdit_quick_manual_name.setText("홍길동")
                window.ui.lineEdit_quick_manual_phone.setText("010-1234-5678")
                window.ui.lineEdit_quick_manual_address.setText("서울 강남구 테헤란로 1")
                window.ui.pushButton_quick_excel_gen.click()

            step("단건 송장(수동)", quick_manual)
        result = {
            "steps": steps,
            "files": sorted(p.name.split("_")[0] for p in (work / "output").glob("*.xlsx")),
            "notice": getattr(getattr(window, "notice_bar", None), "text", lambda: "")(),
        }
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


def measure(root=None):
    with tempfile.TemporaryDirectory() as directory:
        env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
        if root:
            env["EF_TEST_ROOT"] = str(root)
        completed = subprocess.run(
            [sys.executable, str(Path(__file__).resolve()), "--scenario"],
            cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=180,
        )
    line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
    if line is None:
        raise AssertionError(completed.stdout[-3000:] + completed.stderr[-3000:])
    return json.loads(line[len("RESULT "):])


class SuccessPathDialogTests(unittest.TestCase):
    def test_daily_flow_shows_no_dialogs_on_success(self):
        result = measure()
        # 흐름이 실제로 끝까지 갔는지(파일이 만들어졌는지) 먼저 본다.
        self.assertIn("일괄발송", result["files"])
        self.assertTrue(any(name.startswith("퀵") for name in result["files"]), result["files"])
        self.assertEqual({step: shown for step, shown in result["steps"].items() if shown}, {})
        self.assertIn("송장 엑셀 파일이 생성되었습니다", result["notice"])  # 마지막 단계(단건 송장)의 알림


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
