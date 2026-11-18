"""단건 송장 탭(알림·상태바 계획 D-2)을 검증한다.

- 묶음 제목 「단건 송장 엑셀」, 「생성」·「초기화」는 맨 아래 줄(생성은 강조)
- 수동 모드에서는 이름·전화번호·기본주소를 다 넣어야 「생성」이 켜지고(빈 칸은 툴팁), 클립보드 모드는 항상 켜짐
- Ctrl+Enter(탭 안)로 생성, 켜져 있을 때만
- 「오늘 만든 단건 송장」 목록(오늘 파일만, 최근 것부터, 수취인 표시)

MainWindow를 만들어야 하므로 자식 프로세스에서 실행한다.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run_scenario():
    import importlib.util
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from openpyxl import Workbook
    from PySide6.QtGui import QShortcut
    from PySide6.QtWidgets import QApplication, QMessageBox

    import sheets.google_sheets_oauth as oauth

    work = Path.cwd()
    (work / "output").mkdir()
    today = datetime.now()

    def make(name, recipient):
        workbook = Workbook()
        workbook.active.append(["주문번호", "수취인명", "우편번호"])
        workbook.active.append(["", recipient, ""])
        workbook.save(work / "output" / name)

    make(f"퀵_네이버_{today:%Y%m%d}093000.xlsx", "김오전")
    make(f"퀵_수동_{today:%Y%m%d}141500.xlsx", "이오후")
    make(f"퀵_쿠팡_{today - timedelta(days=1):%Y%m%d}101010.xlsx", "어제사람")

    spec = importlib.util.spec_from_file_location("easy_fulfill_quick_tab", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])
    result = {}
    silent = lambda *a, **k: QMessageBox.StandardButton.Ok  # noqa: E731
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(QMessageBox, "warning", side_effect=silent), \
            patch.object(QMessageBox, "information", side_effect=silent):
        window = module.MainWindow()
        app.processEvents()
        ui = window.ui
        grid = ui.groupBox.layout()

        def cell(widget):
            index = grid.indexOf(widget)
            return list(grid.getItemPosition(index)[:2])

        result["title"] = ui.groupBox.title()
        result["gen_cell"] = cell(ui.pushButton_quick_excel_gen)
        result["reset_cell"] = cell(ui.pushButton_quick_excel_reset)
        result["gen_role"] = ui.pushButton_quick_excel_gen.property("role")
        result["shortcuts"] = sorted(s.key().toString() for s in ui.tab_quick_invoice.findChildren(QShortcut))

        ui.comboBox_store_select.setCurrentText("수동")
        for name in ("lineEdit_quick_manual_name", "lineEdit_quick_manual_phone", "lineEdit_quick_manual_address"):
            getattr(ui, name).clear()
        states = [(ui.pushButton_quick_excel_gen.isEnabled(), ui.pushButton_quick_excel_gen.toolTip())]
        ui.lineEdit_quick_manual_name.setText("홍길동")
        ui.lineEdit_quick_manual_phone.setText("010-1234-5678")
        states.append((ui.pushButton_quick_excel_gen.isEnabled(), ui.pushButton_quick_excel_gen.toolTip()))
        created = []
        window.generate_quick_excel = lambda: created.append("gen")
        ui.pushButton_quick_excel_gen.clicked.disconnect()
        ui.pushButton_quick_excel_gen.clicked.connect(window.generate_quick_excel)
        window._on_quick_excel_shortcut()  # 꺼져 있으면 아무 일도 없다
        ui.lineEdit_quick_manual_address.setText("서울 강남구 테헤란로 1")
        states.append((ui.pushButton_quick_excel_gen.isEnabled(), ui.pushButton_quick_excel_gen.toolTip()))
        window._on_quick_excel_shortcut()
        result["created"] = created
        ui.lineEdit_quick_manual_name.clear()
        ui.comboBox_store_select.setCurrentText("네이버")  # 클립보드 모드는 항상 켠다
        states.append((ui.pushButton_quick_excel_gen.isEnabled(), ui.pushButton_quick_excel_gen.toolTip()))
        result["states"] = states

        window._refresh_quick_today_list()
        lst = window.listWidget_quick_today
        result["items"] = [lst.item(i).text() for i in range(lst.count())]
        result["item_paths"] = [Path(lst.item(i).data(256)).name for i in range(lst.count())]
        for path in (work / "output").glob("*.xlsx"):
            path.unlink()
        window._refresh_quick_today_list()
        result["empty"] = [lst.item(i).text() for i in range(lst.count())]
        result["empty_selectable"] = bool(lst.item(0).flags() & module.Qt.ItemFlag.ItemIsEnabled)
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


class QuickInvoiceTabTests(unittest.TestCase):
    def test_layout(self):
        result = scenario()
        self.assertEqual(result["title"], "단건 송장 엑셀")
        self.assertEqual(result["reset_cell"], [8, 0])
        self.assertEqual(result["gen_cell"], [8, 2])  # 마지막 줄 오른쪽
        self.assertEqual(result["gen_role"], "primary")
        self.assertIn("Ctrl+Return", result["shortcuts"])

    def test_generate_enabled_only_when_required_fields_filled(self):
        states = scenario()["states"]
        self.assertEqual(states[0][0], False)
        self.assertIn("이름, 전화번호, 기본주소", states[0][1])
        self.assertEqual(states[1], [False, "빈 칸을 채우면 누를 수 있습니다: 기본주소"])
        self.assertEqual(states[2][0], True)
        self.assertIn("Ctrl+Enter", states[2][1])
        self.assertEqual(states[3][0], True)  # 클립보드 모드
        self.assertIn("클립보드", states[3][1])
        self.assertEqual(scenario()["created"], ["gen"])  # 꺼져 있을 때 단축키는 무시

    def test_today_list(self):
        result = scenario()
        self.assertEqual(len(result["items"]), 2)  # 어제 파일은 빠진다
        self.assertTrue(result["items"][0].startswith("14:15  이오후"))  # 최근 것부터
        self.assertTrue(result["items"][1].startswith("09:30  김오전"))
        self.assertTrue(result["items"][1].endswith("· 네이버"))
        self.assertTrue(result["item_paths"][0].startswith("퀵_수동_"))
        self.assertEqual(result["empty"], ["오늘 만든 단건 송장이 없습니다."])
        self.assertFalse(result["empty_selectable"])


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
