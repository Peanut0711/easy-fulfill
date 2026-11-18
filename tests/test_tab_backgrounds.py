"""코드로 만든 스크롤 영역(문서 만들기·상세페이지)이 창 배경색(회색)을 칠하지 않는지 검증한다(알림·상태바 계획 D-3).

다른 탭은 탭 바탕(흰색)이 그대로 보이는데, 스크롤 영역이 바탕을 칠하면 이 두 탭만 회색으로 보였다.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run_scenario():
    import importlib.util
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtWidgets import QApplication, QScrollArea

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_backgrounds", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")):
        window = module.MainWindow()
        tabs = window.ui.tabWidget
        for name in ("tab_documents", "tab_detail_page"):
            page = next(tabs.widget(i) for i in range(tabs.count()) if tabs.widget(i).objectName() == name)
            tabs.setCurrentWidget(page)  # 문서 만들기는 처음 들어갈 때 화면을 만든다
            app.processEvents()
            area = page.findChild(QScrollArea)
            result[name] = [area.viewport().autoFillBackground(), area.widget().autoFillBackground()]
    print("RESULT " + json.dumps(result), flush=True)
    os._exit(0)


class TabBackgroundTests(unittest.TestCase):
    def test_scroll_tabs_do_not_paint_window_background(self):
        with tempfile.TemporaryDirectory() as directory:
            env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
            completed = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), "--scenario"],
                cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=180,
            )
        line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
        self.assertIsNotNone(line, completed.stdout[-2000:] + completed.stderr[-2000:])
        self.assertEqual(json.loads(line[len("RESULT "):]),
                         {"tab_documents": [False, False], "tab_detail_page": [False, False]})


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
