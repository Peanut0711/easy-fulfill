"""단축키(알림·상태바 계획 E-1)를 검증한다.

- 탭 이동 Ctrl+1~6, 송장 불러오기 Ctrl+Shift+O
- 메뉴 동작·창 단축키가 겹치지 않음(탭 안 단축키 Ctrl+Enter도 메뉴와 겹치지 않음)
- 「도움말 › 단축키」 목록에 탭 이동·메뉴 단축키·탭 안 단축키가 모두 보임
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run_scenario():
    import importlib.util
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QAction, QShortcut
    from PySide6.QtWidgets import QApplication, QDialog, QTableWidget

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_shortcuts", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")):
        window = module.MainWindow()
        app.processEvents()
        result["action_keys"] = [a.shortcut().toString() for a in window.findChildren(QAction)
                                 if not a.shortcut().isEmpty()]
        window_shortcuts = [s for s in window.findChildren(QShortcut)
                            if s.context() == Qt.ShortcutContext.WindowShortcut]
        result["window_keys"] = [s.key().toString() for s in window_shortcuts]
        result["tab_local_keys"] = [s.key().toString() for s in window.findChildren(QShortcut)
                                    if s.context() != Qt.ShortcutContext.WindowShortcut]
        result["load_invoice"] = window.act_load_invoice.shortcut().toString()
        tabs = window.ui.tabWidget
        moved = []
        for shortcut in sorted(window_shortcuts, key=lambda s: s.key().toString()):
            shortcut.activated.emit()
            moved.append(tabs.currentWidget().objectName())
        result["moved"] = moved
        result["rows"] = window.shortcut_help_rows()
        shown = {}

        def fake_exec(dialog):
            table = dialog.findChild(QTableWidget, "tableWidget_shortcuts")
            shown["title"] = dialog.windowTitle()
            shown["rows"] = table.rowCount()
            return 0

        with patch.object(QDialog, "exec", fake_exec):
            window.act_shortcut_help.trigger()
        result["dialog"] = shown
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


class ShortcutTests(unittest.TestCase):
    def test_no_duplicate_shortcuts(self):
        result = scenario()
        keys = result["action_keys"] + result["window_keys"]
        duplicates = {key: n for key, n in Counter(keys).items() if n > 1}
        self.assertEqual(duplicates, {})
        for key in result["tab_local_keys"]:
            self.assertNotIn(key, keys)

    def test_tab_shortcuts_move_in_order(self):
        result = scenario()
        self.assertEqual(sorted(result["window_keys"]), [f"Ctrl+{i}" for i in range(1, 7)])
        self.assertEqual(result["moved"], ["tab", "tab_tracking", "tab_quick_invoice", "tab_documents",
                                           "tab_detail_page", "tab_settings"])
        self.assertEqual(result["load_invoice"], "Ctrl+Shift+O")

    def test_shortcut_help_lists_everything(self):
        result = scenario()
        rows = result["rows"]
        self.assertEqual(sum(1 for row in rows if row[0] == "탭 이동"), 6)
        self.assertIn(["탭 이동", "「설정」 탭", "Ctrl+6"], rows)
        self.assertTrue(any(row[1] == "송장 불러오기" and row[2] == "Ctrl+Shift+O" for row in rows))
        self.assertTrue(any(row[2] == "Ctrl+Enter" for row in rows))
        self.assertEqual(result["dialog"], {"title": "단축키", "rows": len(rows)})


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
