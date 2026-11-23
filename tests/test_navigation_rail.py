"""메인 탭 측면 툴바(탭 이름 줄 대신 왼쪽 세로 버튼)를 검증한다.

- 탭 이름 줄은 숨고, 왼쪽 툴바에 짧은 이름 버튼 8개가 있다(설정은 맨 아래, 발주는 주문 아래, 문의는 설정 바로 위).
- 버튼을 누르면 그 탭으로 가고, 단축키·코드로 탭이 바뀌어도 툴바 선택 표시가 따라간다.
- 툴팁에 전체 탭 이름과 단축키가 보이고, 오른쪽 클릭 메뉴로 툴바를 숨길 수 없다.
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
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication, QToolBar, QToolButton

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_nav_rail", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")):
        window = module.MainWindow()
        window.show()
        app.processEvents()
        tabs = window.ui.tabWidget
        rail = window.findChild(QToolBar, "navigationRail")
        result["tab_bar_visible"] = tabs.tabBar().isVisible()
        result["rail_area"] = window.toolBarArea(rail) == Qt.ToolBarArea.LeftToolBarArea
        result["rail_toggle_visible"] = rail.toggleViewAction().isVisible()
        buttons = sorted(rail.findChildren(QToolButton, "navigationButton"), key=lambda b: b.geometry().y())
        result["labels"] = [b.defaultAction().text() for b in buttons]
        result["settings_gap"] = buttons[-1].geometry().top() - buttons[-2].geometry().bottom()
        result["other_gap"] = buttons[1].geometry().top() - buttons[0].geometry().bottom()

        def checked():
            return [name for name, action in window.nav_rail_actions.items() if action.isChecked()]

        result["initial"] = [tabs.currentWidget().objectName(), checked()]
        window.nav_rail_actions["tab_tracking"].trigger()
        app.processEvents()
        result["after_click"] = [tabs.currentWidget().objectName(), checked()]
        ctrl4 = next(s for s in window._tab_shortcuts if s.key().toString() == "Ctrl+4")
        ctrl4.activated.emit()
        app.processEvents()
        result["after_shortcut"] = [tabs.currentWidget().objectName(), checked()]
        settings = next(tabs.widget(i) for i in range(tabs.count()) if tabs.widget(i).objectName() == "tab_settings")
        tabs.setCurrentWidget(settings)  # 상태 칩 등 코드로 설정 페이지에 갈 때
        app.processEvents()
        result["after_code"] = [tabs.currentWidget().objectName(), checked()]
        result["tooltips"] = {name: action.toolTip() for name, action in window.nav_rail_actions.items()}
        hint = window.minimumSizeHint()
        window.resize(hint)
        app.processEvents()
        result["visible_at_minimum"] = [b.defaultAction().text() for b in buttons if b.isVisible()]
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


class NavigationRailTests(unittest.TestCase):
    def test_rail_replaces_tab_bar(self):
        result = scenario()
        self.assertFalse(result["tab_bar_visible"])
        self.assertTrue(result["rail_area"])
        self.assertFalse(result["rail_toggle_visible"])
        self.assertEqual(result["labels"], ["주문", "발주", "추적", "송장", "문서", "상세", "문의", "설정"])

    def test_settings_is_pushed_to_bottom(self):
        result = scenario()
        self.assertGreater(result["settings_gap"], result["other_gap"] + 50)

    def test_rail_and_tabs_stay_in_sync(self):
        result = scenario()
        self.assertEqual(result["initial"], ["tab", ["tab"]])
        self.assertEqual(result["after_click"], ["tab_tracking", ["tab_tracking"]])
        self.assertEqual(result["after_shortcut"], ["tab_documents", ["tab_documents"]])
        self.assertEqual(result["after_code"], ["tab_settings", ["tab_settings"]])

    def test_no_button_hidden_at_minimum_size(self):
        # 버튼이 많아져 창 최소 높이보다 길어지면 「»」 확장 메뉴로 숨는다(문의 탭 추가 때 설정이 숨었음).
        result = scenario()
        self.assertEqual(result["visible_at_minimum"], result["labels"])

    def test_tooltips_show_full_name_and_shortcut(self):
        tips = scenario()["tooltips"]
        self.assertEqual(tips["tab"], "주문·발송 (Ctrl+1)")
        self.assertEqual(tips["tab_documents"], "문서 만들기 (Ctrl+4)")
        self.assertEqual(tips["tab_settings"], "설정 (Ctrl+6)")
        self.assertEqual(tips["tab_inquiries"], "문의 (Ctrl+7)")
        self.assertEqual(tips["tab_naver_orders"], "네이버 발주 (Ctrl+8)")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
