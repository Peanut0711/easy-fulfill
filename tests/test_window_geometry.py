"""메인 창 기본 크기·가운데 배치·위치 기억을 검증한다."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from shared.window_geometry import (  # noqa: E402
    FRAME_ESTIMATE, centered_position, default_window_size, geometry_to_setting, is_restorable, parse_saved_geometry,
)

MINIMUM = (680, 560)
FHD_100 = (0, 0, 1920, 1032)   # 작업 표시줄 48px
FHD_125 = (0, 0, 1536, 826)
FHD_150 = (0, 0, 1280, 688)


class DefaultSizeTests(unittest.TestCase):
    def test_fhd_100_uses_base_size_under_half_width(self):
        width, height = default_window_size(FHD_100[2:], MINIMUM)
        self.assertEqual((width, height), (900, 800))
        self.assertLess(width + FRAME_ESTIMATE[0], 1920 / 2)
        self.assertGreater(FHD_100[3] - (height + FRAME_ESTIMATE[1]), 150)  # 위아래 여유

    def test_fhd_125_shrinks_by_ratio(self):
        width, height = default_window_size(FHD_125[2:], MINIMUM)
        self.assertEqual((width, height), (721, 638))
        self.assertLess(width + FRAME_ESTIMATE[0], 1536 / 2)
        self.assertLess(height + FRAME_ESTIMATE[1], 826)

    def test_never_smaller_than_minimum(self):
        self.assertEqual(default_window_size(FHD_150[2:], MINIMUM), MINIMUM)

    def test_centered_in_work_area(self):
        x, y = centered_position(FHD_100, (900, 800))
        frame_w, frame_h = FRAME_ESTIMATE
        outer_left, outer_top = x - frame_w // 2, y - (frame_h - frame_w // 2)
        self.assertEqual(outer_left, (1920 - 900 - frame_w) // 2)
        self.assertEqual(outer_top, (1032 - 800 - frame_h) // 2)

    def test_centered_on_secondary_screen_offset(self):
        x, y = centered_position((1920, 0, 1920, 1032), (900, 800))
        self.assertGreaterEqual(x, 1920)


class RestoreTests(unittest.TestCase):
    def test_saved_rect_on_screen_is_restorable(self):
        self.assertTrue(is_restorable((1000, 100, 900, 800), [FHD_100], MINIMUM))

    def test_rect_on_removed_monitor_is_rejected(self):
        self.assertFalse(is_restorable((2500, 100, 900, 800), [FHD_100], MINIMUM))

    def test_title_bar_above_screen_is_rejected(self):
        self.assertFalse(is_restorable((100, 5, 900, 800), [FHD_100], MINIMUM))

    def test_rect_larger_than_screen_after_scale_change_is_rejected(self):
        self.assertFalse(is_restorable((100, 40, 900, 800), [FHD_125], MINIMUM))

    def test_rect_smaller_than_minimum_is_rejected(self):
        self.assertFalse(is_restorable((100, 100, 300, 200), [FHD_100], MINIMUM))

    def test_mostly_offscreen_rect_is_rejected(self):
        self.assertFalse(is_restorable((1850, 100, 900, 800), [FHD_100], MINIMUM))

    def test_setting_round_trip_and_bad_values(self):
        value = geometry_to_setting((10, 40, 900, 800), True)
        self.assertEqual(parse_saved_geometry(value), ((10, 40, 900, 800), True))
        self.assertIsNone(parse_saved_geometry({"x": 1}))
        self.assertIsNone(parse_saved_geometry("900x800"))
        self.assertIsNone(parse_saved_geometry({"x": "a", "y": 0, "w": 1, "h": 1}))


def _scenario(name):
    import importlib.util
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtCore import QRect
    from PySide6.QtWidgets import QApplication, QMessageBox

    app = QApplication([])
    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_geometry", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    settings = Path("database") / "app_settings.json"
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(QMessageBox, "warning", side_effect=lambda *a, **k: None), \
            patch.object(QMessageBox, "critical", side_effect=lambda *a, **k: None):
        spec.loader.exec_module(module)
        area = app.primaryScreen().availableGeometry()
        result["area"] = [area.x(), area.y(), area.width(), area.height()]
        if name == "restore":
            settings.parent.mkdir(exist_ok=True)
            settings.write_text(json.dumps({"main_window_geometry": {
                "x": 20, "y": 40, "w": 760, "h": 740, "maximized": False}}), encoding="utf-8")
        elif name == "offscreen_saved":
            settings.parent.mkdir(exist_ok=True)
            settings.write_text(json.dumps({"main_window_geometry": {
                "x": 5000, "y": 40, "w": 760, "h": 740, "maximized": False}}), encoding="utf-8")
        window = module.MainWindow()
        geometry = window.geometry()
        result["initial"] = [geometry.x(), geometry.y(), geometry.width(), geometry.height()]
        hint = window.minimumSizeHint()
        result["minimum"] = [hint.width(), hint.height()]
        if name == "close_without_show":
            window._close_ready = False
            window.closeEvent(type("E", (), {"accept": lambda self: None, "ignore": lambda self: None})())
            result["saved"] = settings.exists() and "main_window_geometry" in settings.read_text(encoding="utf-8")
        if name == "save":
            window.show()
            app.processEvents()
            window.setGeometry(QRect(30, 50, 770, 745))
            app.processEvents()
            window.close()
            app.processEvents()
            result["saved"] = json.loads(settings.read_text(encoding="utf-8")).get("main_window_geometry")
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


def run_scenario(name):
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


# offscreen 화면은 800×800이고 창 최소 크기가 그보다 조금 작아서(측면 툴바 포함 폭 736), 저장값은 최소 크기 이상·화면 이하로 둔다.
class MainWindowGeometryTests(unittest.TestCase):
    def test_first_run_is_centered_and_not_below_minimum(self):
        result = run_scenario("default")
        x, y, width, height = result["initial"]
        ax, ay, aw, ah = result["area"]
        self.assertGreaterEqual(width, result["minimum"][0])
        self.assertGreaterEqual(height, result["minimum"][1])
        self.assertAlmostEqual((x - ax) * 2 + width, aw, delta=40)

    def test_saved_geometry_is_restored(self):
        self.assertEqual(run_scenario("restore")["initial"], [20, 40, 760, 740])

    def test_saved_geometry_off_screen_falls_back_to_default(self):
        result = run_scenario("offscreen_saved")
        self.assertLess(result["initial"][0], result["area"][2])

    def test_close_saves_geometry_after_show(self):
        saved = run_scenario("save")["saved"]
        self.assertEqual([saved["x"], saved["y"], saved["w"], saved["h"]], [30, 50, 770, 745])
        self.assertFalse(saved["maximized"])

    def test_never_shown_window_does_not_save(self):
        self.assertFalse(run_scenario("close_without_show")["saved"])


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--scenario":
        _scenario(sys.argv[2])
    unittest.main()
