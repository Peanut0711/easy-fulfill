"""팝업 창(상세 HTML 편집기·이미지 보기) 크기·위치(알림·상태바 계획 E-2)를 검증한다."""

import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from shared.window_geometry import FRAME_ESTIMATE, dialog_default_size  # noqa: E402

FHD_100 = (1920, 1032)
FHD_125 = (1536, 826)
FHD_150 = (1280, 688)


class DialogDefaultSizeTests(unittest.TestCase):
    def test_original_size_when_it_fits(self):
        self.assertEqual(dialog_default_size(FHD_100, (1280, 760)), (1280, 760))
        self.assertEqual(dialog_default_size(FHD_100, (1024, 768)), (1024, 768))

    def test_shrinks_to_work_area_at_higher_scaling(self):
        width, height = dialog_default_size(FHD_125, (1280, 760))
        self.assertEqual(width, 1280)
        self.assertLessEqual(height + FRAME_ESTIMATE[1], int(FHD_125[1] * 0.9))  # 프레임까지 작업 영역 90% 안
        width, height = dialog_default_size(FHD_150, (1280, 760))
        self.assertLessEqual(width, int(FHD_150[0] * 0.9))
        self.assertLessEqual(height + FRAME_ESTIMATE[1], int(FHD_150[1] * 0.9))


class DialogGeometryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def owner(self):
        from PySide6.QtWidgets import QWidget

        class Owner(QWidget):
            def __init__(self):
                super().__init__()
                self.settings = {}

            def get_app_setting(self, key, default=None):
                return self.settings.get(key, default)

            def set_app_setting(self, key, value):
                self.settings[key] = value

        return Owner()

    def test_default_fits_screen_then_restores_saved(self):
        from PySide6.QtCore import QRect
        from PySide6.QtGui import QGuiApplication
        from PySide6.QtWidgets import QDialog
        from shared.dialog_geometry import apply_dialog_geometry, save_dialog_geometry

        owner = self.owner()
        area = QGuiApplication.primaryScreen().availableGeometry()
        dialog = QDialog(owner)
        apply_dialog_geometry(dialog, "test_geometry", (1280, 760))
        geometry = dialog.geometry()
        self.assertLessEqual(geometry.width(), min(1280, area.width()))
        self.assertLessEqual(geometry.height(), min(760, area.height()))
        self.assertTrue(area.contains(geometry.center()))  # 화면 가운데 쪽

        moved = QRect(area.x() + 20, area.y() + 60, 400, 300)
        dialog.setGeometry(moved)
        save_dialog_geometry(dialog, "test_geometry")
        self.assertEqual(owner.settings["test_geometry"],
                         {"x": moved.x(), "y": moved.y(), "w": 400, "h": 300, "maximized": False})
        again = QDialog(owner)
        apply_dialog_geometry(again, "test_geometry", (1280, 760))
        self.assertEqual(again.geometry(), moved)

        # 화면 밖 위치(모니터를 뺀 경우 등)는 되살리지 않고 기본 위치로
        owner.settings["test_geometry"] = {"x": 99999, "y": 99999, "w": 400, "h": 300}
        third = QDialog(owner)
        apply_dialog_geometry(third, "test_geometry", (1280, 760))
        self.assertTrue(area.contains(third.geometry().center()))

    def test_without_settings_owner(self):
        from PySide6.QtWidgets import QDialog
        from shared.dialog_geometry import apply_dialog_geometry, save_dialog_geometry

        dialog = QDialog()
        apply_dialog_geometry(dialog, "test_geometry", (1024, 768))
        save_dialog_geometry(dialog, "test_geometry")  # 저장할 곳이 없으면 아무 일도 하지 않는다
        self.assertGreater(dialog.width(), 0)

    def test_dialogs_use_helpers(self):
        source = (ROOT / "easy-fulfill.py").read_text(encoding="utf-8")
        self.assertNotIn("self.resize(1280, 760)", source)
        self.assertNotIn("self.resize(1024, 768)", source)
        self.assertEqual(source.count("apply_dialog_geometry(self, "), 2)
        self.assertEqual(source.count("save_dialog_geometry(self, "), 2)


if __name__ == "__main__":
    unittest.main()
