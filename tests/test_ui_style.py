"""공통 버튼 역할·상태 표시(shared/ui_style.py)를 검증한다."""

import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class UiStyleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def test_colors_come_from_theme_only(self):
        # 색 값은 shared/theme.py 한 곳에서 정한다(UI 개편 2-1부터, 4-1의 「속성 선택자만」 규칙을 대체).
        import re
        source = (ROOT / "shared" / "ui_style.py").read_text(encoding="utf-8")
        self.assertEqual(re.findall(r"#[0-9a-fA-F]{6}\b", source), [])
        from shared import theme, ui_style
        self.assertIs(ui_style.APP_STYLESHEET, theme.APP_STYLESHEET)
        self.assertEqual(ui_style.COLOR_DANGER, theme.DANGER)

    def test_button_role(self):
        from PySide6.QtWidgets import QPushButton
        from shared.ui_style import set_button_role
        button = QPushButton("실행")
        set_button_role(button, "primary")
        self.assertEqual(button.property("role"), "primary")
        set_button_role(button, "danger")
        self.assertEqual(button.property("role"), "danger")
        set_button_role(button, None)
        self.assertEqual(button.property("role"), "")
        with self.assertRaises(ValueError):
            set_button_role(button, "blue")

    def test_status_marks_first_line_and_tooltip(self):
        from PySide6.QtWidgets import QLabel
        from shared.ui_style import clear_status, set_status
        label = QLabel()
        set_status(label, "error", "연결 실패\n상세 원문")
        self.assertEqual(label.text(), "✗ 연결 실패")
        self.assertEqual(label.toolTip(), "연결 실패\n상세 원문")
        self.assertEqual(label.property("status"), "error")
        set_status(label, "ok", "연결됨", "토큰 유효")
        self.assertEqual((label.text(), label.toolTip()), ("✓ 연결됨", "토큰 유효"))
        set_status(label, "ok", "✓ 이미 표시", mark=True)
        self.assertEqual(label.text(), "✓ 이미 표시")
        set_status(label, "idle", "확인 중…")
        self.assertEqual(label.text(), "확인 중…")
        set_status(label, "warn", "두 줄\n그대로", mark=False, single_line=False)
        self.assertEqual((label.text(), label.toolTip()), ("두 줄\n그대로", ""))
        clear_status(label)
        self.assertEqual(label.property("status"), "")

    def test_band_shows_with_margins_and_hides(self):
        from PySide6.QtWidgets import QLabel
        from shared.ui_style import BAND_MARGINS, set_band
        label = QLabel("테스트 모드")
        label.ensurePolished()  # 이미 polish된 라벨에도 여백이 들어가야 한다
        set_band(label, "warn")
        margins = label.contentsMargins()
        self.assertEqual((margins.left(), margins.top(), margins.right(), margins.bottom()), BAND_MARGINS)
        self.assertEqual(label.property("band"), "warn")
        self.assertFalse(label.isHidden())
        set_band(label, None)
        self.assertTrue(label.isHidden())
        self.assertEqual(label.contentsMargins().left(), 0)
        with self.assertRaises(ValueError):
            set_band(label, "error")

    def test_status_color_applied_through_app_stylesheet(self):
        from PySide6.QtGui import QPalette
        from PySide6.QtWidgets import QLabel
        from shared.ui_style import COLOR_DANGER, APP_STYLESHEET, set_status
        previous = self.app.styleSheet()
        self.app.setStyleSheet(APP_STYLESHEET)
        try:
            label = QLabel()
            set_status(label, "error", "실패")
            label.ensurePolished()
            self.assertEqual(label.palette().color(QPalette.ColorRole.WindowText).name(), COLOR_DANGER)
        finally:
            self.app.setStyleSheet(previous)


if __name__ == "__main__":
    unittest.main()
