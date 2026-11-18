"""앱 테마(shared/theme.py)의 기준값·스타일시트·적용을 검증한다."""

import os
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from shared import theme  # noqa: E402


class ContrastTests(unittest.TestCase):
    """docs/ui-design-guide.md 「검증 규칙」: 본문 글자 대비 4.5:1 이상."""

    def assertReadable(self, fg, bg, minimum=4.5):
        ratio = theme.contrast_ratio(fg, bg)
        self.assertGreaterEqual(ratio, minimum, f"{fg} / {bg} 대비 {ratio:.2f}")

    def test_text_on_backgrounds(self):
        for fg in (theme.TEXT, theme.TEXT_SECONDARY, theme.TEXT_MUTED):
            for bg in (theme.BG_APP, theme.BG_SURFACE, theme.BG_SUBTLE):
                with self.subTest(fg=fg, bg=bg):
                    self.assertReadable(fg, bg)

    def test_status_colors_on_their_backgrounds(self):
        for fg, bg in ((theme.OK, theme.OK_SUBTLE), (theme.WARN, theme.WARN_SUBTLE),
                       (theme.DANGER, theme.DANGER_SUBTLE), (theme.OK, theme.BG_SURFACE),
                       (theme.WARN, theme.BG_SURFACE), (theme.DANGER, theme.BG_SURFACE)):
            with self.subTest(fg=fg, bg=bg):
                self.assertReadable(fg, bg)

    def test_button_text(self):
        self.assertReadable("#ffffff", theme.PRIMARY)
        self.assertReadable(theme.DANGER, theme.BG_SURFACE)
        self.assertReadable(theme.TEXT, theme.PRIMARY_SUBTLE)  # 선택 행
        self.assertReadable(theme.TOOLTIP_TEXT, theme.TOOLTIP_BG)

    def test_contrast_ratio_reference_values(self):
        self.assertAlmostEqual(theme.contrast_ratio("#000000", "#ffffff"), 21.0, places=2)
        self.assertAlmostEqual(theme.contrast_ratio("#ffffff", "#ffffff"), 1.0, places=2)


class StylesheetTests(unittest.TestCase):
    def test_no_unformatted_placeholders(self):
        self.assertNotRegex(theme.APP_STYLESHEET, r"\{[A-Z_]+\}")
        self.assertEqual(theme.APP_STYLESHEET.count("{"), theme.APP_STYLESHEET.count("}"))

    def test_icons_exist(self):
        for url in re.findall(r"url\(([^)]+)\)", theme.APP_STYLESHEET):
            with self.subTest(url=url):
                self.assertTrue(Path(url).is_file(), url)

    def test_colors_in_stylesheet_are_theme_values(self):
        values = {v.lower() for k, v in vars(theme).items()
                  if k.isupper() and isinstance(v, str) and v.startswith("#")}
        extra = {"#ffffff", "#f9d6d2", "#d9a29b", "#f0d4d0"}  # 흰색·위험 버튼 눌림/꺼짐
        for color in re.findall(r"#[0-9a-fA-F]{6}\b", theme.APP_STYLESHEET):
            with self.subTest(color=color):
                self.assertIn(color.lower(), values | extra)


class ApplyThemeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def test_apply_theme_sets_fusion_palette_font_and_stylesheet(self):
        from PySide6.QtGui import QPalette
        previous = (self.app.style().name(), self.app.palette(), self.app.font(), self.app.styleSheet())
        try:
            theme.apply_theme(self.app)
            self.assertEqual(self.app.styleSheet(), theme.APP_STYLESHEET)
            self.app.setStyleSheet("")  # 스타일시트가 있으면 style()이 감싼 객체를 돌려준다
            self.assertEqual(self.app.style().metaObject().className(), "QFusionStyle")
            self.app.setStyleSheet(theme.APP_STYLESHEET)
            self.assertEqual(self.app.palette().color(QPalette.ColorRole.Highlight).name(), theme.PRIMARY)
            self.assertEqual(self.app.palette().color(QPalette.ColorRole.Window).name(), theme.BG_APP)
            self.assertEqual(self.app.font().pointSizeF(), theme.FONT_SIZE)
            self.assertEqual(self.app.styleSheet(), theme.APP_STYLESHEET)
        finally:
            if previous[0]:
                self.app.setStyle(previous[0])
            self.app.setPalette(previous[1])
            self.app.setFont(previous[2])
            self.app.setStyleSheet(previous[3])


if __name__ == "__main__":
    unittest.main()
