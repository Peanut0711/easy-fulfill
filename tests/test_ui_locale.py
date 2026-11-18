"""Qt 기본 문구 한국어 번역(shared/ui_locale.py)을 검증한다."""

import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class KoreanTranslationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def test_standard_buttons_are_korean(self):
        from PySide6.QtWidgets import QDialogButtonBox
        from shared.ui_locale import install_korean_translations
        self.assertEqual(install_korean_translations(self.app), ["qtbase_ko"])
        self.assertEqual(install_korean_translations(self.app), ["qtbase_ko"])  # 두 번 불러도 한 번만 설치
        self.assertEqual(len(self.app._easy_fulfill_translators), 1)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Yes | QDialogButtonBox.StandardButton.No
                               | QDialogButtonBox.StandardButton.Cancel)
        texts = [box.button(b).text().replace("&", "") for b in (
            QDialogButtonBox.StandardButton.Yes, QDialogButtonBox.StandardButton.No,
            QDialogButtonBox.StandardButton.Cancel)]
        for text in texts:
            self.assertFalse(text.isascii(), texts)


if __name__ == "__main__":
    unittest.main()
