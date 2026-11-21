"""화면 코드에 색 값·위젯 스타일시트·글꼴을 직접 넣지 않는지 검사한다(디자인 개편 5-1 재발 방지).

색·간격·글꼴은 shared/theme.py 한 곳에서 정하고, 화면 코드는 테마 이름이나 동적 속성을 쓴다
(docs/ui-design-guide.md). 아래 허용 목록은 화면 모양이 아니라 「데이터」인 색이다. 새로 직접 넣어야 하면
테마에 이름을 만들고, 정말 데이터라면 이유와 함께 허용 목록에 줄 내용을 더한다.
"""

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# 화면 코드(테마 파일 제외). 업무 로직 모듈(orders/, product_detail/ 등)이 만드는 HTML·엑셀 색은 대상이 아니다.
UI_FILES = sorted({
    ROOT / "easy-fulfill.py",
    *(ROOT / "ui").rglob("*.py"),
    *(path for path in (ROOT / "shared").glob("*.py") if path.name != "theme.py"),
    ROOT / "tracking" / "controller.py",
    ROOT / "documents" / "quotation_statement_widget.py",
})

PATTERN = re.compile(r"#[0-9a-fA-F]{6}\b|setStyleSheet\(|QFont\(")

# (파일 이름, 줄에 들어 있는 글자) → 이유
ALLOWED = {
    ("easy-fulfill.py", '("검정", "#222222")'): "상세 편집기 글자색 견본(문서에 넣는 색)",
    ("easy-fulfill.py", '("주황", "#E65100")'): "상세 편집기 글자색 견본(문서에 넣는 색)",
    ("easy-fulfill.py", "chip.setStyleSheet("): "글자색 견본 버튼마다 다른 바탕색(데이터), 테두리는 테마",
    ("easy-fulfill.py", 'QColor("#222222")'): "색 선택 창 기본값 = 첫 견본 색",
    ("easy-fulfill.py", "[data-ef-spacer] { position:relative; outline:1px dashed"): "웹 미리보기 안에 넣는 CSS",
    ("easy-fulfill.py", "outline-offset:-1px !important; background:"): "웹 미리보기 안에 넣는 CSS",
    ("easy-fulfill.py", "color:#475569; pointer-events:none;"): "웹 미리보기 안에 넣는 CSS",
    ("easy-fulfill.py", "border: '2px solid #1976d2',"): "웹 미리보기 안에서 실행하는 JS",
    ("order_ship_panel.py", "font = QFont()"): "QPainter로 스토어 글자 배지를 그리는 글꼴(위젯 스타일 아님)",
    ("chrome.py", "font = QFont()"): "QPainter로 측면 툴바 숫자 배지를 그리는 글꼴(위젯 스타일 아님)",
}


def _offending_lines():
    found = []
    for path in UI_FILES:
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if PATTERN.search(line):
                found.append((path.name, number, line.strip()))
    return found


class StyleGuardTests(unittest.TestCase):
    def test_no_direct_colors_styles_or_fonts_outside_theme(self):
        unexpected = [f"{name}:{number}: {line}" for name, number, line in _offending_lines()
                      if not any(name == file and snippet in line for (file, snippet) in ALLOWED)]
        self.assertEqual(unexpected, [], "색·스타일시트·글꼴은 shared/theme.py의 이름이나 동적 속성으로 쓴다")

    def test_allowlist_entries_still_used(self):
        lines = _offending_lines()
        stale = [key for key in ALLOWED
                 if not any(name == key[0] and key[1] in line for name, _number, line in lines)]
        self.assertEqual(stale, [], "코드에서 사라진 허용 항목은 목록에서 지운다")

    def test_ui_file_has_no_inline_styles_or_font_sizes(self):
        text = (ROOT / "ui" / "main_window.ui").read_text(encoding="utf-8")
        self.assertNotIn('name="styleSheet"', text)
        self.assertNotIn("<pointsize>", text)


if __name__ == "__main__":
    unittest.main()
