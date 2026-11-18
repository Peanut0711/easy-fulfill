"""상세페이지 탭 정리(4단계 계획 4-5)를 검증한다.

- 네이버 상품번호 칸이 하나이고 ②(HTML 만들기)와 ③(WING에 채우기)가 모두 그 칸을 읽는다.
- 단계 묶음 제목 ①②③과 끝난 단계의 ✓(입력 초기화·다른 상품으로 바꾸면 ②③을 지움).
- 작업 로그는 기본 접힘이고, 오류가 기록되면 펼친다.
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class DetailPageTabTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from tests.test_order_load_sheets import load_app_module
        cls.app = load_app_module()

    def window(self):
        from PySide6.QtWidgets import QTabWidget, QWidget
        from tests.test_order_load_sheets import main_window_members
        # staticmethod 등 설명자를 그대로 옮기려고 클래스 __dict__(믹스인 포함)에서 꺼낸다.
        members = main_window_members(self.app)
        names = [name for name in members if "detail" in name.lower() and not name.startswith("__")]
        attrs = {name: members[name] for name in names}
        window = type("FakeWindow", (QWidget,), attrs)()
        window.ui = type("Ui", (), {})()
        window.ui.tabWidget = QTabWidget(window)
        window._detail_process = None
        window._detail_process_holding_wing = False
        window._detail_created_product_no = ""
        window._mount_detail_page_tab()
        self.started = []
        window._start_detail_process = lambda title, arguments, hold_wing=False: self.started.append(
            (title, arguments))
        window._open_detail_editor = lambda product_no: None
        self.output = Path(tempfile.mkdtemp(prefix="ef-detail-"))
        window._detail_output_dir = lambda product_no: self.output / product_no
        return window

    def titles(self, window):
        """단계 카드 「제목 · 배지」(디자인 개편 4-5: 끝난 단계는 번호 배지가 ✓)."""
        return [f"{window._detail_step_boxes[step].title_label.text()} · {window._detail_step_boxes[step].step_badge.text()}"
                for step in (1, 2, 3)]

    def finish(self, window, title, exit_code=0, output=""):
        window._detail_process_title = title
        window._detail_process_output = output
        window._on_detail_process_finished(exit_code, None)

    def test_single_product_field_used_by_create_and_apply(self):
        from PySide6.QtWidgets import QLineEdit
        window = self.window()
        page = window.ui.tabWidget.widget(0)
        edits = {edit.placeholderText() for edit in page.findChildren(QLineEdit)}
        self.assertEqual(edits, {"네이버 스마트스토어 상품번호 또는 상품 주소", "쿠팡 등록상품 ID"})
        window.lineEdit_detail_naver.setText("https://smartstore.naver.com/shop/products/12345?x=1")
        self.assertEqual(window.lineEdit_detail_naver.text(), "12345")  # 주소를 번호로 바꾼다
        window.pushButton_detail_create.click()
        (self.output / "12345").mkdir()
        (self.output / "12345" / "coupang-paste.html").write_text("<p>x</p>", encoding="utf-8")
        window.lineEdit_detail_vendor_inventory.setText("987")
        window.pushButton_detail_stage.click()
        self.assertEqual([args[-2:] if title.startswith("기존") else args[-2:-1] for title, args in self.started],
                         [["12345"], ["12345", "987"]])

    def test_step_titles_and_done_marks(self):
        window = self.window()
        self.assertEqual(self.titles(window),
                         ["쿠팡 WING 연결 · 1", "상품 상세 HTML 만들기 · 2", "쿠팡 상품 HTML 적용 · 3"])
        with patch.object(self.app.QTimer, "singleShot"):
            self.finish(window, "쿠팡 로그인 연결", output='{"seller": ""}')
            window._detail_created_product_no = "12345"
            window.lineEdit_detail_naver.setText("12345")
            self.finish(window, "상품 상세 HTML 생성")
        self.assertEqual(self.titles(window)[:2], ["쿠팡 WING 연결 · ✓", "상품 상세 HTML 만들기 · ✓"])

        class FakeProcess:
            def readAllStandardOutput(self):
                return b"\xec\xa4\x80\xeb\xb9\x84 \xec\x99\x84\xeb\xa3\x8c: \xec\xbf\xa0\xed\x8c\xa1 \xec\x83\x81\xed\x92\x88 987"

        window._detail_process = FakeProcess()  # 「준비 완료: 쿠팡 상품 987」
        window._detail_process_title = "기존 쿠팡 상품 HTML 채우기"
        window._detail_process_output = ""
        window._on_detail_process_output()
        self.assertTrue(self.titles(window)[2].endswith("✓"))

        # 같은 상품번호를 주소로 다시 붙여 넣어도 ✓는 그대로, 다른 상품이면 ②③을 지운다.
        window.lineEdit_detail_naver.setText("https://smartstore.naver.com/shop/products/12345")
        self.assertTrue(self.titles(window)[1].endswith("✓"))
        window.lineEdit_detail_naver.setText("55555")
        self.assertEqual(self.titles(window), ["쿠팡 WING 연결 · ✓", "상품 상세 HTML 만들기 · 2", "쿠팡 상품 HTML 적용 · 3"])

        window._set_detail_step_done(2)
        window.pushButton_detail_clear_inputs.click()
        self.assertEqual(self.titles(window)[1:], ["상품 상세 HTML 만들기 · 2", "쿠팡 상품 HTML 적용 · 3"])
        self.assertTrue(self.titles(window)[0].endswith("✓"))  # 로그인 연결은 입력이 아니다
        self.finish(window, "쿠팡 로그인 세션 삭제")
        self.assertEqual(self.titles(window)[0], "쿠팡 WING 연결 · 1")

    def test_log_collapsed_by_default_and_opens_on_error(self):
        window = self.window()
        log, toggle = window.textEdit_detail_log, window.toolButton_detail_log_toggle
        self.assertTrue(log.isHidden())
        self.assertEqual(toggle.text(), "작업 로그 ▸")
        toggle.click()
        self.assertFalse(log.isHidden())
        self.assertEqual(toggle.text(), "작업 로그 ▾")
        toggle.click()
        window._detail_log("[상품 상세 HTML 생성] 시작")
        self.assertTrue(log.isHidden())  # 보통 기록은 펼치지 않는다
        window._detail_log("[오류] WING 화면을 찾지 못했습니다.")
        self.assertFalse(log.isHidden())
        self.assertTrue(toggle.isChecked())
        toggle.click()
        with patch.object(self.app.QTimer, "singleShot"):
            self.finish(window, "상품 상세 HTML 생성", exit_code=1)
        self.assertFalse(log.isHidden())  # 실패로 끝나면 펼친다
        self.assertIn("실패 (1)", log.toPlainText())


if __name__ == "__main__":
    unittest.main()
