"""프로그램 로그(shared/app_log)와 「도움말 › 프로그램 로그」 창을 검증한다.

- print 출력은 콘솔에 그대로 나가면서 로그에도 남는다(줄을 나눠 써도, 여러 스레드에서 써도 줄 단위로)
- log()로 남긴 줄은 콘솔에 나가지 않는다
- 창은 모인 줄을 시각과 함께 보이고, 열려 있는 동안 새 줄을 붙이며, 「지우기」로 비운다
"""

import io
import os
import sys
import threading
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from shared import app_log  # noqa: E402


def texts():
    return [entry[2] for entry in app_log.entries_since(0)]


class AppLogTests(unittest.TestCase):
    def setUp(self):
        app_log.clear()
        self.console = io.StringIO()
        self.console_err = io.StringIO()
        self.saved = (sys.stdout, sys.stderr)
        sys.stdout, sys.stderr = self.console, self.console_err
        app_log.install()

    def tearDown(self):
        sys.stdout, sys.stderr = self.saved
        app_log.clear()

    def test_print_goes_to_console_and_log(self):
        print("✓ 동기화 완료")
        print("! 경고", file=sys.stderr)
        self.assertEqual(self.console.getvalue(), "✓ 동기화 완료\n")
        self.assertEqual(self.console_err.getvalue(), "! 경고\n")
        entries = app_log.entries_since(0)
        self.assertEqual([(e[2], e[3]) for e in entries], [("✓ 동기화 완료", False), ("! 경고", True)])

    def test_log_only_not_on_console(self):
        app_log.log("초기화 완료")
        self.assertEqual(self.console.getvalue(), "")
        self.assertEqual(texts(), ["초기화 완료"])

    def test_partial_writes_join_into_lines(self):
        print("가", end="")
        print("나")
        sys.stdout.write("다\n라\n")
        self.assertEqual(texts(), ["가나", "다", "라"])

    def test_install_twice_wraps_once(self):
        app_log.install()
        print("한 번")
        self.assertEqual(texts(), ["한 번"])
        self.assertEqual(self.console.getvalue(), "한 번\n")

    def test_threads_do_not_mix_partial_lines(self):
        def worker(name):
            for i in range(50):
                print(f"{name}-", end="")
                print(i)

        threads = [threading.Thread(target=worker, args=(name,)) for name in ("a", "b", "c")]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        lines = texts()
        self.assertEqual(len(lines), 150)
        for name in ("a", "b", "c"):
            self.assertEqual(sorted(int(line.split("-")[1]) for line in lines if line.startswith(f"{name}-")),
                             list(range(50)))

    def test_missing_console_still_logs(self):
        sys.stdout = app_log._TeeStream(None, False)
        print("콘솔 없음")
        self.assertEqual(texts(), ["콘솔 없음"])

    def test_since_and_clear_keep_numbers_increasing(self):
        app_log.log("하나")
        last = app_log.entries_since(0)[-1][0]
        app_log.clear()
        app_log.log("둘")
        self.assertEqual([e[2] for e in app_log.entries_since(last)], ["둘"])


class LogWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        app_log.clear()

    def test_shows_lines_and_appends_new_ones(self):
        from ui.log_window import LogWindow
        app_log.log("프로그램 시작")
        window = LogWindow()
        window.show()
        self.assertRegex(window.text.toPlainText(), r"^\d\d:\d\d:\d\d  프로그램 시작$")
        app_log.log("초기화 완료")
        window.refresh()
        self.assertEqual(window.text.toPlainText().splitlines()[-1][10:], "초기화 완료")
        window.clear_log()
        self.assertEqual(window.text.toPlainText(), "")
        app_log.log("다시")
        window.refresh()
        self.assertEqual(window.text.toPlainText()[10:], "다시")
        window.close()

    def test_error_lines_are_marked(self):
        from ui.log_window import format_entry
        self.assertEqual(format_entry((1, "09:00:00", "Traceback", True)), "09:00:00  [오류] Traceback")
        self.assertEqual(format_entry((2, "09:00:00", "보통", False)), "09:00:00  보통")


if __name__ == "__main__":
    unittest.main()
