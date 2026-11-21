"""새 미답변 문의가 생겼을 때 「문의」 탭 목록 반영을 검증한다.

- 배지 조회가 처음 보는 문의 번호를 찾으면:
  · 다른 탭에 있으면 표시만 해 두고, 「문의」 탭에 들어올 때 기간 「오늘」·「미답변만」으로 바꿔 한 번 새로고침
  · 「문의」 탭을 보고 있으면 지금 기간·필터 그대로 바로 새로고침(쓰던 답변은 그대로)
- 건수가 아니라 번호로 비교(한 건 답변 + 한 건 새로 옴도 잡음), 앱을 켠 뒤 첫 결과는 기준으로만,
  한 종류만 센 결과는 기존 번호에 더하기만, 실패는 무시, 탭을 30일·미답변만으로 불러오면 기준을 맞춤
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]


def run_scenario():
    import importlib.util
    import time

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtWidgets import QApplication

    import sheets.google_sheets_oauth as oauth
    from integrations import naver_commerce

    spec = importlib.util.spec_from_file_location("easy_fulfill_auto_refresh", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = QApplication.instance() or QApplication([])

    def qna(qid, answered=False):
        return naver_commerce.normalize_product_qna({"questionId": qid, "createDate": "2026-11-22T09:00:00.0+09:00",
                                                     "productName": f"상품 {qid}", "question": f"질문 {qid}",
                                                     "answered": answered})

    rows = {"value": [qna(1)]}
    list_calls = []

    def fake_list(days, unanswered_only=True):
        list_calls.append([days, unanswered_only])
        return {"ok": True, "rows": rows["value"], "errors": [], "sample_keys": {}, "fetched_at": "10:00",
                "days": days, "unanswered_only": unanswered_only}

    def wait_idle(window):
        deadline = time.time() + 10
        while time.time() < deadline:
            app.processEvents()
            if window._inquiry_thread is None and not window._inquiry_refresh_pending:
                app.processEvents()
                return
            time.sleep(0.01)
        raise AssertionError("문의 조회가 끝나지 않음")

    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(module, "run_naver_inquiry_list_worker", fake_list), \
            patch.object(module, "run_naver_qna_templates_worker", lambda: {"ok": True, "templates": []}), \
            patch.object(module, "run_naver_inquiry_count_worker",
                         lambda: {"ok": False, "error": "테스트에서는 직접 넣음", "checked_at": ""}):
        window = module.MainWindow()
        window.show()
        app.processEvents()
        combo, check = window.comboBox_inquiry_period, window.checkBox_inquiry_unanswered

        def counted(ids, partial=False):
            window._on_inquiry_badge_counted({"ok": True, "count": len(ids), "partial": partial,
                                              "ids": sorted(ids), "checked_at": "10:00"})
            app.processEvents()

        def go(name):
            window.nav_rail_actions[name].trigger()
            wait_idle(window)

        # 앱을 켠 뒤 첫 결과는 기준만, 탭을 열기 전에 새 번호가 생기면 표시만
        counted({"상품문의:1"})
        result["baseline"] = [window._inquiry_new_pending, len(list_calls)]
        counted({"상품문의:1", "고객문의:2"})
        result["pending_before_first_open"] = [window._inquiry_new_pending, len(list_calls)]
        go("tab_inquiries")
        result["first_open"] = {"calls": list(list_calls), "pending": window._inquiry_new_pending,
                                "status": window.label_inquiry_status.text()}

        # 사용자가 필터를 바꿔 둔 채 다른 탭으로 → 새 문의 → 들어오면 오늘·미답변만으로 한 번
        combo.setCurrentIndex(1)
        wait_idle(window)
        check.setChecked(False)
        wait_idle(window)
        go("tab")
        counted({"상품문의:1", "고객문의:2"})  # 바뀐 것 없음
        result["no_change"] = window._inquiry_new_pending
        counted({"고객문의:2", "상품문의:3"})  # 1 답변 + 3 새로 옴(건수 그대로)
        result["same_count_new_id"] = window._inquiry_new_pending
        before = len(list_calls)
        rows["value"] = [qna(3)]
        go("tab_inquiries")
        result["reenter"] = {"calls": list_calls[before:], "period": combo.currentText(),
                             "unanswered": check.isChecked(), "pending": window._inquiry_new_pending,
                             "status": window.label_inquiry_status.text()}
        go("tab")
        go("tab_inquiries")  # 새 문의 없이 다시 들어오면 조회하지 않는다
        result["reenter_without_new"] = len(list_calls) - before

        # 탭을 보고 있을 때 새 문의 → 지금 필터 그대로 바로, 쓰던 답변은 그대로
        combo.setCurrentIndex(1)
        wait_idle(window)
        window.tableWidget_inquiries.selectRow(0)
        app.processEvents()
        window.plainTextEdit_inquiry_answer.setPlainText("쓰던 답변")
        before = len(list_calls)
        rows["value"] = [qna(3), qna(4)]
        counted({"고객문의:2", "상품문의:3", "상품문의:4"})
        wait_idle(window)
        result["while_shown"] = {
            "calls": list_calls[before:],
            "pending": window._inquiry_new_pending,
            "period": combo.currentText(),
            "draft": window.plainTextEdit_inquiry_answer.toPlainText(),
            "visible": window.card_inquiry_answer.isVisible(),
            "rows": window.tableWidget_inquiries.rowCount(),
            "status": window.label_inquiry_status.text(),
        }

        # 한 종류만 센 결과는 더하기만(빠진 번호를 답변됨으로 보지 않음), 실패는 무시
        go("tab")
        counted({"상품문의:4"}, partial=True)
        counted({"고객문의:2", "상품문의:3", "상품문의:4"})
        result["partial_then_full"] = window._inquiry_new_pending
        window._on_inquiry_badge_counted({"ok": False, "error": "HTTP 500", "checked_at": "10:10"})
        result["failure"] = [window._inquiry_new_pending, sorted(window._inquiry_known_ids)]

        # 탭을 30일·미답변만으로 빠짐없이 불러오면 그 번호로 기준을 맞춘다(배지 조회 전에 본 것은 새 것 아님)
        go("tab_inquiries")
        rows["value"] = [qna(3), qna(4), qna(9)]
        combo.setCurrentIndex(2)
        wait_idle(window)
        go("tab")
        counted({"상품문의:3", "상품문의:4", "상품문의:9"})
        result["list_sets_baseline"] = window._inquiry_new_pending
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


class AutoRefreshTests(unittest.TestCase):
    def test_first_count_is_only_baseline(self):
        self.assertEqual(scenario()["baseline"], [False, 0])

    def test_new_before_first_open_loads_once_on_entry(self):
        r = scenario()
        self.assertEqual(r["pending_before_first_open"], [True, 0])
        first = r["first_open"]
        self.assertEqual(first["calls"], [[0, True]])
        self.assertFalse(first["pending"])
        self.assertTrue(first["status"].endswith("· 새 문의 자동 반영"))

    def test_new_id_with_same_count_is_caught(self):
        r = scenario()
        self.assertFalse(r["no_change"])
        self.assertTrue(r["same_count_new_id"])

    def test_entering_resets_to_today_unanswered(self):
        r = scenario()
        reenter = r["reenter"]
        self.assertEqual(reenter["calls"], [[0, True]])  # 필터를 바꾸면서 따로 조회하지 않고 한 번만
        self.assertEqual([reenter["period"], reenter["unanswered"]], ["오늘", True])
        self.assertFalse(reenter["pending"])
        self.assertIn("새 문의 자동 반영", reenter["status"])
        self.assertEqual(r["reenter_without_new"], 1)  # 새 문의 없이 다시 들어오면 그대로

    def test_refreshes_in_place_when_shown(self):
        shown = scenario()["while_shown"]
        self.assertEqual(shown["calls"], [[7, True]])  # 보고 있는 필터 그대로
        self.assertFalse(shown["pending"])
        self.assertEqual(shown["period"], "최근 7일")
        self.assertEqual(shown["draft"], "쓰던 답변")
        self.assertTrue(shown["visible"])
        self.assertEqual(shown["rows"], 2)
        self.assertIn("새 문의 자동 반영", shown["status"])

    def test_partial_and_failure(self):
        r = scenario()
        self.assertFalse(r["partial_then_full"])
        self.assertEqual(r["failure"], [False, ["고객문의:2", "상품문의:3", "상품문의:4"]])

    def test_list_load_sets_baseline(self):
        self.assertFalse(scenario()["list_sets_baseline"])


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--scenario":
        run_scenario()
    else:
        unittest.main()
