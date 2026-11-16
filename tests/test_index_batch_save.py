"""주문 파일 처리 중 인덱스 JSON을 주문마다 다시 쓰지 않고 끝날 때 한 번 저장하는지 검증한다."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run_scenario(name):
    import importlib.util
    import time
    from unittest.mock import patch

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, str(ROOT))
    from PySide6.QtWidgets import QApplication

    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_index_batch", ROOT / "easy-fulfill.py")
    app_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app_module)
    app = QApplication.instance() or QApplication([])
    result = {}
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")):
        window = app_module.MainWindow()
        window._startup_order_index_sync_started = True
        writes = []
        original = window._persist_index_values_to_json
        window._persist_index_values_to_json = lambda: (writes.append(window.current_idx_naver), original())[1]
        start = window.current_idx_naver
        started = time.perf_counter()
        with window._batched_index_saves():
            for _ in range(50):
                window.update_naver_index()
            window.update_11st_index()
        result["batch_seconds"] = time.perf_counter() - started
        result["writes"] = writes
        result["final"] = window.current_idx_naver
        result["start"] = start
        result["push_scheduled"] = window._index_sheet_push_timer.isActive()
        result["reason"] = getattr(window, "_pending_index_reason", "")
        saved = json.loads(Path(window.index_file_path).read_text(encoding="utf-8"))
        result["saved_naver"] = next(iter(saved.values()))["naver"]
        window._index_sheet_push_timer.stop()
        writes.clear()
        with window._batched_index_saves():
            window.update_11st_index()
        result["eleventh_only_push"] = window._index_sheet_push_timer.isActive()
        result["eleventh_only_writes"] = len(writes)
    print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
    os._exit(0)


class IndexBatchSaveTests(unittest.TestCase):
    def test_batch_writes_index_json_once_with_final_value(self):
        with tempfile.TemporaryDirectory() as directory:
            env = dict(os.environ, QT_QPA_PLATFORM="offscreen", PYTHONIOENCODING="utf-8")
            completed = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), "--scenario", "batch"],
                cwd=directory, capture_output=True, text=True, encoding="utf-8", env=env, timeout=120,
            )
        line = next((line for line in completed.stdout.splitlines() if line.startswith("RESULT ")), None)
        self.assertIsNotNone(line, completed.stdout[-2000:] + completed.stderr[-2000:])
        result = json.loads(line[len("RESULT "):])
        self.assertEqual(result["final"], result["start"] + 50)
        self.assertEqual(result["writes"], [result["final"]])  # 50건 처리에도 저장은 끝날 때 한 번
        self.assertEqual(result["saved_naver"], result["final"])
        self.assertTrue(result["push_scheduled"])
        self.assertEqual(result["reason"], "주문 로드")
        self.assertFalse(result["eleventh_only_push"])  # 11번가만 바뀌면 시트 반영 없음
        self.assertEqual(result["eleventh_only_writes"], 1)


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--scenario":
        run_scenario(sys.argv[2])
    else:
        unittest.main()
