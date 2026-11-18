"""메인 창과 탭의 최소 크기가 창 크기 예산 안에 드는지 검사한다.

예산은 docs/gui-window-settings-plan.md 「크기 기준」을 따른다. 글꼴에 따라 값이 달라서
Windows(맑은 고딕)에서만 잰다. 창은 화면에 띄우지 않는다.
"""

import json
import os
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tools.measure_window_sizes import PAGE_MIN_BUDGET, WINDOW_MIN_BUDGET  # noqa: E402

# 아직 예산을 넘는 탭(B단계 「설정」 탭으로 옮기면 없어진다)
KNOWN_OVER_BUDGET = {"tab_admin"}

_cached = None


def measured():
    global _cached
    if _cached is None:
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        env.pop("QT_QPA_PLATFORM", None)
        completed = subprocess.run(
            [sys.executable, str(ROOT / "tools" / "measure_window_sizes.py"), "--json"],
            capture_output=True, text=True, encoding="utf-8", env=env, timeout=180,
        )
        line = next((l for l in completed.stdout.splitlines() if l.startswith("RESULT ")), None)
        if line is None:
            raise AssertionError(completed.stdout[-2000:] + completed.stderr[-2000:])
        _cached = json.loads(line[len("RESULT "):])
    return _cached


@unittest.skipUnless(sys.platform == "win32", "맑은 고딕 기준 측정은 Windows에서만 한다")
class WindowSizeBudgetTests(unittest.TestCase):
    def test_each_tab_fits_page_budget(self):
        over = {}
        for tab in measured()["tabs"]:
            width, height = tab["min"]
            if (width > PAGE_MIN_BUDGET[0] or height > PAGE_MIN_BUDGET[1]) and tab["name"] not in KNOWN_OVER_BUDGET:
                over[tab["name"]] = tab["min"]
        self.assertEqual(over, {}, f"탭 최소 크기 예산 {PAGE_MIN_BUDGET} 초과")

    def test_known_over_budget_tabs_still_exist(self):
        # 예외 목록이 실제로 남은 탭만 가리키게 한다(옮긴 뒤에는 목록에서 지운다).
        names = {tab["name"] for tab in measured()["tabs"]}
        self.assertLessEqual(KNOWN_OVER_BUDGET, names)

    @unittest.skipIf(KNOWN_OVER_BUDGET, "예산을 넘는 탭이 남아 있는 동안은 창 최소 크기도 넘는다")
    def test_window_minimum_fits_budget(self):
        width, height = measured()["window_min"]
        self.assertLessEqual(width, WINDOW_MIN_BUDGET[0])
        self.assertLessEqual(height, WINDOW_MIN_BUDGET[1])


if __name__ == "__main__":
    unittest.main()
