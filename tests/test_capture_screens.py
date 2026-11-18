"""화면 캡처 도구(tools/capture_screens.py)의 배율·창 크기 계산과 전후 비교를 검증한다.

실제 캡처(Qt 창 생성)는 느리고 Windows 글꼴에 따라 달라 여기서는 하지 않는다.
"""

import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

import capture_screens as cs  # noqa: E402


def _png(path, size=(20, 10), color="white", dot=None):
    path.parent.mkdir(parents=True, exist_ok=True)
    image = Image.new("RGB", size, color)
    if dot is not None:
        image.putpixel(dot, (0, 0, 0))
    image.save(path)
    return path


class ScaleTests(unittest.TestCase):
    def test_parse_scales(self):
        self.assertEqual(cs.parse_scales("100,125,150"), (100, 125, 150))
        self.assertEqual(cs.parse_scales("125"), (125,))

    def test_parse_scales_rejects_unknown_or_empty(self):
        for text in ("175", "", "100,90"):
            with self.subTest(text=text), self.assertRaises(ValueError):
                cs.parse_scales(text)

    def test_window_size_follows_app_default_per_scale(self):
        minimum = (667, 443)
        self.assertEqual(cs.capture_window_size(100, minimum), (900, 800))
        self.assertEqual(cs.capture_window_size(125, minimum), (721, 638))
        # 150% FHD 작업 영역의 기본 폭(601)은 최소 폭보다 작아 최소 폭을 쓴다
        self.assertEqual(cs.capture_window_size(150, minimum), (667, 526))

    def test_safe_name(self):
        self.assertEqual(cs.safe_name("주문·발송(&O)"), "주문_발송__O")
        self.assertEqual(cs.safe_name("..."), "unnamed")


class CompareTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_same_images(self):
        a = _png(self.tmp / "a.png")
        b = _png(self.tmp / "b.png")
        self.assertEqual(cs.compare_images(a, b)["status"], "same")

    def test_changed_pixel_reports_count_bbox_and_diff_image(self):
        a = _png(self.tmp / "a.png")
        b = _png(self.tmp / "b.png", dot=(3, 4))
        diff = self.tmp / "diff.png"
        result = cs.compare_images(a, b, diff)
        self.assertEqual(result["status"], "changed")
        self.assertEqual(result["changed_pixels"], 1)
        self.assertEqual(result["bbox"], [3, 4, 4, 5])
        self.assertTrue(diff.exists())

    def test_size_change(self):
        a = _png(self.tmp / "a.png", size=(20, 10))
        b = _png(self.tmp / "b.png", size=(21, 10))
        result = cs.compare_images(a, b)
        self.assertEqual(result["status"], "size")
        self.assertEqual((result["before_size"], result["after_size"]), ([20, 10], [21, 10]))

    def test_compare_dirs_matches_relative_paths(self):
        before, after, out = self.tmp / "before", self.tmp / "after", self.tmp / "out"
        _png(before / "100" / "tab-1.png")
        _png(after / "100" / "tab-1.png", dot=(0, 0))
        _png(before / "100" / "old.png")
        _png(after / "125" / "new.png")
        rows = {row["path"]: row["status"] for row in cs.compare_dirs(before, after, out)}
        self.assertEqual(rows, {"100/tab-1.png": "changed", "100/old.png": "removed", "125/new.png": "added"})
        self.assertTrue((out / "diff" / "100" / "tab-1.png").exists())
        cs.write_compare_html(cs.compare_dirs(before, after, out), before, after, out)
        self.assertIn("100/tab-1.png", (out / "index.html").read_text(encoding="utf-8"))

    def test_fail_on_diff_exit_code(self):
        before, after = self.tmp / "before", self.tmp / "after"
        _png(before / "100" / "tab-1.png")
        _png(after / "100" / "tab-1.png")
        out = str(self.tmp / "out")
        self.assertEqual(cs.main(["--compare", str(before), str(after), "--out", out, "--fail-on-diff"]), 0)
        _png(after / "100" / "tab-1.png", dot=(1, 1))
        self.assertEqual(cs.main(["--compare", str(before), str(after), "--out", out, "--fail-on-diff"]), 1)


if __name__ == "__main__":
    unittest.main()
