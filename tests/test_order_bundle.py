"""옵션명의 묶음 개수(「3개」 옵션)를 읽어 포장 수량을 계산하는 규칙을 검증한다."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from orders.bundle import bundle_note, bundle_size, packed_quantity, quantity_markdown  # noqa: E402


class BundleSizeTests(unittest.TestCase):
    def test_reads_single_bundle_count(self):
        # 실제 사례: 쿠팡 「3개 ESP32-C3(핀헤더 미납땜)」 옵션 1세트 구매
        self.assertEqual(bundle_size("3개 ESP32-C3(핀헤더 미납땜)"), 3)
        self.assertEqual(bundle_size("개수: 5 개 / 타입: C3"), 5)
        self.assertEqual(bundle_size("10개입"), 10)
        self.assertEqual(bundle_size("3개, 3개 세트"), 3)  # 같은 개수가 반복되면 그대로

    def test_no_bundle_returns_one(self):
        for option in ("", None, "없음", "N16R8", "1개", "WiFi", "0개"):
            with self.subTest(option=option):
                self.assertEqual(bundle_size(option), 1)

    def test_ambiguous_or_non_count_returns_one(self):
        self.assertEqual(bundle_size("5개 (+1개 증정)"), 1)  # 서로 다른 개수가 섞이면 곱하지 않는다
        self.assertEqual(bundle_size("12개월 보증"), 1)
        self.assertEqual(bundle_size("1.5개"), 1)
        self.assertEqual(bundle_size("1,000개"), 1)


class QuantityMarkdownTests(unittest.TestCase):
    def test_bundle_option_shows_pieces_and_formula(self):
        # 쿠팡(bundles=True)만 묶음 개수를 곱한다
        self.assertEqual(quantity_markdown(1, "3개 ESP32-C3", bundles=True), "**[ 3 개 ]** (1세트×3개)")
        self.assertEqual(quantity_markdown(2, "3개 ESP32-C3", bundles=True), "**[ 6 개 ]** (2세트×3개)")
        self.assertEqual(packed_quantity(2, "3개"), 6)
        self.assertEqual(bundle_note(2, "3개"), "2세트×3개")

    def test_plain_option_keeps_previous_format(self):
        self.assertEqual(quantity_markdown(2, "N16R8", bundles=True), "**[ 2 개 ]**")
        self.assertEqual(quantity_markdown(1, "", bundles=True), "**[ 1 개 ]**")
        self.assertEqual(bundle_note(1, "없음"), "")

    def test_default_shows_order_quantity_only(self):
        # 네이버·지마켓·11번가(기본값): 옵션명에 「N개」가 있어도 주문 수량만 적는다
        self.assertEqual(quantity_markdown(12, "회전각:360도-12개"), "**[ 12 개 ]**")
        self.assertEqual(quantity_markdown(1, "3개"), "**[ 1 개 ]**")


if __name__ == "__main__":
    unittest.main()
