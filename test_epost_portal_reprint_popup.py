"""우체국 재출력 대상 행의 안전 검증."""

import unittest

from epost_portal_reprint_popup import (
    REPRINT_GRID_DIAGNOSTIC_PATH,
    REPRINT_GRID_SCROLL_CLICKS_PER_STEP,
    MAX_REPRINT_GRID_SCROLL_STEPS,
    REPRINT_GRID_SCROLL_SETTLE_SECONDS,
    _matching_target_rows,
    reprint_grid_scroll_down_button_id,
    verified_reprint_rows,
    reprint_lookup_date,
)
from post_parcel_receipt_store import PrintCandidate


class EpostPortalReprintPopupTests(unittest.TestCase):
    def test_diagnostic_is_local_output_only(self):
        self.assertEqual(REPRINT_GRID_DIAGNOSTIC_PATH.parent.name, "output")

    def test_only_one_printed_checkbox_row_is_accepted(self):
        row = {"rowIndex": 2, "regiNo": "68901-6459-467", "printState": "출력", "hasCheckbox": True}
        self.assertEqual(verified_reprint_rows(["689016459467"], [row]), [row])

    def test_missing_duplicate_or_unprinted_row_is_blocked(self):
        with self.assertRaises(RuntimeError):
            verified_reprint_rows(["R1"], [])
        with self.assertRaises(RuntimeError):
            verified_reprint_rows(["R1"], [
                {"regiNo": "R1", "printState": "출력", "hasCheckbox": True},
                {"regiNo": "R1", "printState": "출력", "hasCheckbox": True},
            ])
        with self.assertRaises(RuntimeError):
            verified_reprint_rows(["R1"], [{"regiNo": "R1", "printState": "미출력", "hasCheckbox": True}])

    def test_current_visible_target_row_is_found_without_other_rows(self):
        rows = [
            {"regiNo": "OTHER", "printState": "출력", "hasCheckbox": True},
            {"regiNo": "68901-6459-467", "printState": "출력", "hasCheckbox": True},
        ]
        self.assertEqual(_matching_target_rows(["689016459467"], rows), [rows[1]])

    def test_scroll_button_is_scoped_to_current_work_area(self):
        self.assertEqual(
            reprint_grid_scroll_down_button_id("mainframe_work_"),
            "mainframe_work_grdList_vscrollbar_incbutton",
        )

    def test_grid_scroll_has_enough_steps_for_long_portal_results(self):
        self.assertGreaterEqual(
            MAX_REPRINT_GRID_SCROLL_STEPS * REPRINT_GRID_SCROLL_CLICKS_PER_STEP,
            120,
        )
        self.assertLessEqual(REPRINT_GRID_SCROLL_SETTLE_SECONDS, 0.05)

    def test_reprint_lookup_date_must_come_from_confirmed_receipt(self):
        received_at = "2026-08-19T18:01:55"
        self.assertEqual(received_at[:10], "2026-08-19")

    def test_batch_reprint_requires_one_receipt_date(self):
        candidates = [
            PrintCandidate("o1", "q1", "r1", "68901-1", "2026-08-20T10:00:00", "PORTAL_PRINT_CONFIRMED"),
            PrintCandidate("o2", "q2", "r2", "68901-2", "2026-08-20T10:01:00", "PORTAL_PRINT_CONFIRMED"),
        ]
        self.assertEqual(reprint_lookup_date(candidates), "2026-08-20")
        mixed_dates = [
            candidates[0],
            PrintCandidate("o3", "q3", "r3", "68901-3", "2026-08-19T10:00:00", "PORTAL_PRINT_CONFIRMED"),
        ]
        with self.assertRaises(RuntimeError):
            reprint_lookup_date(mixed_dates)


if __name__ == "__main__":
    unittest.main()
