"""우체국 신규출력 읽기 전용 대조의 순수 로직 검증."""

import unittest

from epost_portal_lookup import (
    GRID_DIAGNOSTIC_PATH,
    QUERY_CONTROL_DIAGNOSTIC_PATH,
    TEXT_CONTROL_DIAGNOSTIC_PATH,
    _work_area_prefix,
    grid_row_indexes_from_ids,
    matched_registration_numbers,
    normalize_registration_number,
    pending_candidates_for_date,
    print_target_combo_keys,
    registration_filter_values,
    registration_input_matches,
    selected_print_target_matches,
    selected_date_matches,
    preferred_outer_text_control_id,
    preferred_query_button_id,
    total_count_from_page_text,
    verify_target_rows,
)
from post_parcel_receipt_store import PrintCandidate


def candidate(regi_no: str, received_at: str = "2026-08-19T12:00:00") -> PrintCandidate:
    return PrintCandidate(
        order_no="ORDER-001",
        req_no="REQ-001",
        res_no="RES-001",
        regi_no=regi_no,
        received_at=received_at,
        print_status="PENDING",
    )


class EpostPortalLookupTests(unittest.TestCase):
    def test_query_control_diagnostic_is_local_output_path(self):
        self.assertEqual(QUERY_CONTROL_DIAGNOSTIC_PATH.parent.name, "output")
        self.assertEqual(GRID_DIAGNOSTIC_PATH.parent.name, "output")
        self.assertEqual(TEXT_CONTROL_DIAGNOSTIC_PATH.parent.name, "output")

    def test_work_area_prefix_is_derived_from_current_tab_input(self):
        class _DateInput:
            def get_attribute(self, name):
                self.assertEqual(name, "id")
                return "mainframe_work_winSHIP_4_1_139_form_div_work_Div00_ipbStartDate_calendaredit_input"

            def assertEqual(self, actual, expected):
                if actual != expected:
                    raise AssertionError(actual)

        self.assertEqual(
            _work_area_prefix(_DateInput()),
            "mainframe_work_winSHIP_4_1_139_form_div_work_",
        )

    def test_query_button_prefers_outer_nexacro_component(self):
        self.assertEqual(
            preferred_query_button_id([
                "mainframe_work_Div00_btnInfoLoad",
                "mainframe_work_Div00_btnInfoLoadTextBoxElement",
            ]),
            "mainframe_work_Div00_btnInfoLoad",
        )
        self.assertEqual(
            preferred_query_button_id(["btnOne", "btnTwo"]),
            "",
        )

    def test_dropdown_text_prefers_outer_component_over_text_child(self):
        self.assertEqual(
            preferred_outer_text_control_id([
                "popup_list_item_0",
                "popup_list_item_0TextBoxElement",
            ]),
            "popup_list_item_0",
        )
        self.assertEqual(
            preferred_outer_text_control_id(["popup_item_0", "popup_item_1"]),
            "",
        )

    def test_all_print_target_uses_the_already_identified_combo(self):
        self.assertEqual(print_target_combo_keys("전체"), ("Home", "Enter"))
        self.assertEqual(print_target_combo_keys("신규출력"), ())
        self.assertTrue(selected_print_target_matches(" 전체 ", "전체"))
        self.assertFalse(selected_print_target_matches("신규출력", "전체"))
        self.assertTrue(selected_date_matches("2026-08-19", "2026-08-19"))
        self.assertTrue(selected_date_matches("20260819", "2026-08-19"))
        self.assertFalse(selected_date_matches("2026-08-20", "2026-08-19"))

    def test_registration_filter_uses_digits_for_both_range_inputs(self):
        self.assertEqual(
            registration_filter_values("68901-6459-467"),
            ("689016459467", "689016459467"),
        )
        self.assertTrue(registration_input_matches("68901-6459-467", "689016459467"))
        self.assertFalse(registration_input_matches("689016459468", "689016459467"))
        with self.assertRaises(RuntimeError):
            registration_filter_values("-")

    def test_grid_rows_require_only_expected_unprinted_rows(self):
        self.assertEqual(
            grid_row_indexes_from_ids([
                "prefix_grdList_body_gridrow_3",
                "prefix_grdList_body_gridrow_3_cell_3_3",
                "prefix_grdList_body_gridrow_1",
            ]),
            [1, 3],
        )
        verified = verify_target_rows(["68901-6459-467"], [{
            "rowIndex": 0, "regiNo": "689016459467", "printState": "미출력", "hasCheckbox": True,
        }])
        self.assertTrue(verified["verified"])
        self.assertEqual(normalize_registration_number("68901-6459-467"), "689016459467")

        blocked = verify_target_rows(["689016459467"], [{
            "rowIndex": 0, "regiNo": "689016459467", "printState": "출력", "hasCheckbox": True,
        }, {
            "rowIndex": 1, "regiNo": "689016459468", "printState": "미출력", "hasCheckbox": True,
        }])
        self.assertFalse(blocked["verified"])
        self.assertEqual(blocked["unexpectedRowCount"], 1)

        no_checkbox = verify_target_rows(["689016459467"], [{
            "rowIndex": 0, "regiNo": "689016459467", "printState": "미출력", "hasCheckbox": False,
        }])
        self.assertFalse(no_checkbox["verified"])

    def test_pending_candidates_are_limited_to_lookup_date(self):
        self.assertEqual(
            [item.regi_no for item in pending_candidates_for_date(
                [candidate("689016459467"), candidate("689016459468", "2026-08-18T23:59:59")],
                "2026-08-19",
            )],
            ["689016459467"],
        )

    def test_total_count_and_known_registration_matches_are_extracted(self):
        candidates = [candidate("689016459467"), candidate("689016459468")]
        body_text = "총건수 : 2 등기번호 689016459467 다른 값 16890164594670"
        self.assertEqual(total_count_from_page_text(body_text), 2)
        self.assertEqual(matched_registration_numbers(body_text, candidates), ["689016459467"])


if __name__ == "__main__":
    unittest.main()
