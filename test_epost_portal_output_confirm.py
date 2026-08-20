"""포털 출력여부 읽기 전용 확인의 순수 로직 검증."""

import unittest

from epost_portal_output_confirm import (
    OUTPUT_CONFIRM_DIAGNOSTIC_PATH,
    confirmed_output_regi_nos,
)


class EpostPortalOutputConfirmTests(unittest.TestCase):
    def test_diagnostic_is_saved_only_under_local_output(self):
        self.assertEqual(OUTPUT_CONFIRM_DIAGNOSTIC_PATH.parent.name, "output")

    def test_only_exact_once_printed_rows_are_confirmed(self):
        self.assertEqual(
            confirmed_output_regi_nos(["R1", "R2"], [
                {"regiNo": "R1", "printState": "출력"},
                {"regiNo": "R2", "printState": "미출력"},
                {"regiNo": "OTHER", "printState": "출력"},
            ]),
            ["R1"],
        )
        self.assertEqual(
            confirmed_output_regi_nos(["R1"], [
                {"regiNo": "R1", "printState": "출력"},
                {"regiNo": "R1", "printState": "출력"},
            ]),
            [],
        )


if __name__ == "__main__":
    unittest.main()
