"""우체국 계약소포 테스트 접수 요청값 변환 검증."""

import unittest
import xml.etree.ElementTree as ET

from post_parcel import (
    ParcelValidationError,
    address_suggestion,
    build_real_order_values,
    build_test_order_values,
    normalize_content_code,
    _postcode_response_items,
)


SETTINGS = {
    "epost_parcel_cust_no": "1234567890",
    "epost_parcel_appr_no": "1234567890",
    "epost_parcel_pay_type": "1",
    "epost_parcel_office_ser": "250428756",
    "epost_parcel_default_weight_kg": "2",
    "epost_parcel_default_volume_cm": "60",
    "epost_parcel_micro_yn": "N",
    "epost_parcel_content_code": "29",
    "epost_parcel_print_yn": "N",
}

ROW = {
    "주문번호": "",
    "고객주문처명": "",
    "수취인명": "테스트 수취인",
    "우편번호": "12345",
    "수취인 주소": "서울시 테스트로 1",
    "수취인 전화번호": "02-1234-5678",
    "수취인 이동통신": "010-1234-5678",
    "상품명": "테스트 상품",
}


class PostParcelRequestTests(unittest.TestCase):
    def test_content_code_is_normalized_to_three_digits(self):
        self.assertEqual(normalize_content_code("29"), "029")

    def test_blank_order_number_is_generated_and_blank_source_is_kept(self):
        values = build_test_order_values(ROW, SETTINGS)

        self.assertTrue(values["orderNo"].startswith("EFTEST-"))
        self.assertEqual(values["ordCompNm"], "")
        self.assertEqual(values["recAddr2"], "-")
        self.assertEqual(values["contCd"], "029")

    def test_real_receipt_generates_distinct_order_number_and_disables_self_print(self):
        values = build_real_order_values(ROW, SETTINGS)

        self.assertTrue(values["orderNo"].startswith("EFREAL-"))
        self.assertEqual(values["testYn"], "N")
        self.assertEqual(values["printYn"], "N")

    def test_real_receipt_blocks_self_print_setting(self):
        with self.assertRaisesRegex(ParcelValidationError, "운송장 자체 출력"):
            build_real_order_values(ROW, {**SETTINGS, "epost_parcel_print_yn": "Y"})

    def test_missing_recipient_address_is_blocked_before_api_call(self):
        row = {**ROW, "수취인 주소": ""}

        with self.assertRaisesRegex(ParcelValidationError, "수취인 주소"):
            build_test_order_values(row, SETTINGS)

    def test_phone_only_row_is_mapped_to_both_contact_fields(self):
        row = {**ROW, "수취인 이동통신": ""}

        values = build_test_order_values(row, SETTINGS)
        self.assertEqual(values["recTel"], "0212345678")
        self.assertEqual(values["recMob"], "0212345678")

    def test_standard_address_and_original_tail_are_split(self):
        suggestion = address_suggestion(
            "충남 천안시 서북구 성거읍 정자1길 10 생산기술팀",
            {
                "postcd": "31045",
                "address": "충청남도 천안시 서북구 성거읍 정자1길 10 (오색당리, 세라젬)",
            },
        )

        self.assertEqual(suggestion.postcode, "31045")
        self.assertEqual(suggestion.address1, "충청남도 천안시 서북구 성거읍 정자1길 10")
        self.assertEqual(suggestion.address2, "(오색당리, 세라젬) 생산기술팀")

    def test_postcode_not_found_response_allows_next_query(self):
        root = ET.fromstring(
            "<error><error_code>ERR-125</error_code><message>결과 없음</message></error>",
        )

        self.assertEqual(_postcode_response_items(root), [])


if __name__ == "__main__":
    unittest.main()
