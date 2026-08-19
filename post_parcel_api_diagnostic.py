"""우체국 계약소포 OpenAPI 조회 결과를 비밀값 없이 콘솔에 표시합니다.

사용 예:
    python post_parcel_api_diagnostic.py cust
    python post_parcel_api_diagnostic.py appr
    python post_parcel_api_diagnostic.py office

`설정` 시트의 epost_parcel_* 값을 읽습니다. 소포 신청(InsertOrder),
취소, 운송장 출력은 호출하지 않습니다.
"""

from __future__ import annotations

import argparse
import sys
import xml.etree.ElementTree as ET

import requests

from google_sheets_oauth import get_authorized_gspread_client
from post_parcel import call_api, element_text, required_setting
from tracking.repository import open_config_worksheet, read_config_values_map
from tracking.workers import SPREADSHEET_ID, CONFIG_SHEET_HEADERS, CONFIG_SHEET_TITLE


CONFIG_KEY_API = "epost_parcel_api_key"
CONFIG_KEY_SECURITY = "epost_parcel_security_key"
CONFIG_KEY_MEMBER_ID = "epost_parcel_member_id"
CONFIG_KEY_CUST_NO = "epost_parcel_cust_no"


def read_settings() -> dict[str, str]:
    """공유 설정 시트에서 필요한 값만 읽습니다."""
    client = get_authorized_gspread_client()
    worksheet = open_config_worksheet(
        client, SPREADSHEET_ID, CONFIG_SHEET_TITLE, CONFIG_SHEET_HEADERS,
    )
    return read_config_values_map(worksheet)


def print_result(api_name: str, root: ET.Element) -> bool:
    error_code = element_text(root.find(".//error_code"))
    message = element_text(root.find(".//message"))
    print(f"[우체국 계약소포 OpenAPI] {api_name}")
    if error_code:
        print("결과: 실패")
        print(f"오류코드: {error_code}")
        print(f"오류내용: {message or '우체국 API 오류 응답'}")
        return False

    print("결과: 성공")
    print("오류코드: 없음")
    return True


def main() -> int:
    parser = argparse.ArgumentParser(
        description="공유 설정을 사용해 우체국 계약소포 조회 API만 호출합니다.",
    )
    parser.add_argument(
        "query", nargs="?", default="appr", choices=("cust", "appr", "office"),
        help="cust=고객번호, appr=계약 승인번호(기본), office=공급지 조회",
    )
    args = parser.parse_args()

    try:
        settings = read_settings()
        api_key = required_setting(settings, CONFIG_KEY_API, "소포신청 인증키")
        security_key = required_setting(settings, CONFIG_KEY_SECURITY, "접수용 보안키")

        if args.query == "cust":
            member_id = required_setting(settings, CONFIG_KEY_MEMBER_ID, "인터넷우체국 ID")
            root = call_api("GetCustNo", f"memberID={member_id}", api_key, security_key)
            return 0 if print_result("고객번호 조회(GetCustNo)", root) else 2

        customer_number = required_setting(settings, CONFIG_KEY_CUST_NO, "고객번호")
        if not customer_number.isdigit() or len(customer_number) != 10:
            raise ValueError("고객번호는 대시 없이 10자리 숫자로 입력해야 합니다.")

        api_name = "GetApprNo" if args.query == "appr" else "GetOfficeInfo"
        title = "계약 승인번호 조회(GetApprNo)" if args.query == "appr" else "공급지 조회(GetOfficeInfo)"
        root = call_api(api_name, f"custNo={customer_number}", api_key, security_key)
        return 0 if print_result(title, root) else 2
    except requests.RequestException as error:
        print("[우체국 계약소포 OpenAPI] 통신 실패")
        print(f"오류내용: {error}")
        return 1
    except (ET.ParseError, ValueError) as error:
        print("[우체국 계약소포 OpenAPI] 실행 실패")
        print(f"오류내용: {error}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
