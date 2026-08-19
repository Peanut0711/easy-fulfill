"""우체국 계약소포 OpenAPI의 단건 접수 공통 기능.

테스트 접수와 명시적으로 확인한 실접수만 제공한다. 운송장 출력과 네이버
발송은 여기서 수행하지 않으며, 인증키와 수취인 개인정보를 로그에 출력하지
않는다.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import re
import secrets
from typing import Mapping, Sequence
import xml.etree.ElementTree as ET

import requests
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes


API_BASE_URL = "http://ship.epost.go.kr/api.{}.jparcel"
POSTCODE_API_URL = "http://biz.epost.go.kr/KpostPortal/openapi2"
USER_AGENT = "easy-fulfill-contract-parcel/1.0"

CONFIG_KEY_API = "epost_parcel_api_key"
CONFIG_KEY_SECURITY = "epost_parcel_security_key"
CONFIG_KEY_CUST_NO = "epost_parcel_cust_no"
CONFIG_KEY_APPR_NO = "epost_parcel_appr_no"
CONFIG_KEY_PAY_TYPE = "epost_parcel_pay_type"
CONFIG_KEY_OFFICE_SER = "epost_parcel_office_ser"
CONFIG_KEY_WEIGHT = "epost_parcel_default_weight_kg"
CONFIG_KEY_VOLUME = "epost_parcel_default_volume_cm"
CONFIG_KEY_MICRO_YN = "epost_parcel_micro_yn"
CONFIG_KEY_CONTENT_CODE = "epost_parcel_content_code"
CONFIG_KEY_PRINT_YN = "epost_parcel_print_yn"
CONFIG_KEY_POSTCODE_API = "epost_postcode_api_key"


class ParcelValidationError(ValueError):
    """우체국 API 호출 전에 발견한 송장 행 또는 설정 오류."""


class ParcelApiError(RuntimeError):
    """우체국 API가 반환한 오류."""

    def __init__(self, code: str, message: str):
        self.code = code
        self.message = message
        super().__init__(f"{code}: {message}" if code else message)


@dataclass(frozen=True)
class ParcelReceipt:
    """소포신청 및 재조회 결과의 개인정보 없는 요약."""

    order_no: str
    req_no: str
    res_no: str
    regi_no: str
    rechecked: bool


# 이전 테스트 접수 호출부와의 호환성을 위한 이름이다.
ParcelTestReceipt = ParcelReceipt


@dataclass(frozen=True)
class RecipientAddressSuggestion:
    """우편번호 API 후보에서 만든 수취인 주소1·상세주소 제안."""

    postcode: str
    address1: str
    address2: str


def compact_text(value: str | None) -> str:
    return " ".join((value or "").split())


def element_text(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return compact_text("".join(element.itertext()))


def required_setting(settings: Mapping[str, object], key: str, label: str) -> str:
    value = str(settings.get(key, "") or "").strip()
    if not value:
        raise ParcelValidationError(f"설정 탭의 {label}({key}) 값이 비어 있습니다.")
    return value


def digits(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def text(value: object) -> str:
    value = "" if value is None else str(value)
    return "" if value.strip().lower() == "nan" else value.strip()


def _normalized_address_token(value: str) -> str:
    short_provinces = {
        "서울": "서울특별시", "부산": "부산광역시", "대구": "대구광역시",
        "인천": "인천광역시", "광주": "광주광역시", "대전": "대전광역시",
        "울산": "울산광역시", "세종": "세종특별자치시", "경기": "경기도",
        "강원": "강원특별자치도", "충북": "충청북도", "충남": "충청남도",
        "전북": "전북특별자치도", "전남": "전라남도", "경북": "경상북도",
        "경남": "경상남도", "제주": "제주특별자치도",
    }
    return short_provinces.get(value.strip(), value.strip())


def _split_standard_address(value: str) -> tuple[str, str]:
    """우편번호 API의 '도로명주소 (법정동, 건물명)'을 두 부분으로 나눈다."""
    match = re.match(r"^(.*?)(\s*\([^)]*\))$", text(value))
    if not match:
        return text(value), ""
    return match.group(1).strip(), match.group(2).strip()


def address_suggestion(source_address: str, candidate: Mapping[str, object]) -> RecipientAddressSuggestion:
    """표준 도로명주소와 원문 뒷부분을 합쳐 API용 주소1·상세주소를 제안한다.

    도로명과 건물번호가 원문 앞부분과 순서대로 일치할 때만 원문의 남은 부분을
    상세주소로 취급한다. 일치하지 않으면 후보 선택 후에도 원문을 억지로 분리하지
    않기 위해 표준주소의 괄호 정보만 상세주소로 반환한다.
    """
    standard_address = text(candidate.get("address"))
    address1, standard_detail = _split_standard_address(standard_address)
    if not address1:
        raise ParcelValidationError("우편번호 API 결과에 도로명주소가 없습니다.")

    source_tokens = text(source_address).split()
    base_tokens = address1.split()
    source_detail = ""
    if len(source_tokens) >= len(base_tokens):
        same_base = all(
            _normalized_address_token(source) == _normalized_address_token(base)
            for source, base in zip(source_tokens, base_tokens)
        )
        if same_base:
            source_detail = " ".join(source_tokens[len(base_tokens):]).strip()

    detail_parts = [part for part in (standard_detail, source_detail) if part]
    detail = " ".join(detail_parts).strip() or "-"
    return RecipientAddressSuggestion(
        postcode=digits(candidate.get("postcd")),
        address1=address1,
        address2=detail,
    )


def _postcode_response_items(root: ET.Element) -> list[dict[str, str]]:
    error = next((element for element in root.iter() if element.tag.rsplit("}", 1)[-1] == "error"), None)
    if error is not None:
        error_fields = {
            child.tag.rsplit("}", 1)[-1]: text(child.text)
            for child in error
        }
        # KpostPortal은 검색 결과가 없을 때도 HTTP 200 대신 ERR-125 XML을
        # 반환한다. 상세주소 토큰을 제거해 다음 검색어를 시도할 수 있게 한다.
        if error_fields.get("error_code") == "ERR-125":
            return []
        raise ParcelApiError(
            error_fields.get("error_code", ""),
            error_fields.get("message", "우편번호 API 오류 응답"),
        )

    items: list[dict[str, str]] = []
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] != "item":
            continue
        item = {
            child.tag.rsplit("}", 1)[-1]: text(child.text)
            for child in element
        }
        if item.get("postcd") and item.get("address"):
            items.append(item)
    return items


def postcode_lookup_items(query: str, api_key: str) -> list[dict[str, str]]:
    """우편번호 OpenAPI의 도로명주소 후보를 조회한다."""
    response = requests.get(
        POSTCODE_API_URL,
        params={
            "regkey": api_key,
            "target": "postNew",
            "query": text(query),
            "countPerPage": 20,
            "currentPage": 1,
        },
        timeout=15,
    )
    response.raise_for_status()
    return _postcode_response_items(ET.fromstring(response.content))


def _postcode_queries(source_address: str) -> Sequence[str]:
    """상세주소를 하나씩 제거해 표준 도로명주소까지 검색 범위를 넓힌다."""
    tokens = text(source_address).split()
    queries = []
    for count in range(len(tokens), max(len(tokens) - 8, 1), -1):
        query = " ".join(tokens[:count])
        if len(query) >= 2 and query not in queries:
            queries.append(query)
    return queries


def resolve_recipient_address(source_address: str, api_key: str) -> list[RecipientAddressSuggestion]:
    """원문 주소에서 표준주소 후보를 찾고 수취인 주소 필드 제안을 만든다."""
    source = text(source_address)
    if not source:
        raise ParcelValidationError("수취인 주소가 비어 있습니다.")
    if not text(api_key):
        raise ParcelValidationError("설정 탭의 우편번호 API 키가 비어 있습니다.")

    for query in _postcode_queries(source):
        items = postcode_lookup_items(query, text(api_key))
        if not items:
            continue
        suggestions = []
        seen = set()
        for item in items:
            suggestion = address_suggestion(source, item)
            identity = (suggestion.postcode, suggestion.address1, suggestion.address2)
            if identity not in seen:
                seen.add(identity)
                suggestions.append(suggestion)
        return suggestions
    return []


def normalize_content_code(value: object) -> str:
    """계약소포의 3자리 내용품코드를 정규화한다(예: 29 -> 029)."""
    code = digits(value)
    if not code or len(code) > 3:
        raise ParcelValidationError("내용품코드는 3자리 숫자로 입력해야 합니다.")
    return code.zfill(3)


def new_order_no(prefix: str) -> str:
    """빈 송장 주문번호 대신 사용할 충돌 가능성이 낮은 내부 식별자."""
    return f"{prefix}-{datetime.now():%Y%m%d%H%M%S}-{secrets.token_hex(3).upper()}"


def new_test_order_no() -> str:
    return new_order_no("EFTEST")


def new_real_order_no() -> str:
    return new_order_no("EFREAL")


def encrypt_reg_data(security_key: str, plain_text: str) -> str:
    """우체국 제공 Java/PHP 예제와 같은 SEED128 ECB/0x00 패딩을 적용한다."""
    key_bytes = security_key.encode("utf-8")
    if len(key_bytes) < 16:
        raise ParcelValidationError("접수용 보안키는 UTF-8 기준 최소 16바이트여야 합니다.")

    plain_bytes = plain_text.encode("utf-8")
    padded_length = max(16, ((len(plain_bytes) + 15) // 16) * 16)
    padded = plain_bytes.ljust(padded_length, b"\x00")
    cipher = Cipher(algorithms.SEED(key_bytes[:16]), modes.ECB())
    return cipher.encryptor().update(padded).hex()


def _request_plain_data(values: Mapping[str, object]) -> str:
    """암호화 전 key=value 문자열을 만든다.

    구분자인 ``&`` 또는 ``=``가 값에 포함되면 우체국 서버가 다른 항목으로 해석할 수
    있으므로 요청 전에 명확히 중단한다. 개인정보를 예외 메시지에 포함하지 않는다.
    """
    parts: list[str] = []
    for key, value in values.items():
        value_text = text(value)
        if "&" in value_text or "=" in value_text:
            raise ParcelValidationError(f"{key} 값에 사용할 수 없는 문자가 포함되어 있습니다.")
        parts.append(f"{key}={value_text}")
    return "&".join(parts)


def call_api(api_name: str, plain_data: str, api_key: str, security_key: str) -> ET.Element:
    response = requests.get(
        API_BASE_URL.format(api_name),
        params={"key": api_key, "regData": encrypt_reg_data(security_key, plain_data)},
        headers={
            "Connection": "keep-alive",
            "Host": "ship.epost.go.kr",
            "User-Agent": USER_AGENT,
        },
        timeout=20,
    )
    response.raise_for_status()
    return ET.fromstring(response.content)


def raise_if_api_error(root: ET.Element) -> None:
    code = element_text(root.find(".//error_code"))
    if code:
        raise ParcelApiError(code, element_text(root.find(".//message")) or "우체국 API 오류 응답")


def _first_result(root: ET.Element) -> dict[str, str]:
    result = {
        name: element_text(root.find(f".//{name}"))
        for name in ("reqNo", "resNo", "regiNo", "orderNo")
    }
    return result


def build_order_values(
    row: Mapping[str, object],
    settings: Mapping[str, object],
    *,
    test_yn: str,
) -> dict[str, str]:
    """기존 11열 송장 행을 계약소포 신청 파라미터로 변환한다."""
    if test_yn not in {"Y", "N"}:
        raise ValueError("test_yn은 Y 또는 N이어야 합니다.")
    recipient_name = text(row.get("수취인명"))
    recipient_zip = digits(row.get("우편번호"))
    recipient_address = text(row.get("수취인 주소"))
    recipient_detail_address = text(row.get("수취인 상세주소")) or "-"
    recipient_tel = digits(row.get("수취인 전화번호"))
    recipient_mobile = digits(row.get("수취인 이동통신"))
    goods_name = text(row.get("상품명"))

    missing = []
    if not recipient_name:
        missing.append("수취인명")
    if len(recipient_zip) != 5:
        missing.append("우편번호(5자리)")
    if not recipient_address:
        missing.append("수취인 주소")
    if not recipient_tel and not recipient_mobile:
        missing.append("수취인 연락처")
    if not goods_name:
        missing.append("상품명")
    if missing:
        raise ParcelValidationError("송장 행의 필수값이 비어 있습니다: " + ", ".join(missing))

    # 우체국 API는 recTel을 필수로 표시한다. 한쪽만 있는 기존 송장 행도 접수할 수
    # 있도록 존재하는 연락처를 두 필드에 넣는다.
    recipient_tel = recipient_tel or recipient_mobile
    recipient_mobile = recipient_mobile or recipient_tel
    order_no = text(row.get("주문번호")) or (
        new_test_order_no() if test_yn == "Y" else new_real_order_no()
    )
    if len(order_no) > 50:
        raise ParcelValidationError("주문번호는 50자 이하여야 합니다.")

    values = {
        "custNo": required_setting(settings, CONFIG_KEY_CUST_NO, "고객번호"),
        "apprNo": required_setting(settings, CONFIG_KEY_APPR_NO, "계약 승인번호"),
        "payType": required_setting(settings, CONFIG_KEY_PAY_TYPE, "요금납부구분"),
        "reqType": "1",
        "officeSer": required_setting(settings, CONFIG_KEY_OFFICE_SER, "공급지 코드"),
        "weight": required_setting(settings, CONFIG_KEY_WEIGHT, "기본 중량"),
        "volume": required_setting(settings, CONFIG_KEY_VOLUME, "기본 부피"),
        "microYn": required_setting(settings, CONFIG_KEY_MICRO_YN, "초소형소포 여부"),
        "orderNo": order_no,
        # ordCompNm은 API가 공란을 허용하므로 송장 원본 값을 그대로 사용한다.
        "ordCompNm": text(row.get("고객주문처명")),
        "recNm": recipient_name,
        "recZip": recipient_zip,
        "recAddr1": recipient_address,
        # 상세주소가 없는 기존 11열 송장 양식은 '-'로 보완한다.
        "recAddr2": recipient_detail_address,
        "recTel": recipient_tel,
        "recMob": recipient_mobile,
        "contCd": normalize_content_code(
            required_setting(settings, CONFIG_KEY_CONTENT_CODE, "내용품코드"),
        ),
        "goodsNm": goods_name,
        "printYn": required_setting(settings, CONFIG_KEY_PRINT_YN, "운송장 자체 출력 여부"),
        "testYn": test_yn,
    }

    if values["payType"] not in {"1", "2"}:
        raise ParcelValidationError("요금납부구분은 1 또는 2여야 합니다.")
    if values["microYn"] not in {"Y", "N"}:
        raise ParcelValidationError("초소형소포 여부는 Y 또는 N이어야 합니다.")
    if values["printYn"] not in {"Y", "N"}:
        raise ParcelValidationError("운송장 자체 출력 여부는 Y 또는 N이어야 합니다.")
    if not digits(values["weight"]) or not digits(values["volume"]):
        raise ParcelValidationError("기본 중량과 부피는 정수여야 합니다.")
    return values


def build_test_order_values(row: Mapping[str, object], settings: Mapping[str, object]) -> dict[str, str]:
    """기존 호출부용 테스트 소포신청 파라미터 변환기."""
    return build_order_values(row, settings, test_yn="Y")


def build_real_order_values(row: Mapping[str, object], settings: Mapping[str, object]) -> dict[str, str]:
    """실제 소포신청 파라미터를 만들고 자체 운송장 출력을 금지한다."""
    values = build_order_values(row, settings, test_yn="N")
    if values["printYn"] != "N":
        raise ParcelValidationError(
            "실접수 확인은 운송장 자체 출력 여부(epost_parcel_print_yn)를 N으로 설정해야 합니다.",
        )
    return values


def submit_order(
    row: Mapping[str, object],
    settings: Mapping[str, object],
    *,
    test_yn: str,
) -> ParcelReceipt:
    """소포신청 후 같은 주문번호로 접수 결과를 재조회한다."""
    api_key = required_setting(settings, CONFIG_KEY_API, "소포신청 인증키")
    security_key = required_setting(settings, CONFIG_KEY_SECURITY, "접수용 보안키")
    if test_yn == "Y":
        values = build_test_order_values(row, settings)
    elif test_yn == "N":
        values = build_real_order_values(row, settings)
    else:
        values = build_order_values(row, settings, test_yn=test_yn)
    root = call_api("InsertOrder", _request_plain_data(values), api_key, security_key)
    raise_if_api_error(root)
    result = _first_result(root)
    if not result["reqNo"] or not result["resNo"] or not result["regiNo"]:
        raise ParcelApiError("", "소포신청 응답에 신청번호·예약번호·등기번호가 모두 없습니다.")

    lookup_values = {
        "custNo": values["custNo"],
        "reqType": values["reqType"],
        "orderNo": values["orderNo"],
        "reqYmd": datetime.now().strftime("%Y%m%d"),
    }
    lookup_root = call_api("GetResInfo", _request_plain_data(lookup_values), api_key, security_key)
    raise_if_api_error(lookup_root)
    rechecked = _first_result(lookup_root)
    if rechecked["regiNo"] != result["regiNo"]:
        raise ParcelApiError("", "소포신청 재조회 결과의 등기번호가 최초 응답과 일치하지 않습니다.")

    return ParcelReceipt(
        order_no=values["orderNo"],
        req_no=result["reqNo"],
        res_no=result["resNo"],
        regi_no=result["regiNo"],
        rechecked=True,
    )


def submit_test_order(row: Mapping[str, object], settings: Mapping[str, object]) -> ParcelReceipt:
    """테스트 소포신청(testYn=Y) 후 같은 주문번호로 접수 결과를 재조회한다."""
    return submit_order(row, settings, test_yn="Y")


def submit_real_order(row: Mapping[str, object], settings: Mapping[str, object]) -> ParcelReceipt:
    """운송장 자체 출력 없이 실제 소포신청(testYn=N)을 한 건 접수한다."""
    return submit_order(row, settings, test_yn="N")
