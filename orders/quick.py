"""퀵 엑셀용 클립보드 파싱과 송장 Excel 파일 생성."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd


INVOICE_COLUMNS = [
    "주문번호", "고객주문처명", "수취인명", "우편번호", "수취인 주소",
    "수취인 전화번호", "수취인 이동통신", "상품명", "상품모델", "배송메세지", "비고",
]


def extract_zip_code(text: object) -> str:
    """텍스트에서 첫 5자리 우편번호를 반환한다."""
    match = re.search(r"(?:\(|\[)?(\d{5})(?:\)|\])?", str(text or ""))
    return match.group(1) if match else ""


def parse_generic_quick_clipboard(clipboard_text: object) -> dict[str, str]:
    """일반 라벨 형식의 단건 주문정보를 추출한다."""
    result = {"수취인명": "", "연락처": "", "주소": "", "우편번호": ""}
    lines = [line.strip() for line in str(clipboard_text or "").splitlines() if line.strip()]
    first_line = lines[0] if lines else ""
    for line in lines:
        parts = re.split(r"\s*[:：]\s*|\t+", line, maxsplit=1)
        if len(parts) != 2:
            continue
        key = re.sub(r"[\s()（）_-]", "", parts[0])
        value = parts[1].strip()
        field = None
        if re.fullmatch(r"(?:상품)?(?:수령인|수취인)(?:명)?", key):
            field = "수취인명"
        elif re.fullmatch(
            r"(?:수령인|수취인)?(?:연락처(?:안심번호|\d*)?|전화번호|"
            r"휴대폰(?:번호)?|휴대전화(?:번호)?|핸드폰(?:번호)?|모바일(?:번호)?)", key,
        ):
            field = "연락처"
        elif re.fullmatch(r"(?:수령인|수취인)?(?:주소|배송주소|배송지(?:주소)?)", key):
            field = "주소"
        elif key == "우편번호":
            field = "우편번호"
        if field and value and not result[field]:
            result[field] = value
    if (not result["수취인명"] and result["연락처"] and result["주소"]
            and not re.search(r"[:：\t]", first_line)):
        result["수취인명"] = first_line
    result["우편번호"] = extract_zip_code(result["우편번호"]) or extract_zip_code(result["주소"])
    return result


def _clipboard_tokens(clipboard_text: object) -> list[str]:
    return [
        token.strip()
        for line in str(clipboard_text or "").replace("\r", "").split("\n")
        if line.strip()
        for token in line.split("\t")
        if token.strip()
    ]


def detect_quick_store(clipboard_text: object) -> str | None:
    """클립보드 텍스트의 스토어 양식을 판별한다."""
    tokens = _clipboard_tokens(clipboard_text)
    if "연락처(안심번호)" in tokens or "배송주소" in tokens:
        return "coupang"
    gmarket_address = tokens[tokens.index("배송지주소") + 1] if "배송지주소" in tokens and tokens.index("배송지주소") + 1 < len(tokens) else ""
    if "상품수령인" in tokens or "배송 요청사항" in tokens or re.match(r"^\d{5}\b", gmarket_address or ""):
        return "gmarket"
    if "연락처1" in tokens or "배송지" in tokens:
        return "naver"
    generic = parse_generic_quick_clipboard(clipboard_text)
    return "generic" if any(generic.values()) else None


def parse_coupang_quick_clipboard(clipboard_text: object) -> dict[str, str]:
    result = {key: "" for key in ("수취인명", "연락처(안심번호)", "배송주소", "배송메모", "우편번호")}
    tokens = _clipboard_tokens(clipboard_text)
    for index, token in enumerate(tokens[:-1]):
        if token in result:
            result[token] = tokens[index + 1]
    result["우편번호"] = result["우편번호"] or extract_zip_code(result["배송주소"])
    return result


def _parse_continuation_fields(clipboard_text: object, fields: tuple[str, ...]) -> dict[str, str]:
    result = {field: "" for field in fields}
    current = None
    for token in _clipboard_tokens(clipboard_text):
        if token in result:
            current = token
        elif current:
            result[current] = f"{result[current]} {token}".strip()
    return result


def parse_naver_quick_clipboard(clipboard_text: object) -> dict[str, str]:
    return _parse_continuation_fields(clipboard_text, ("수취인명", "연락처1", "연락처2", "배송지", "배송메모"))


def parse_gmarket_quick_clipboard(clipboard_text: object) -> dict[str, str]:
    result = _parse_continuation_fields(clipboard_text, ("상품수령인", "연락처1", "연락처2", "배송지주소", "배송 요청사항", "우편번호"))
    zipcode = extract_zip_code(result["배송지주소"])
    result["우편번호"] = zipcode
    if zipcode:
        result["배송지주소"] = re.sub(rf"^\s*{re.escape(zipcode)}\s*", "", result["배송지주소"]).strip()
    return result


def save_invoice_excel(invoice_data, filename_prefix: str, output_dir: Path | str = "output") -> Path:
    """송장 데이터를 기존 양식의 Excel 파일로 저장하고 경로를 반환한다."""
    directory = Path(output_dir)
    directory.mkdir(exist_ok=True)
    output_file = (directory / f"{filename_prefix}_{datetime.now():%Y%m%d%H%M%S}.xlsx").resolve()
    dataframe = pd.DataFrame(invoice_data)
    if "배송메세지" in dataframe.columns:
        dataframe["배송메세지"] = dataframe["배송메세지"].fillna("")
    for column in ("상품명", "상품모델"):
        if column in dataframe.columns:
            dataframe[column] = "전자제품"
    dataframe = dataframe[INVOICE_COLUMNS]
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        workbook = writer.book
        center = workbook.add_format({"align": "center", "valign": "vcenter"})
        header = workbook.add_format({"align": "center", "valign": "vcenter", "bold": True})
        for index, column in enumerate(dataframe.columns):
            maximum = max(dataframe[column].astype(str).apply(len).max(), len(str(column)))
            width = maximum * 2 if any("ㄱ" <= char <= "ㆎ" or "가" <= char <= "힣" for char in str(column)) else maximum
            worksheet.set_column(index, index, width + 2, center)
        for index, column in enumerate(dataframe.columns):
            worksheet.write(0, index, column, header)
    return output_file
