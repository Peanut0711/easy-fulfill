"""대량 주문 Excel DataFrame을 화면 독립 주문 구조로 변환하는 함수."""

from __future__ import annotations

from typing import Mapping

import pandas as pd


def find_columns(dataframe: pd.DataFrame, required: tuple[str, ...]) -> tuple[dict[str, object], list[str]]:
    """공백을 제외한 헤더 정확 일치로 필요한 열을 찾는다."""
    columns = {name: None for name in required}
    for column in dataframe.columns:
        name = str(column).strip()
        if name in columns:
            columns[name] = column
    return columns, [name for name, column in columns.items() if column is None]


def _text(value) -> str:
    return "" if pd.isna(value) else str(value)


def _number(value) -> float:
    if pd.isna(value):
        return 0.0
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


def _quantity(value) -> int:
    if pd.isna(value):
        return 1
    try:
        return int(value)
    except (TypeError, ValueError):
        return 1


def build_coupang_orders(
    dataframe: pd.DataFrame,
    columns: Mapping[str, object],
    product_codes: Mapping[str, str],
    option_to_product_no: Mapping[str, str],
    normalize_mapping_key,
) -> dict[str, dict]:
    """쿠팡 주문 DataFrame을 기존 orders 딕셔너리 구조로 변환한다."""
    orders: dict[str, dict] = {}
    for _, row in dataframe.iterrows():
        order_number = _text(row[columns["주문번호"]]).strip()
        if not order_number:
            continue
        if order_number not in orders:
            orders[order_number] = {
                "수취인이름": _text(row[columns["수취인이름"]]),
                "수취인주소": _text(row[columns["수취인 주소"]]),
                "수취인전화번호": _text(row[columns["수취인전화번호"]]),
                "배송메세지": _text(row[columns["배송메세지"]]),
                "우편번호": _text(row[columns["우편번호"]]),
                "상품목록": [], "결제액": 0,
            }
        payment_column = columns.get("결제액")
        if payment_column is not None:
            orders[order_number]["결제액"] += _number(row[payment_column])
        option_id = normalize_mapping_key(row[columns["옵션ID"]])
        orders[order_number]["상품목록"].append({
            "상품명": _text(row[columns["노출상품명(옵션명)"]]),
            "옵션": _text(row[columns["등록옵션명"]]),
            "수량": _quantity(row[columns["구매수(수량)"]]),
            "상품코드": product_codes.get(option_id, ""),
            "쿠팡상품번호": option_to_product_no.get(option_id, "") if option_id else "",
        })
    return orders


def build_11st_orders(dataframe: pd.DataFrame, columns: Mapping[str, object]) -> dict[str, dict]:
    """11번가 주문 DataFrame을 기존 orders 딕셔너리 구조로 변환한다."""
    orders: dict[str, dict] = {}
    for _, row in dataframe.iterrows():
        order_number = _text(row[columns["주문번호"]]).strip()
        if not order_number:
            continue
        if order_number not in orders:
            orders[order_number] = {
                "수취인명": _text(row[columns["수취인"]]),
                "주소": _text(row[columns["주소"]]),
                "휴대폰번호": _text(row[columns["휴대폰번호"]]),
                "전화번호": _text(row[columns["전화번호"]]),
                "우편번호": _text(row[columns["우편번호"]]),
                "배송메시지": _text(row[columns.get("배송메시지")]) if columns.get("배송메시지") is not None else "",
                "상품목록": [],
                "주문금액": _number(row[columns["주문금액"]]) if columns.get("주문금액") is not None else 0,
            }
        option = _text(row[columns["옵션"]]).strip() or "없음"
        orders[order_number]["상품목록"].append({
            "상품명": _text(row[columns["상품명"]]), "옵션": option,
            "수량": _quantity(row[columns["수량"]]),
        })
    return orders
