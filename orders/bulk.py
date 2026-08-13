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


def build_naver_orders(
    dataframe: pd.DataFrame,
    columns: Mapping[str, object],
    product_codes: Mapping[str, str],
    normalize_mapping_key,
) -> dict[str, dict]:
    """네이버 주문 DataFrame을 기존 orders 딕셔너리 구조로 변환한다.

    주문번호 앞 13자리를 하나의 주문으로 묶는 기존 규칙과,
    '택배,등기,소포' 요청이 하나라도 있으면 해당 주문 전체에 적용하는 규칙을 보존한다.
    """
    orders: dict[str, dict] = {}
    grouped_rows: dict[str, list[tuple[str, pd.Series]]] = {}

    for _, row in dataframe.iterrows():
        order_number = _text(row[columns["주문번호"]]).strip()
        if not order_number:
            continue
        pattern = order_number[:13] if len(order_number) >= 13 else order_number
        grouped_rows.setdefault(pattern, []).append((order_number, row))

    amount_column = columns.get("최종 상품별 총 주문금액")
    for pattern, rows in grouped_rows.items():
        first_order = rows[0][1]
        delivery_column = columns["배송방법(구매자 요청)"]
        first_delivery = _text(first_order[delivery_column]) or "배송방법 오류"
        order = {
            "주문번호목록": [order_number for order_number, _ in rows],
            "수취인명": _text(first_order[columns["수취인명"]]),
            "수취인연락처1": _text(first_order[columns["수취인연락처1"]]),
            "통합배송지": _text(first_order[columns["통합배송지"]]),
            "구매자연락처": _text(first_order[columns["구매자연락처"]]),
            "배송메세지": _text(first_order[columns["배송메세지"]]),
            "우편번호": _text(first_order[columns["우편번호"]]),
            "배송방법": first_delivery,
            "상품수": 0,
            "상품목록": [],
            "주문총액": 0,
        }
        has_standard_delivery = False
        for _, row in rows:
            delivery_method = _text(row[delivery_column])
            if delivery_method == "택배,등기,소포":
                has_standard_delivery = True
            product_number = normalize_mapping_key(row[columns["상품번호"]])
            product_amount = _number(row[amount_column]) if amount_column is not None else 0
            order["주문총액"] += product_amount
            order["상품수"] += 1
            order["상품목록"].append({
                "상품명": _text(row[columns["상품명"]]),
                "수량": _quantity(row[columns["수량"]]),
                "옵션": _text(row[columns["옵션정보"]]) or "없음",
                "상품코드": product_codes.get(product_number, "") or "        ",
                "금액": product_amount,
                "상품번호": product_number,
            })
        if has_standard_delivery:
            order["배송방법"] = "택배,등기,소포"
        orders[pattern] = order
    return orders


def build_gmarket_orders(dataframe: pd.DataFrame, columns: Mapping[str, object]) -> dict[str, dict]:
    """지마켓 주문 DataFrame을 기존 orders 딕셔너리 구조로 변환한다."""
    orders: dict[str, dict] = {}
    sale_column = columns.get("판매금액")
    shipping_column = columns.get("배송비 금액")
    additional_column = columns.get("추가구성")

    for _, row in dataframe.iterrows():
        order_number = _text(row[columns["주문번호"]]).strip()
        if not order_number:
            continue
        if order_number not in orders:
            orders[order_number] = {
                "수령인명": _text(row[columns["수령인명"]]),
                "주소": _text(row[columns["주소"]]),
                "수령인 전화번호": _text(row[columns["수령인 전화번호"]]),
                "수령인 휴대폰": _text(row[columns["수령인 휴대폰"]]),
                "배송시 요구사항": _text(row[columns["배송시 요구사항"]]),
                "우편번호": _text(row[columns["우편번호"]]),
                "상품목록": [],
                "판매금액": _number(row[sale_column]) if sale_column is not None else 0,
                "배송비 금액": _number(row[shipping_column]) if shipping_column is not None else 0,
            }
        option = _text(row[columns["옵션"]]).strip()
        if option.lower() in ("nan", ""):
            option = "없음"
        additional_config = _text(row[additional_column]).strip() if additional_column is not None else ""
        if additional_config.lower() == "nan":
            additional_config = ""
        orders[order_number]["상품목록"].append({
            "상품명": _text(row[columns["상품명"]]),
            "옵션": option,
            "수량": _quantity(row[columns["수량"]]),
            "추가구성": additional_config,
        })
    return orders


def consolidate_gmarket_orders(orders: Mapping[str, dict]) -> dict[str, dict]:
    """기존 지마켓 표시 규칙대로 같은 수령인명의 주문을 통합한다."""
    consolidated_orders: dict[str, dict] = {}
    for order_number, info in orders.items():
        customer_name = info["수령인명"]
        if customer_name not in consolidated_orders:
            consolidated_orders[customer_name] = {
                "수령인명": customer_name,
                "주문번호목록": [order_number],
                "상품목록": info["상품목록"].copy(),
                "총판매금액": info.get("판매금액", 0),
                "총배송비금액": info.get("배송비 금액", 0),
            }
            continue
        target = consolidated_orders[customer_name]
        target["주문번호목록"].append(order_number)
        target["상품목록"].extend(info["상품목록"])
        target["총판매금액"] += info.get("판매금액", 0)
        target["총배송비금액"] += info.get("배송비 금액", 0)
    return consolidated_orders
