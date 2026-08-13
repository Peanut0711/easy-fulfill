"""스토어 주문 구조를 하이제니스 송장 행으로 변환하는 순수 함수."""

from __future__ import annotations

from collections.abc import Mapping

import pandas as pd


INVOICE_COLUMNS = (
    "주문번호",
    "고객주문처명",
    "수취인명",
    "우편번호",
    "수취인 주소",
    "수취인 전화번호",
    "수취인 이동통신",
    "상품명",
    "상품모델",
    "배송메세지",
    "비고",
)


def _text(value) -> str:
    if pd.isna(value) or str(value).lower() == "nan":
        return ""
    return str(value)


def _zipcode(value) -> str:
    zipcode = _text(value).strip()
    return zipcode.zfill(5) if zipcode.isdigit() else zipcode


def _products_text(products) -> str:
    return "\n".join(
        f"{product['상품명']} (옵션: {product['옵션']}) - {product['수량']}개"
        for product in products
    )


def _invoice_row(
    order_number,
    recipient_name,
    zipcode,
    address,
    phone,
    mobile,
    products,
    message,
) -> dict[str, object]:
    return {
        "주문번호": order_number,
        "고객주문처명": "",
        "수취인명": recipient_name,
        "우편번호": _zipcode(zipcode),
        "수취인 주소": address,
        "수취인 전화번호": phone,
        "수취인 이동통신": mobile,
        "상품명": _products_text(products),
        "상품모델": "전자제품",
        "배송메세지": _text(message),
        "비고": "",
    }


def build_invoice_rows(store_type: str, orders: Mapping[str, dict]) -> list[dict[str, object]]:
    """기존 스토어별 orders 딕셔너리를 송장 Excel 행 목록으로 변환한다."""
    if store_type not in {"naver", "coupang", "gmarket", "11st"}:
        raise ValueError(f"지원하지 않는 스토어 유형입니다: {store_type}")

    rows: list[dict[str, object]] = []

    for order_number, info in orders.items():
        if store_type == "naver":
            if info["배송방법"] != "택배,등기,소포":
                continue
            rows.append(_invoice_row(
                order_number,
                info["수취인명"],
                info["우편번호"],
                info["통합배송지"],
                info["수취인연락처1"],
                info["수취인연락처1"],
                info["상품목록"],
                info.get("배송메세지", ""),
            ))
        elif store_type == "coupang":
            rows.append(_invoice_row(
                order_number,
                info["수취인이름"],
                info.get("우편번호", ""),
                info["수취인주소"],
                info["수취인전화번호"],
                info["수취인전화번호"],
                info["상품목록"],
                info.get("배송메세지", ""),
            ))
        elif store_type == "gmarket":
            rows.append(_invoice_row(
                order_number,
                info["수령인명"],
                info.get("우편번호", ""),
                info["주소"],
                info["수령인 전화번호"],
                info["수령인 휴대폰"],
                info["상품목록"],
                info.get("배송시 요구사항", ""),
            ))
        elif store_type == "11st":
            rows.append(_invoice_row(
                order_number,
                info["수취인명"],
                info.get("우편번호", ""),
                info["주소"],
                info["전화번호"],
                info["휴대폰번호"],
                info["상품목록"],
                info.get("배송메시지", ""),
            ))
    return rows
