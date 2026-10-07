"""송장번호를 주문서 행에 안전하게 연결하는 규칙."""

from __future__ import annotations

import pandas as pd


def _normalized_text(value) -> str:
    return "" if pd.isna(value) else " ".join(str(value).split())


def _phone_digits(value) -> str:
    return "".join(char for char in _normalized_text(value) if char.isdigit())


def match_naver_invoice_rows(
    order_df: pd.DataFrame,
    order_number_series: pd.Series,
    order_number: str,
    columns: dict[str, object],
    invoice_name: str,
    invoice_phone: str,
    invoice_addresses: str | tuple[str, ...],
) -> pd.DataFrame:
    """같은 주문번호에 배송지가 여럿이면 송장의 수취 정보로 한 배송지만 고른다."""
    candidates = order_df[order_number_series == order_number]
    if candidates.empty:
        return candidates

    def identity(row) -> tuple[str, str, str, str]:
        return (
            _normalized_text(row[columns["수취인명"]]),
            _normalized_text(row[columns["통합배송지"]]),
            _normalized_text(row[columns["우편번호"]]),
            _phone_digits(row[columns["수취인연락처1"]]),
        )

    identities = candidates.apply(identity, axis=1)
    if len(set(identities)) == 1:
        return candidates

    name = _normalized_text(invoice_name)
    phone = _phone_digits(invoice_phone)
    addresses = (invoice_addresses,) if isinstance(invoice_addresses, str) else invoice_addresses
    normalized_addresses = {_normalized_text(address) for address in addresses} - {""}
    if not name or not phone or not normalized_addresses:
        raise ValueError(
            f"주문번호 {order_number}: 배송지가 여러 건이므로 송장 수취인명·전화번호·주소가 모두 필요합니다."
        )

    matching_identities = {
        key for key in identities
        if key[0] == name and key[1] in normalized_addresses and key[3] == phone
    }
    if len(matching_identities) != 1:
        raise ValueError(
            f"주문번호 {order_number}: 배송지가 여러 건인데 송장 주소로 한 건을 특정할 수 없습니다."
        )
    selected = next(iter(matching_identities))
    return candidates[identities.map(lambda key: key == selected)]
