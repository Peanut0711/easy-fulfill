"""송장번호를 주문서 행에 안전하게 연결하는 규칙."""

from __future__ import annotations

from collections.abc import Callable, Hashable, Sequence

from shared.lazy_import import deferred_module

pd = deferred_module("pandas")  # 첫 화면 이후 처음 쓸 때 불러온다.


def _normalized_text(value) -> str:
    return "" if pd.isna(value) else " ".join(str(value).split())


def _phone_digits(value) -> str:
    return "".join(char for char in _normalized_text(value) if char.isdigit())


def normalize_coupang_text(value) -> str:
    """쿠팡 주문서·송장의 수취인명/전화번호 비교용 문자열(빈 값·'nan'은 빈 문자열)."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


_NAVER_IDENTITY_KEYS = ("수취인명", "통합배송지", "우편번호", "수취인연락처1")


def _naver_identity(values: Sequence[object]) -> tuple[str, str, str, str]:
    name, address, zipcode, phone = values
    return (
        _normalized_text(name),
        _normalized_text(address),
        _normalized_text(zipcode),
        _phone_digits(phone),
    )


def _select_naver_labels(
    labels: Sequence[Hashable],
    identity: Callable[[Hashable], tuple[str, str, str, str]],
    order_number: str,
    invoice_name: str,
    invoice_phone: str,
    invoice_addresses: str | tuple[str, ...],
) -> list[Hashable]:
    """같은 주문번호 후보 행 중 송장의 수취 정보에 해당하는 행을 고른다."""
    if not labels:
        return []

    identities = [identity(label) for label in labels]
    if len(set(identities)) == 1:
        return list(labels)

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
    return [label for label, key in zip(labels, identities) if key == selected]


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
    selected = _select_naver_labels(
        list(candidates.index),
        lambda label: _naver_identity(
            [order_df.at[label, columns[key]] for key in _NAVER_IDENTITY_KEYS]),
        order_number, invoice_name, invoice_phone, invoice_addresses,
    )
    return order_df.loc[selected]


class NaverOrderIndex:
    """정규화한 주문번호(앞 13자리)별 주문서 행을 한 번에 묶어 송장마다 재사용한다.

    송장마다 주문서 전체를 비교하던 방식과 같은 후보 행·같은 판별 규칙을 쓰며,
    주문서의 수취 정보 열은 매칭 중에 바뀌지 않는다는 전제로 만든다.
    """

    def __init__(self, order_df: pd.DataFrame, order_number_series: pd.Series, columns: dict[str, object]):
        self._order_df = order_df
        self._columns = columns
        self._labels = list(order_df.index)
        groups: dict[str, list[int]] = {}
        position_of = {label: position for position, label in enumerate(self._labels)}
        for label, number in zip(order_number_series.index, order_number_series.tolist()):
            groups.setdefault(number, []).append(position_of[label])
        self._groups = groups
        self._identity_values: list[list[object]] | None = None
        # 수취 정보 열이 없으면 기존처럼 판별 단계에서 오류가 나도록 단축 경로를 쓰지 않는다.
        self._identity_columns_found = all(columns[key] is not None for key in _NAVER_IDENTITY_KEYS)

    def _identity(self, position: int) -> tuple[str, str, str, str]:
        # 수취 정보 열은 배송지가 여러 곳인 주문번호를 만났을 때 처음 한 번만 읽는다.
        if self._identity_values is None:
            self._identity_values = [
                self._order_df[self._columns[key]].tolist() for key in _NAVER_IDENTITY_KEYS
            ]
        return _naver_identity([values[position] for values in self._identity_values])

    def match(
        self,
        order_number: str,
        invoice_name: str,
        invoice_phone: str,
        invoice_addresses: str | tuple[str, ...],
    ) -> list[Hashable]:
        """송장 한 건에 해당하는 주문서 행 라벨(주문서 순서)을 반환한다."""
        positions = self._groups.get(order_number, [])
        if len(positions) == 1 and self._identity_columns_found:
            # 후보가 한 행이면 기존 규칙에서도 수취 정보와 무관하게 그 행이 선택된다.
            return [self._labels[positions[0]]]
        selected = _select_naver_labels(
            positions, self._identity, order_number,
            invoice_name, invoice_phone, invoice_addresses,
        )
        return [self._labels[position] for position in selected]


class CoupangOrderIndex:
    """정규화한 (수취인명, 전화번호)별 주문서 행과 이름별 전화번호를 미리 묶는다."""

    def __init__(
        self,
        order_df: pd.DataFrame,
        name_column: object,
        phone_column: object,
        normalize: Callable[[object], str] = normalize_coupang_text,
    ):
        names = order_df[name_column].map(normalize).tolist()
        phones = order_df[phone_column].map(normalize).tolist()
        self._rows: dict[tuple[str, str], list[Hashable]] = {}
        self._phones_by_name: dict[str, list[str]] = {}
        for label, name, phone in zip(order_df.index, names, phones):
            self._rows.setdefault((name, phone), []).append(label)
            self._phones_by_name.setdefault(name, []).append(phone)

    def match(self, name: str, phone: str) -> list[Hashable]:
        """수취인명과 전화번호가 모두 같은 주문서 행 라벨(주문서 순서)을 반환한다."""
        return list(self._rows.get((name, phone), []))

    def phones_for_name(self, name: str) -> list[str] | None:
        """같은 수취인명의 주문서 전화번호 목록. 이름이 없으면 None."""
        return self._phones_by_name.get(name)
