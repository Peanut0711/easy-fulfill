import pandas as pd

from orders.matching import (
    CoupangOrderIndex,
    NaverOrderIndex,
    match_naver_invoice_rows,
    normalize_coupang_text,
)


columns = {
    "주문번호": "order", "수취인명": "name", "통합배송지": "address",
    "우편번호": "zipcode", "수취인연락처1": "phone",
}
orders = pd.DataFrame([
    ["1234567890123-A", "김", "서울 1", "12345", "010-1111-2222"],
    ["1234567890123-B", "김", "부산 2", "54321", "010-1111-2222"],
    ["1234567890123-C", "김", "서울 1", "12345", "010-1111-2222"],
], columns=columns.values())
prefixes = orders["order"].str[:13]

seoul = match_naver_invoice_rows(
    orders, prefixes, "1234567890123", columns,
    "김", "01011112222", "서울 1",
)
assert seoul["order"].tolist() == ["1234567890123-A", "1234567890123-C"]

busan = match_naver_invoice_rows(
    orders, prefixes, "1234567890123", columns,
    "김", "010-1111-2222", "부산 2",
)
assert busan["order"].tolist() == ["1234567890123-B"]

joined_address = match_naver_invoice_rows(
    orders, prefixes, "1234567890123", columns,
    "김", "010-1111-2222", ("서울", "1", "서울 1"),
)
assert joined_address["order"].tolist() == ["1234567890123-A", "1234567890123-C"]

for address in ("", "대전 3"):
    try:
        match_naver_invoice_rows(
            orders, prefixes, "1234567890123", columns,
            "김", "010-1111-2222", address,
        )
    except ValueError:
        pass
    else:
        raise AssertionError("주소가 없는 송장이나 다른 배송지는 자동 매칭하면 안 됩니다.")

same_address = orders.iloc[[0, 2]]
assert len(match_naver_invoice_rows(
    same_address, same_address["order"].str[:13], "1234567890123",
    columns, "", "", "",
)) == 2


# 주문번호 인덱스는 주문서 전체 비교와 같은 행을 같은 순서로 고른다.
indexed_orders = pd.concat([
    orders,
    pd.DataFrame([
        ["0012345678901-A", "이", "대전 3", "34567", "010-3333-4444"],
        ["0012345678901-B", "이", "대전 3", "34567", "010-3333-4444"],
        ["5555555555555-A", "박", None, None, None],
    ], columns=columns.values()),
], ignore_index=True)
indexed_prefixes = indexed_orders["order"].str[:13]
order_index = NaverOrderIndex(indexed_orders, indexed_prefixes, columns)
for args in (
    ("1234567890123", "김", "01011112222", "서울 1"),
    ("1234567890123", "김", "010-1111-2222", ("부산", "2", "부산 2")),
    ("0012345678901", "", "", ""),
    ("5555555555555", "", "", ""),
    ("9999999999999", "김", "01011112222", "서울 1"),
):
    expected = match_naver_invoice_rows(indexed_orders, indexed_prefixes, args[0], columns, *args[1:])
    assert order_index.match(*args) == list(expected.index), args
for address in ("", "대전 3"):
    try:
        order_index.match("1234567890123", "김", "010-1111-2222", address)
    except ValueError as error:
        assert "1234567890123" in str(error)
    else:
        raise AssertionError("인덱스 경로도 배송지를 특정할 수 없으면 중단해야 합니다.")

# 수취 정보 열을 찾지 못한 주문서는 후보가 한 행이어도 기존처럼 판별 단계에서 실패한다.
missing_zip = dict(columns, 우편번호=None)
try:
    NaverOrderIndex(indexed_orders, indexed_prefixes, missing_zip).match("5555555555555", "", "", "")
except KeyError:
    pass
else:
    raise AssertionError("우편번호 열이 없으면 기존처럼 오류가 나야 합니다.")


coupang_orders = pd.DataFrame({
    "name": ["김", " 김 ", "김", None, "nan", "이"],
    "phone": ["010-1111-2222", "010-1111-2222", "010-9999-8888", "010-5555-6666", "050", None],
})
coupang_index = CoupangOrderIndex(coupang_orders, "name", "phone")
assert coupang_index.match("김", "010-1111-2222") == [0, 1]
assert coupang_index.match("", "010-5555-6666") == [3]
assert coupang_index.match("", "050") == [4]
assert coupang_index.match("이", "") == [5]
assert coupang_index.match("박", "010-1111-2222") == []
assert sorted(coupang_index.phones_for_name("김")) == ["010-1111-2222", "010-1111-2222", "010-9999-8888"]
assert coupang_index.phones_for_name("박") is None
assert normalize_coupang_text(float("nan")) == ""
assert normalize_coupang_text(" NaN ") == ""
assert normalize_coupang_text(" 010 ") == "010"
