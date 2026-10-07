import pandas as pd

from orders.matching import match_naver_invoice_rows


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
