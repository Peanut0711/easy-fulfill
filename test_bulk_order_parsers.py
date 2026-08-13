import pandas as pd

from orders.bulk import build_11st_orders, build_coupang_orders


coupang_columns = {
    "주문번호": "order", "수취인이름": "name", "수취인 주소": "address",
    "수취인전화번호": "phone", "배송메세지": "message", "우편번호": "zip",
    "노출상품명(옵션명)": "product", "등록옵션명": "option", "구매수(수량)": "quantity",
    "옵션ID": "option_id", "결제액": "payment",
}
coupang_orders = build_coupang_orders(
    pd.DataFrame([
        ["C1", "김", "서울", "010", "문앞", "01234", "상품1", "빨강", 2, "100.0", "1,000"],
        ["C1", "김", "서울", "010", "문앞", "01234", "상품2", "파랑", None, "200", "2,500"],
    ], columns=coupang_columns.values()),
    coupang_columns, {"100": "P100", "200": "P200"}, {"100": "VP100"},
    lambda value: str(value).replace(".0", "") if not pd.isna(value) else "",
)
assert coupang_orders["C1"]["결제액"] == 3500.0
assert [item["수량"] for item in coupang_orders["C1"]["상품목록"]] == [2, 1]
assert coupang_orders["C1"]["상품목록"][0]["상품코드"] == "P100"
assert coupang_orders["C1"]["상품목록"][0]["쿠팡상품번호"] == "VP100"

st11_columns = {
    "주문번호": "order", "수취인": "name", "상품명": "product", "옵션": "option",
    "수량": "quantity", "주문금액": "amount", "휴대폰번호": "mobile", "전화번호": "phone",
    "우편번호": "zip", "주소": "address", "배송메시지": "message",
}
st11_orders = build_11st_orders(
    pd.DataFrame([
        ["S1", "이", "상품1", None, None, "1,200", "010", "02", "12345", "서울", "문앞"],
        ["S1", "이", "상품2", "대", 3, "1,200", "010", "02", "12345", "서울", "문앞"],
    ], columns=st11_columns.values()),
    st11_columns,
)
assert st11_orders["S1"]["주문금액"] == 1200.0
assert st11_orders["S1"]["상품목록"] == [
    {"상품명": "상품1", "옵션": "없음", "수량": 1},
    {"상품명": "상품2", "옵션": "대", "수량": 3},
]
