import pandas as pd

from orders.bulk import (
    build_11st_orders,
    build_coupang_orders,
    build_gmarket_orders,
    build_naver_orders,
    consolidate_11st_orders,
    consolidate_gmarket_orders,
)
from orders.invoice import build_invoice_rows


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

st11_delivery_orders = build_11st_orders(
    pd.DataFrame([
        ["S1", "이", "상품1", "", 1, 1000, "010", "02", "12345", "서울", ""],
        ["S2", "이", "상품2", "", 1, 2000, "010", "02", "54321", "부산", ""],
        ["S3", "이", "상품3", "", 1, 3000, "010", "02", "12345", "서울", ""],
    ], columns=st11_columns.values()),
    st11_columns,
)
st11_consolidated = consolidate_11st_orders(st11_delivery_orders)
assert list(st11_consolidated) == ["S1", "S2"]
assert st11_consolidated["S1"]["주문번호목록"] == ["S1", "S3"]
assert st11_consolidated["S2"]["주문번호목록"] == ["S2"]
assert [(row["주문번호"], row["수취인 주소"]) for row in build_invoice_rows("11st", st11_delivery_orders)] == [
    ("S1", "서울"), ("S2", "부산"), ("S3", "서울"),
]
st11_same_number = build_11st_orders(
    pd.DataFrame([
        ["S9", "이", "상품1", "", 1, 1000, "010", "02", "12345", "서울", ""],
        ["S9", "이", "상품2", "", 1, 2000, "010", "02", "54321", "부산", ""],
    ], columns=st11_columns.values()), st11_columns,
)
assert list(st11_same_number) == ["S9", "S9|배송지2"]
assert [(row["주문번호"], row["수취인 주소"]) for row in build_invoice_rows("11st", st11_same_number)] == [
    ("S9", "서울"), ("S9", "부산"),
]

naver_columns = {
    "주문번호": "order", "수취인명": "name", "수취인연락처1": "recipient_phone",
    "통합배송지": "address", "구매자연락처": "buyer_phone", "배송메세지": "message",
    "상품명": "product", "옵션정보": "option", "수량": "quantity", "우편번호": "zip",
    "상품번호": "product_no", "배송방법(구매자 요청)": "delivery", "최종 상품별 총 주문금액": "amount",
}
naver_orders = build_naver_orders(
    pd.DataFrame([
        ["N123456789012-A", "박", "010", "서울", "011", "문앞", "상품1", None, 2, "12345", "100.0", "일반배송", "1,200"],
        ["N123456789012-B", "박", "010", "서울", "011", "문앞", "상품2", "대", None, "12345", "200", "택배,등기,소포", "800"],
    ], columns=naver_columns.values()),
    naver_columns, {"100": "P100", "200": "P200"},
    lambda value: str(value).replace(".0", "") if not pd.isna(value) else "",
)
assert list(naver_orders) == ["N123456789012"]
assert naver_orders["N123456789012"]["주문총액"] == 2000.0
assert naver_orders["N123456789012"]["배송방법"] == "택배,등기,소포"
assert naver_orders["N123456789012"]["상품목록"][0]["옵션"] == "없음"
assert naver_orders["N123456789012"]["상품목록"][1]["수량"] == 1

naver_split_orders = build_naver_orders(
    pd.DataFrame([
        ["N123456789012-A", "박", "010", "서울", "011", "문앞", "상품1", "", 1, "12345", "100", "택배,등기,소포", 1000],
        ["N123456789012-B", "박", "010", "부산", "011", "문앞", "상품2", "", 1, "54321", "200", "택배,등기,소포", 2000],
        ["N123456789012-C", "박", "010", "서울", "011", "문앞", "상품3", "", 1, "12345", "300", "택배,등기,소포", 3000],
    ], columns=naver_columns.values()),
    naver_columns, {}, lambda value: str(value),
)
assert list(naver_split_orders) == ["N123456789012", "N123456789012|배송지2"]
assert naver_split_orders["N123456789012"]["주문번호목록"] == ["N123456789012-A", "N123456789012-C"]
assert naver_split_orders["N123456789012|배송지2"]["주문번호목록"] == ["N123456789012-B"]
assert [(row["주문번호"], row["수취인 주소"]) for row in build_invoice_rows("naver", naver_split_orders)] == [
    ("N123456789012", "서울"), ("N123456789012", "부산"),
]

gmarket_columns = {
    "주문번호": "order", "수령인명": "name", "주소": "address", "수령인 전화번호": "phone",
    "수령인 휴대폰": "mobile", "상품명": "product", "옵션": "option", "수량": "quantity",
    "배송시 요구사항": "message", "우편번호": "zip", "판매금액": "sale",
    "추가구성": "additional", "배송비 금액": "shipping",
}
gmarket_orders = build_gmarket_orders(
    pd.DataFrame([
        ["G1", "최", "서울", None, "010", "상품1", None, 2, "문앞", "12345", "10,000", "사은품", "3,000"],
        ["G1", "최", "서울", None, "010", "상품2", "대", None, "문앞", "12345", "10,000", None, "3,000"],
        ["G2", "최", "서울2", "02", "010", "상품3", "소", 1, "경비실", "12345", "5,000", "", "0"],
        ["G3", "최", "서울", None, "010", "상품4", "", 1, "문앞", "12345", "2,000", "", "0"],
    ], columns=gmarket_columns.values()),
    gmarket_columns,
)
assert gmarket_orders["G1"]["판매금액"] == 10000.0
assert gmarket_orders["G1"]["상품목록"][0]["옵션"] == "없음"
assert gmarket_orders["G1"]["상품목록"][1]["수량"] == 1
assert gmarket_orders["G1"]["상품목록"][0]["추가구성"] == "사은품"
gmarket_consolidated = consolidate_gmarket_orders(gmarket_orders)
assert list(gmarket_consolidated) == ["G1", "G2"]
assert gmarket_consolidated["G1"]["주문번호목록"] == ["G1", "G3"]
assert gmarket_consolidated["G2"]["주문번호목록"] == ["G2"]
assert gmarket_consolidated["G1"]["총판매금액"] == 12000.0
assert gmarket_consolidated["G1"]["총배송비금액"] == 3000.0
assert [(row["주문번호"], row["수취인 주소"]) for row in build_invoice_rows("gmarket", gmarket_orders)] == [
    ("G1", "서울"), ("G2", "서울2"), ("G3", "서울"),
]
gmarket_same_number = build_gmarket_orders(
    pd.DataFrame([
        ["G9", "최", "서울", "02", "010", "상품1", "", 1, "", "12345", 1000, "", 0],
        ["G9", "최", "부산", "02", "010", "상품2", "", 1, "", "54321", 2000, "", 0],
    ], columns=gmarket_columns.values()), gmarket_columns,
)
assert list(gmarket_same_number) == ["G9", "G9|배송지2"]
assert [(row["주문번호"], row["수취인 주소"]) for row in build_invoice_rows("gmarket", gmarket_same_number)] == [
    ("G9", "서울"), ("G9", "부산"),
]
