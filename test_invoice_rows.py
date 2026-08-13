import math

from orders.invoice import INVOICE_COLUMNS, build_invoice_rows


products = [{"상품명": "상품A", "옵션": "빨강", "수량": 2}]

naver_rows = build_invoice_rows("naver", {
    "N1": {
        "수취인명": "김네이버", "수취인연락처1": "010-1111-2222",
        "통합배송지": "서울", "배송메세지": math.nan, "우편번호": "123",
        "배송방법": "택배,등기,소포", "상품목록": products,
    },
    "N2": {
        "수취인명": "제외", "수취인연락처1": "010", "통합배송지": "부산",
        "배송메세지": "", "우편번호": "456", "배송방법": "방문수령",
        "상품목록": products,
    },
})
assert len(naver_rows) == 1
assert tuple(naver_rows[0]) == INVOICE_COLUMNS
assert naver_rows[0]["우편번호"] == "00123"
assert naver_rows[0]["배송메세지"] == ""
assert naver_rows[0]["상품명"] == "상품A (옵션: 빨강) - 2개"

coupang_row = build_invoice_rows("coupang", {
    "C1": {
        "수취인이름": "김쿠팡", "수취인전화번호": "010-2222-3333",
        "수취인주소": "대전", "배송메세지": "문 앞", "우편번호": 1234,
        "상품목록": products,
    }
})[0]
assert coupang_row["수취인 전화번호"] == "010-2222-3333"
assert coupang_row["수취인 이동통신"] == "010-2222-3333"
assert coupang_row["우편번호"] == "01234"

gmarket_row = build_invoice_rows("gmarket", {
    "G1": {
        "수령인명": "김지마켓", "수령인 전화번호": "02-1111-2222",
        "수령인 휴대폰": "010-3333-4444", "주소": "광주",
        "배송시 요구사항": "경비실", "우편번호": "54321", "상품목록": products,
    }
})[0]
assert gmarket_row["수취인 전화번호"] == "02-1111-2222"
assert gmarket_row["수취인 이동통신"] == "010-3333-4444"

st11_row = build_invoice_rows("11st", {
    "S1": {
        "수취인명": "김11번가", "전화번호": "031-111-2222",
        "휴대폰번호": "010-4444-5555", "주소": "수원",
        "배송메시지": "부재 시 연락", "우편번호": "16400", "상품목록": products,
    }
})[0]
assert st11_row["수취인 전화번호"] == "031-111-2222"
assert st11_row["수취인 이동통신"] == "010-4444-5555"

try:
    build_invoice_rows("unknown", {})
except ValueError:
    pass
else:
    raise AssertionError("지원하지 않는 스토어는 ValueError여야 합니다.")
