from pathlib import Path
from tempfile import TemporaryDirectory

from orders.quick import (
    detect_quick_store,
    parse_coupang_quick_clipboard,
    parse_generic_quick_clipboard,
    parse_gmarket_quick_clipboard,
    parse_naver_quick_clipboard,
    save_invoice_excel,
)


sample = """수취인명: 조현준
연락처: 010-5375-9148
주소: 서울 강남구 테헤란로 306길 11, 우성빌딩 3층"""
assert parse_generic_quick_clipboard(sample) == {
    "수취인명": "조현준",
    "연락처": "010-5375-9148",
    "주소": "서울 강남구 테헤란로 306길 11, 우성빌딩 3층",
    "우편번호": "",
}
assert detect_quick_store(sample) == "generic"

with_postcode = """수취인명: 홍길동
휴대폰번호: 010-1234-5678
우편 번호: [06173]
배송지 주소: 서울 강남구 테헤란로 1"""
assert parse_generic_quick_clipboard(with_postcode)["우편번호"] == "06173"

unlabeled_recipient = """인천14
주소	인천광역시 서구 청라대로 65번길 20
우편번호	22335
전화번호	070-4732-1863"""
assert parse_generic_quick_clipboard(unlabeled_recipient) == {
    "수취인명": "인천14",
    "연락처": "070-4732-1863",
    "주소": "인천광역시 서구 청라대로 65번길 20",
    "우편번호": "22335",
}

coupang = parse_coupang_quick_clipboard("수취인명\t김고객\n연락처(안심번호)\t010-1\n배송주소\t(12345) 서울")
assert coupang["수취인명"] == "김고객" and coupang["우편번호"] == "12345"
assert detect_quick_store("배송주소\t서울") == "coupang"

naver = parse_naver_quick_clipboard("수취인명\n김고객\n연락처1\n010-2\n배송지\n서울")
assert naver["수취인명"] == "김고객" and naver["배송지"] == "서울"
assert detect_quick_store("연락처1\t010-2") == "naver"

gmarket = parse_gmarket_quick_clipboard("상품수령인\n김고객\n배송지주소\n12345 서울\n배송 요청사항\n문앞")
assert gmarket["우편번호"] == "12345" and gmarket["배송지주소"] == "서울"
assert detect_quick_store("상품수령인\t김고객") == "gmarket"

with TemporaryDirectory() as temp_dir:
    output = save_invoice_excel([{
        "주문번호": "1", "고객주문처명": "", "수취인명": "김고객", "우편번호": "12345",
        "수취인 주소": "서울", "수취인 전화번호": "010", "수취인 이동통신": "010",
        "상품명": "원본", "상품모델": "원본", "배송메세지": None, "비고": "",
    }], "테스트", temp_dir)
    assert output.exists() and output.suffix == ".xlsx"
