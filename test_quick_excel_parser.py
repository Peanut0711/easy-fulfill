from pathlib import Path
from runpy import run_path


module = run_path(
    Path(__file__).with_name("easy-fulfill.py"),
    run_name="easy_fulfill_test",
)
parse = module["_parse_generic_quick_clipboard"]

sample = """수령인: 조현식 파트장님
연락처: 010-5375-9148
주소: 서울 강남구 영동대로106길 11, 현성빌딩 패스트파이브 삼성4호점 3층"""
assert parse(sample) == {
    "수취인명": "조현식 파트장님",
    "연락처": "010-5375-9148",
    "주소": "서울 강남구 영동대로106길 11, 현성빌딩 패스트파이브 삼성4호점 3층",
    "우편번호": "",
}
assert module["MainWindow"].detect_store_from_clipboard(None, sample) == "generic"

with_postcode = """수취인명：홍길동
휴대전화번호: 010-1234-5678
우편 번호: [06173]
배송지 주소: 서울 강남구 테헤란로 1"""
assert parse(with_postcode)["우편번호"] == "06173"
