# Easy Fulfill

오픈마켓(스마트스토어, 쿠팡 등)에서 주문한 제품 리스트 엑셀 파일을 입력받아, 작업자가 출고를 쉽게 할 수 있도록 작업지시서와 송장 양식을 자동 생성하는 Python GUI 프로그램입니다.

## 요구사항

- Python 3.12.2
- PySide6 6.8.1
- pandas 2.2.1
- openpyxl 3.1.2

## 설치 방법

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

## 실행 방법

```bash
python easy-fulfill.py
```

## 기능

- 엑셀 파일(.xlsx) 선택
- 선택된 파일 경로 표시
- 작업지시서 생성 (개발 중)

## 네이버 상세 → 쿠팡 HTML 작성용 변환

네이버 SmartEditor ONE 상세를 읽어 텍스트는 HTML로, 이미지는 쿠팡 CDN 주소로 변환합니다. 쿠팡의 상품 등록·수정·임시저장은 수행하지 않습니다.

한 상품을 준비만 하려면 다음을 실행합니다.

```bash
python naver_to_coupang_html.py 13204504134 --prepare-only
```

여러 상품은 상품번호를 줄바꿈으로 적은 UTF-8 텍스트 파일을 준비한 뒤 실행합니다.

```bash
python naver_to_coupang_html.py --list product_numbers.txt --upload
```

처음 실행 시 열린 전용 쿠팡 WING 브라우저에서 로그인합니다. 로그인 세션은 로컬 `output/coupang-browser-profile`에만 저장됩니다. 완료된 상품별 `output/detail-preview/<상품번호>/coupang-paste.html` 파일의 내용을 쿠팡 **기본 등록 → HTML 작성**에 붙여 넣으면 됩니다.

기존 쿠팡 상품의 옵션별 이미지 상세를 공용 HTML로 바꾸는 시험 도구는 다음처럼 실행합니다. 기본 실행은 HTML을 화면에 채운 뒤 저장하지 않고 종료합니다.

```bash
python coupang_detail_replace.py 13204504134 16078736117
```

`--apply`는 화면을 확인한 뒤 `APPLY`를 한 번 더 입력해야만 **수정 및 검수 요청**을 보냅니다.

GUI의 **상세페이지** 탭에서도 같은 단건 작업을 할 수 있습니다. 최초에는 `쿠팡 로그인 연결/갱신`으로 WING에 직접 로그인하고, 이후에는 이 PC의 로컬 브라우저 세션을 재사용합니다. 쿠팡 ID·비밀번호는 프로그램에 저장하지 않습니다. 기존 상품 작업은 HTML 입력까지만 자동으로 수행하며, 저장은 열린 WING 화면에서 직접 진행합니다.

## 프로젝트 구조

```
easy-fulfill/
├── easy-fulfill.py      # 메인 프로그램 파일
├── google_sheets_oauth.py  # Sheets OAuth 공통 (GUI·database-sync)
├── database-sync.py     # 스프레드시트 동기화 (동일 google-oauth)
├── requirements.txt     # 필요한 패키지 목록
├── README.md           # 프로젝트 설명
├── google-oauth/       # Google OAuth (로컬 전용, .gitignore)
│   ├── credentials.json  # GCP에서 받은 클라이언트 JSON (수동 배치)
│   └── token.json        # 로그인 후 자동 생성
└── ui/                 # UI 파일 디렉토리
    └── main_window.ui  # Qt Designer UI 파일
```

`google-oauth` 안의 파일은 Git에 올리지 마세요. 설정 방법은 [docs/google-sheets-oauth-implementation.md](docs/google-sheets-oauth-implementation.md)를 참고하세요.
