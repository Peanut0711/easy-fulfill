# Google Sheets 인증 정보 보안 개선 가이드

> **업데이트:** 메인 프로그램 `easy-fulfill.py`는 **OAuth(사용자별 로그인 + 로컬 `token.json`)** 로 전환되었습니다.  
> 구현·사용 절차는 [google-sheets-oauth-implementation.md](./google-sheets-oauth-implementation.md)를 참고하세요.  
> 아래 "과거 방식" 설명은 여전히 유효한 **비교·리스크 설명**용입니다.

## 과거 방식 정리 (Service Account 키 공유)

이전에는 `easy-fulfill.py`가 아래와 유사한 방식으로 스프레드시트에 접근했습니다.

- `gspread.service_account(filename=...)` 사용
- 일반 `API Key`가 아니라 **Service Account 키(JSON 파일)** 를 클라이언트 PC에 배포
- 팀원 여러 명이 같은 키 파일을 복사해 사용

이 방식은 단기 테스트에는 편하지만, 장기 운영 시 아래 리스크가 있습니다.

- 키 파일 유출 시 제3자도 동일 권한으로 시트 접근 가능
- 누가 어떤 요청을 했는지 사용자 단위 추적이 어려움
- 키를 교체(로테이션)할 때 모든 PC 재배포 필요

---

## 요구사항 기준 권장 방향

요구사항:

1. 깃에는 키를 올리지 않음 (이미 잘 하고 있음)
2. 사용자 수는 4명
3. 키를 오픈하고 싶지 않음

위 조건이라면, 아래 우선순위를 권장합니다.

### 1순위(권장): 사용자별 OAuth(Installed App) 전환

핵심 아이디어:

- 공용 Service Account 키를 배포하지 않고,
- 각 팀원이 본인 Google 계정으로 최초 1회 로그인 후 토큰 발급
- 이후 로컬 토큰으로 접근

장점:

- 공용 비밀키 배포 제거
- 계정 단위 접근 제어 가능(팀원 추가/제거 쉬움)
- 감사 추적(누가 접근했는지) 개선

주의:

- 최초 로그인 절차가 필요
- 토큰·클라이언트 파일(`google-oauth/token.json`, `credentials.json`)은 저장소에 커밋하지 말고 PC 로컬에서 보호

적합성:

- 사용자 4명 규모에 특히 적합
- 장기 운영 기준에서 가장 현실적인 개선안

---

## 대안 비교

### A. 현재 방식 유지 + 운영 보강(임시)

방법:

- Service Account 키 파일은 계속 사용
- 배포 방식만 강화
  - 개인 메신저 전송 금지
  - 사내 암호화 저장소(예: 비밀관리 툴, 접근권한 폴더)에서만 배포
  - 정기 키 로테이션(예: 월 1회/분기 1회)
  - 퇴사/이동자 발생 시 즉시 키 폐기

평가:

- 구현이 가장 쉬움
- 하지만 "공용 비밀키 공유" 구조는 그대로라 근본 해결은 아님

### B. OAuth 사용자별 토큰(권장)

방법:

- Google Cloud에서 OAuth Client(Desktop) 생성
- 앱 실행 시 사용자 로그인 플로우 수행
- 로컬에 사용자별 토큰 저장

평가:

- 보안/운영 균형이 가장 좋음
- 4명 팀에서 관리 부담도 낮음

### C. 사내 중계 서버(Proxy) 추가

방법:

- 데스크톱 앱은 내부 API 서버만 호출
- Google 인증은 서버에서만 처리

평가:

- 보안은 가장 강력
- 대신 서버 운영/장애대응/배포비용이 필요
- 현재 규모(4명)에서는 과할 수 있음

---

## 권장 실행안 (OAuth 전환)

## 0) 기본 원칙

- 인증 파일은 절대 Git에 커밋 금지
- 사용자 로컬 저장 시 권한 제한(본인 계정만 읽기)
- 토큰 파일 유출 대비 만료/폐기 절차 준비

## 1) Google Cloud 설정

- Google Sheets API 활성화
- OAuth consent screen 설정(내부 사용자 또는 테스트 사용자 4명 등록)
- OAuth Client ID 생성(Desktop App)
- `credentials.json` 다운로드

## 2) 코드 구조 변경

현재:

- `gspread.service_account(filename=CREDENTIAL_PATH)`

변경:

- `google-auth-oauthlib` 기반 사용자 인증
- 최초 로그인 시 브라우저 인증, 이후 `token.json` 재사용

개념 예시:

```python
from pathlib import Path
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # database-sync 쓰기와 동일 토큰 공유 시
APP_ROOT = Path(__file__).resolve().parent  # easy-fulfill.py 기준 저장소 루트
GOOGLE_OAUTH_DIR = APP_ROOT / "google-oauth"
CREDENTIALS_FILE = GOOGLE_OAUTH_DIR / "credentials.json"
TOKEN_FILE = GOOGLE_OAUTH_DIR / "token.json"

def get_gspread_client():
    GOOGLE_OAUTH_DIR.mkdir(parents=True, exist_ok=True)
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return gspread.authorize(creds)
```

## 3) 배포 방식

- 프로그램 배포 시 `credentials.json`은 설치 패키지에 하드코딩하지 않음
- 최초 실행 안내 문서 제공:
  - 프로그램 루트(`easy-fulfill.py`와 같은 폴더) 아래 `google-oauth/credentials.json`에 클라이언트 JSON 배치
  - 앱 실행 후 Google 로그인

## 4) 운영 정책

- 사용자 추가/삭제는 Google 시트 공유 권한으로 관리
- 분기별 토큰/접근권한 점검
- 문제 발생 시 특정 사용자 토큰만 폐기 가능

---

## 단기-중기 마이그레이션 플랜

1. **단기(오늘)**: 기존 Service Account 키 재발급 + 구키 폐기  
2. **단기(1~2일)**: OAuth 로그인 코드 추가, 내부 1명 파일럿  
3. **중기(이번 주)**: 팀원 4명 전환, Service Account 키 배포 중단  
4. **중기(다음 주)**: 토큰/권한 관리 체크리스트 운영

---

## 체크리스트

- [ ] `.gitignore`에 `google-oauth/` 등 인증 관련 경로 포함
- [ ] 코드 내 인증 경로 하드코딩 제거
- [ ] 스코프는 `google_sheets_oauth.py` 기준 (`spreadsheets` 전체; sync·GUI 공통 토큰)
- [ ] 사용자별 계정 접근권한 최소화
- [ ] 키/토큰 유출 대응 절차 문서화

---

## 결론

현재 조건(4명, 키 비공개, 장기 운영)에서는 **공용 Service Account 키 배포를 중단하고 사용자별 OAuth로 전환**하는 것이 가장 적절합니다.  
보안 수준을 높이면서도 운영 복잡도는 크게 늘리지 않는 현실적인 선택입니다.
