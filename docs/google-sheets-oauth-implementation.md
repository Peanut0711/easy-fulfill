# Google Sheets OAuth 구현 및 사용 가이드 (Easy Fulfill)

배경·전략은 [google-sheets-auth-hardening.md](./google-sheets-auth-hardening.md)를 참고하고, 여기서는 **메인 앱(`easy-fulfill.py`)에 반영된 구현**과 **팀원 사용 방법**, **UI**를 정리합니다.

---

## 1. 개념: “1회 로그인 후 로컬 토큰”이란?

### 예전 방식 (Service Account 공유)

- 모든 PC에 **동일한 서비스 계정 JSON 키**를 복사해 둠.
- 키가 곧 “비밀번호”. 한 번 유출되면 누구나 같은 권한으로 접근 가능.

### 지금 방식 (OAuth, 데스크톱 앱)

1. Google Cloud에서 만든 **OAuth 클라이언트 설정**만 앱과 함께 배포합니다. 이 파일(`credentials.json`)은 “앱 식별자 + 로그인 창 띄우는 데 필요한 정보”에 가깝고, **혼자만으로는 시트에 바로 접근할 수 없습니다.**
2. 팀원이 프로그램을 쓰면서 스프레드시트를 처음 읽을 때, **본인 Google 계정으로 브라우저에서 한 번 승인**합니다.
3. Google이 **짧게 쓰는 액세스 토큰**과 **나중에 갱신할 때 쓰는 리프레시 토큰** 등을 앱에 돌려줍니다.
4. 앱은 이걸 **프로젝트 루트 아래 `google-oauth/`** 폴더의 `token.json`에 저장합니다.
5. 다음 실행부터는 **저장된 토큰으로 자동 접근**합니다. 만료되면 리프레시로 갱신하고, 안 되면 그때만 다시 브라우저 로그인이 뜹니다.

정리하면, **공용 비밀키를 팀에 나눠 주지 않고**, **각자 PC·각자 계정 단위로** 접근 권한을 관리합니다. 시트는 여전히 **해당 Google 계정에 공유**되어 있어야 읽을 수 있습니다.

---

## 2. 코드에서 한 일 (요약)

| 항목 | 내용 |
|------|------|
| 인증 방식 | `gspread.service_account(...)` 제거 → OAuth `InstalledAppFlow` + `gspread.authorize(creds)` |
| 클라이언트 시크릿 경로 | 저장소 루트( `easy-fulfill.py`와 같은 디렉터리) 기준 `google-oauth/credentials.json` |
| 토큰 저장 | 같은 폴더의 `token.json` |
| 스코프 | `https://www.googleapis.com/auth/spreadsheets.readonly` (매핑 조회용 읽기 전용) |
| 클라이언트 캐시 | `MainWindow`에서 `_get_gspread_client()`로 한 번 만든 gspread 클라이언트 재사용 |

의존성: `requirements.txt`에 `google-auth-oauthlib` 추가.

> 참고: 저장소 내 `database-sync.py`, `api-key-test.py` 등은 여전히 서비스 계정 경로를 쓸 수 있습니다. 메인 GUI 앱만 OAuth로 전환된 상태입니다.

---

## 3. 팀원 설치·사용 절차

### 3.1 패키지

```bash
pip install -r requirements.txt
```

### 3.2 Google Cloud 콘솔 (관리자 1회)

1. 프로젝트에서 **Google Sheets API** 사용 설정.
2. **OAuth 동의 화면** 구성 (내부 조직이면 Internal, 아니면 테스트 사용자에 팀원 4명 이메일 등록).
3. **OAuth 클라이언트** 유형: **데스크톱 앱**.
4. 클라이언트 JSON 다운로드 → 파일 이름을 `credentials.json`으로 맞춤.

### 3.3 각 PC에 파일 배치

**클론한 저장소 루트**에 `google-oauth` 폴더를 두고, 그 안에 `credentials.json`만 넣습니다 (Git에 넣지 않음 — `.gitignore`에 `google-oauth/` 포함).

- 예: `...\easy-fulfill\google-oauth\credentials.json`
- `google-oauth` 폴더가 없으면 만들어도 되고, 앱이 인증 시도 시 생성하기도 합니다.

### 3.4 실행

```bash
python easy-fulfill.py
```

네이버/쿠팡 처리 등 **스프레드시트 매핑을 처음 읽는 시점**에:

- `credentials.json`이 없으면: 파일 경로 안내와 함께 오류.
- 있으면: 브라우저가 열리고 Google 로그인·권한 승인 → 이후 `token.json` 생성.

### 3.5 시트 공유

OAuth로 로그인한 **그 Google 계정**이 해당 스프레드시트에 **최소 읽기 권한** 이상으로 공유되어 있어야 합니다.

---

## 4. UI는 꼭 넣어야 하나?

**아니요.** 지금 구현만으로도 브라우저 기반 1회(또는 토큰 만료 시) 로그인은 동작합니다.

다만 사무실 배포 시 **다음은 있으면 사용성이 좋습니다** (선택).

| UI·기능 | 목적 |
|---------|------|
| 상태 표시 | “Google: 연결됨 / 미설정” 등 `token.json` 존재 여부 표시 |
| 안내 다이얼로그 | `credentials.json` 없을 때 복사 가능한 경로·설치 순서 표시 |
| “Google 다시 연결” | `token.json` 삭제 후 `_gspread_client = None`으로 재인증 유도 |
| 오류 메시지 정리 | 리다이렉트 URI, 테스트 사용자 미등록 등 흔한 실패 원인 안내 |

필요하면 메뉴/설정 창에 위 항목을 추가하는 것을 권장합니다.

---

## 5. 문제 해결 (자주 나오는 경우)

| 증상 | 확인 |
|------|------|
| `credentials.json` 없음 | `google-oauth/` 경로·파일명 오타 (저장소 루트 기준) |
| 브라우저는 뜨는데 승인 후 실패 | OAuth 클라이언트가 **데스크톱**인지, 동의 화면에 사용자 등록됐는지 |
| 다른 계정으로 바꾸고 싶음 | 해당 PC에서 `token.json` 삭제 후 앱 다시 실행 |
| 권한 없음 | 시트를 로그인한 계정 이메일과 공유했는지 |

---

## 6. 관련 파일

- 앱: `easy-fulfill.py`
- 의존성: `requirements.txt`
- 전략·대안 비교: [google-sheets-auth-hardening.md](./google-sheets-auth-hardening.md)
- 예전 방식 vs OAuth (초심자용 설명): [google-sheets-old-vs-oauth-explained.md](./google-sheets-old-vs-oauth-explained.md)
