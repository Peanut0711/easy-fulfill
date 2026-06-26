# 미답변 문의 알림 (네이버·쿠팡 → 슬랙)

작성일: 2026-06-26

## 배경

배송추적 탭에서 네이버 상품·고객문의 미답변 건을 슬랙으로 알리는 기능이 있었으나, 원래는
**"새 문의 1회 알림 + 첫 실행 시드(기존 건 침묵)"** 모델이었다. 실제 운영 상황이 아래와 같아
모델 자체가 맞지 않았다.

- 작업자 **최대 4~5명**이 각자 PC에서 앱을 켜둠(아침 출근~퇴근, 종종 껐다 켬).
- 미답변 문의는 **누군가 처리할 때까지 주기적으로** 환기돼야 함(한 번 알리고 끝이면 방치됨).

그래서 알림 모델을 **"미답변 리마인더"** 로 교체하고, 같은 구조에 **쿠팡 WING OpenAPI** 문의를
합류시켰다.

## 알림 모델 (핵심)

- **새 미답변 → 즉시 1회 알림(🆕).**
- **아직 미답변인 건 → 「최근알림시각」에서 60분(`NAVER_INQUIRY_REMIND_MIN`) 경과 시 다시 알림(🔁).**
- **누군가 답변하면 자동 중단** — 매 폴링마다 각 API에서 *미답변 전체*를 다시 받아오므로, 답변된
  건은 다음 조회 목록에서 빠져 더는 알리지 않는다(별도의 "처리자 추적" 불필요. 네이버
  `answered=False`, 쿠팡 `NOANSWER`/`NO_ANSWER` 가 진실원천).

### 근무시간 게이트
- 알림은 **평일 10:00~19:00** 에만 발송(`_inquiry_alerts_allowed`).
- 그 외 시간엔 발송도 「최근알림시각」 갱신도 하지 않는다 → 새벽/주말에 들어온 미답변은
  **다음 근무 시작(10시)에 한 번에 환기**된다.

### 다중 PC 중복 방지
- 「최근알림시각」을 **공유 시트**에 저장하므로, 60분 리마인더 창 안에서는 **먼저 조회한 1대만**
  발송하고 나머지 PC는 "방금 알림됨"을 보고 건너뛴다. → **4~5대 모두 켜둬도 됨**.
- 잔여 경쟁: 브랜드뉴 문의가 여러 PC에 동시(수 초 내) 도착하면 중복 발송 가능. PC별 가동시각
  차이로 자연 분산되어 실사용상 드묾(완전 차단이 필요하면 향후 시트 클레임 락 도입 — TODO).

## 시트 스키마 「문의알림」

9열: `문의ID | 유형 | 등록일시 | 대상 | 작성자 | 내용 | 감지시각 | 최근알림시각 | 상태`

- 옛 8열(`…|알림`) 시트는 첫 조회 시 자동 마이그레이션(`add_cols`로 9열 확장 후 헤더 교체).
  옛 「알림」 값(기준/전송)은 「최근알림시각」 자리에서 파싱 불가→미알림 취급되어 첫 근무시간에
  한 번 환기됨(의도된 동작).

### 문의ID 접두사 (채널 구분·충돌 방지)
| 채널 | 접두사 | 예 |
|------|--------|-----|
| 네이버 상품문의 | `Q` | `Q12345` |
| 네이버 고객문의 | `C` | `C67890` |
| 쿠팡 온라인문의(상품) | `KO` | `KO39352323` |
| 쿠팡 콜센터(CS)문의 | `KC` | `KC1015668177` |

모든 채널이 같은 「문의알림」 시트·같은 슬랙 웹훅(`slack_webhook_url`, 배송 위험 알림과 통합)으로
합류한다.

## API 구현

### 네이버 — `naver_commerce.py`
- 인증: OAuth 토큰(`POST /v1/oauth2/token`, bcrypt 전자서명).
- 조회: 상품문의 `GET /v1/contents/qnas`, 고객문의 `GET /v1/pay-user/inquiries` (`answered=false`).

### 쿠팡 — `coupang_commerce.py` (신규)
- **인증이 네이버와 다름**: 토큰이 아니라 **요청마다 HMAC 서명**(CEA HmacSHA256). 표준
  `hmac`/`hashlib` 만 사용 → **새 의존성 없음**.
  - `signed_date` = GMT, 포맷 `yyMMdd'T'HHmmss'Z'`
  - `message` = `signed_date + method + path + query` (query는 `?` 없이)
  - `signature` = `hex(HMAC-SHA256(secret_key, message))`
  - 헤더 `Authorization: CEA algorithm=HmacSHA256, access-key={ak}, signed-date={date}, signature={sig}`
  - **주의**: 서명에 쓰는 query 문자열과 실제 요청 URL 의 query 가 완전히 동일해야 함(인코딩·순서).
    → params 를 직접 `urlencode` 해 URL 에 붙이고 같은 문자열로 서명한다.
- 베이스: `https://api-gateway.coupang.com`, 경로 접두 `/v2/providers/openapi/apis/api/v5/vendors/{vendorId}`
- 조회:
  - 온라인문의 `GET …/onlineInquiries?answeredType=NOANSWER`
  - 콜센터문의 `GET …/callCenterInquiries?partnerCounselingStatus=NO_ANSWER`
- **날짜 범위 최대 ~7일 제한** → `_date_windows`로 30일 조회기간을 7일씩 끊어 호출. 페이지네이션은
  `pageNum` 증가(페이지 항목 수 < `PAGE_SIZE` 면 종료), `MAX_PAGES=20` 안전장치.

## 키 관리 / 설정 UI

- 모든 비밀키는 **비공개 「설정」 탭에만 저장**(코드/레포 금지 — 퍼블릭 레포). 우체국 regkey·네이버
  키와 동일 패턴.
- config keys: `naver_client_id`, `naver_client_secret`, `coupang_vendor_id`,
  `coupang_access_key`, `coupang_secret_key`.
- 설정 팝업 그룹 "미답변 문의 알림 (네이버·쿠팡 → 슬랙)":
  - **[네이버 키]** — client_id/secret 입력→토큰 발급 검증→저장.
  - **[쿠팡 키]** — vendorId/accessKey/secretKey 3종 입력→1일치 조회 검증→저장.
  - **[지금 확인]** — 토글과 무관하게 즉시 1회 조회.
  - 체크박스 "이 PC에서 자동 조회(5분)" — PC별 `app_settings.naver_inquiry_notify`.
- 한쪽(네이버 or 쿠팡) 키만 등록돼 있어도 동작. 둘 다 없을 때만 에러.

## 구현 위치 (코드)

- `coupang_commerce.py` — 쿠팡 HMAC·조회 순수함수(신규).
- `naver_commerce.py` — 네이버 토큰·조회 순수함수(기존).
- `easy-fulfill.py`:
  - `run_naver_inquiry_poll_worker()` — 네이버+쿠팡 통합 조회·리마인더·근무시간 게이트(이름은
    네이버지만 쿠팡도 처리. 향후 `monitoring.py`로 추출 예정 — TODO).
  - `NaverInquiryPollThread`, `CoupangCredsValidateThread`(신규), `NaverCredsValidateThread`.
  - 설정 팝업(`_open_tracking_settings_dialog`)·`_on_coupang_creds_edit_clicked`/`_validated`.
  - 상수: `NAVER_INQUIRY_REMIND_MIN=60`, `LOOKBACK_DAYS=30`, 근무시간 10~19시, `CONFIG_KEY_COUPANG_*`.

## 결정사항 (사용자 확정)

- 알림 형태: **새 건 즉시 + 60분 주기 리마인더**.
- 시간대: **평일 10:00~19:00** 코드 제한.
- 이번 범위: **미답변 문의만**(쿠팡 배송/주문 API는 다음 단계 — TODO).
- 통합 채널: 슬랙은 기존 `slack_webhook_url` 재사용(배송 위험 알림과 같은 채널).
