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
- **예외 — 쿠팡 CS문의(`KC`)는 하루 1회만 재알림(🔁).** 자동 반품수거처럼 회사 도착까지
  판매자가 액션할 수 없는 건이 많아 매시간 알림이 소음이 되어, 「최근알림시각」의 날짜가
  오늘 이전일 때만 재알림한다(그날 이미 알렸으면 스킵 → 다음 근무일 1회). 신규 즉시 알림과
  자동 중단은 다른 채널과 동일.
- **누군가 답변하면 자동 중단** — 매 폴링마다 각 API에서 *미답변 전체*를 다시 받아오므로, 답변된
  건은 다음 조회 목록에서 빠져 더는 알리지 않는다(별도의 "처리자 추적" 불필요. 네이버
  `answered=False`, 쿠팡 `NOANSWER`/`NO_ANSWER` 가 진실원천).

### 근무시간 게이트 (사용자 설정 가능)
- 알림은 **평일 시작시각~종료시각** 에만 발송(`_inquiry_alerts_allowed(now, start, end)`). 기본 10~19시.
- **「관리자」 탭 "미답변 문의 알림" 그룹의 "알림 시간대(평일)" 스핀박스**로 시작/종료 시각을 바꿀 수
  있고, 공유 「설정」 탭(`inquiry_work_start_hour`/`inquiry_work_end_hour`)에 저장돼 **전원 적용**된다.
  종료는 시작보다 최소 1시간 뒤로 자동 보정. 워커는 매 폴링마다 공유 cfg에서 시각을 읽으므로
  한 명이 바꾸면 다음 폴링(최대 5분)에 모든 PC에 반영.
- 그 외 시간엔 발송도 「최근알림시각」 갱신도 하지 않는다 → 새벽/주말에 들어온 미답변은
  **다음 근무 시작 시각에 한 번에 환기**된다.

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

### 네이버 — `integrations/naver_commerce.py`
- 인증: OAuth 토큰(`POST /v1/oauth2/token`, bcrypt 전자서명).
- 조회: 상품문의 `GET /v1/contents/qnas`, 고객문의 `GET /v1/pay-user/inquiries` (`answered=false`).

### 쿠팡 — `integrations/coupang_commerce.py` (신규)
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

## 문의 탭 (조회 전용, `0.0.36`)

- 측면 툴바 「문의」(탭 이름 「네이버 문의」, Ctrl+7). `ui/main/inquiries.py` `InquiriesMixin`,
  작업 스레드 `NaverInquiryListThread` → `run_naver_inquiry_list_worker(days, unanswered_only)`
  → `naver_commerce.fetch_inquiry_list`.
- 위 알림과 별개로 API를 직접 조회해 보여 준다. 「문의알림」 시트·슬랙에는 읽지도 쓰지도 않는다.
- 기간 오늘(0시부터)·최근 7일·최근 30일(기본 오늘, `0.0.39`부터. 그 전에는 7·30일, 기본 30일), 「미답변만」(기본 켬). 처음 들어갈 때 한 번 자동 조회, 이후 「새로고침」·필터 변경 때
  조회. 한 종류(상품문의·고객문의)가 실패해도 다른 종류는 보여 주고 상태 문구에 「일부 실패」로 표시.
- 응답 필드(2026-11-22 실응답, 최근 30일 상품문의 25·고객문의 17건, 읽기 전용 확인):
  - 상품문의: `questionId, question, answer, answered, answers[{answer, createDate}], createDate,
    maskedWriterId, productId, productName`. 답변 시각은 `answers`의 `createDate`에만 있다.
  - 고객문의: `inquiryNo, title, inquiryContent, category, inquiryRegistrationDateTime, answered,
    answerContent, answerContentId, answerRegistrationDateTime, answerTemplateNo, customerId, customerName,
    orderId, productNo, productName, productOrderIdList, productOrderOption`.
  - 조회할 때마다 종류별 첫 항목의 키 목록을 「도움말 › 프로그램 로그」에 남긴다(내용은 남기지 않음).
- 고객문의 주문·배송(`0.0.37`): 고객문의를 고르면 `NaverInquiryOrderThread` →
  `run_naver_inquiry_order_worker` → `naver_commerce.fetch_inquiry_order_summary`(주문번호 →
  상품주문번호 목록 → 상세)로 같은 주문의 상품주문을 모두 읽어 상세 아래 「주문·배송」에 붙인다.
  문의 대상(`productOrderIdList`, 실응답은 상품주문번호 문자열 1개)을 맨 위에 둔다.
  - 보이는 것: 주문일, 상품·옵션·수량, 상태(결제완료면 발주확인 여부), 클레임(종류·상태·사유,
    수거·재배송 송장), 택배사·송장·배송상태·발송·도착 시각(송장번호 오류 표시 포함), 아직 발송 전이면
    발송기한, 수령인·주소·배송메모. 연락처는 넣지 않는다.
  - 고른 주문만 부르고 주문번호별로 기억한다. 「새로고침」하면 비운다. 실패한 주문은 다시 고르면 다시 부른다.
    불러오는 중에 다른 문의를 고르면 끝난 뒤 마지막으로 고른 것만 부른다.
  - 문의 탭 작업자끼리 네이버 토큰을 30분 동안 같이 써서 고를 때마다 공유 「설정」을 읽지 않는다.
    토큰이 거절되면(HTTP 401·403) 한 번만 새로 받아 다시 조회한다.
  - 상태·클레임·사유·택배사 코드는 한글로 바꾸고, 모르는 코드는 원래 값을 그대로 보인다.
  - 실응답 확인(2026-11-22, 읽기 전용, 상태 값과 칸 채움 여부만): 최근 30일 고객문의 17건 모두 조회 성공,
    상품주문 43건(1~7개/주문), 약 0.5초/건.
- 측면 툴바 미답변 배지(`0.0.38`): `run_naver_inquiry_count_worker`가 최근 `NAVER_INQUIRY_LOOKBACK_DAYS`(30)일
  네이버 미답변 상품문의·고객문의 수를 센다. 시작 15초 뒤 한 번, 그 뒤 5분마다(`_inquiry_badge_timer`, 종료 때 멈춤).
  - 위 슬랙 알림 조회를 재사용하지 않는 이유: 알림 조회는 「이 PC에서 자동 조회」를 켠 PC에서만 돌고(기본 꺼짐),
    한 주기에 한 PC만 실제로 조회하며(나머지는 `skipped_other_pc`), 쿠팡 건수도 섞여 있다.
  - PC마다 5분에 네이버 호출 2번. 토큰은 문의 탭 작업자와 같이 써서(30분) 공유 「설정」 읽기는 토큰이 없거나
    거절될 때뿐이다. 두 종류 모두 401·403이면 새 토큰으로 한 번 더 센다.
  - 문의 탭을 「미답변만」·최근 30일로 빠짐없이 불러오면 그 건수로 배지를 바로 맞춘다(7일·전체·일부 실패면 그대로).
  - 셀 수 없으면(키 없음·네트워크·두 종류 모두 실패) 틀린 숫자를 남기지 않게 배지를 떼고 툴팁에
    「미답변 수 확인 실패 · 시각」을 보이며 프로그램 로그에 원인을 남긴다. 한 종류만 실패하면 센 만큼 보이고 「(일부 실패)」.
  - 실제 API(읽기 전용, 2026-11-22): 첫 조회 1.9초(설정 읽기·토큰 포함), 토큰 재사용 때 0.4초, 미답변 0건.
- 상세 화면 디자인(`0.0.40`): 유형·답변 상태 배지와 큰 제목, 문의(파랑)·답변(초록, 미답변 빨강) 막대 카드, 주문·배송 상품별 카드(막대 색 = 상태: 발송 전 주황·배송 중 하늘·완료 초록·취소/반품/교환 빨강, 송장번호 굵게), 받는 분 표. `ui/main/inquiries.py`의 `inquiry_detail_html`·`inquiry_order_html`.
- 답변 등록(`0.0.41`): 미답변 문의를 고르면 상세 아래 「답변 작성」 카드(템플릿 넣기·입력란·「답변 등록」).
  - API: 상품문의 `PUT /v1/contents/qnas/{questionId}` 본문 `{"commentContent": 답변}`(네이버 커머스API 커뮤니티
    사례로 확인), 고객문의 `POST /v1/pay-merchant/inquiries/{inquiryNo}/answer` 본문 `{"answerComment": 답변}`
    (**필드 이름 미확인** — 문서 사이트 404, 틀리면 400으로 거절되고 등록되지 않는다).
    템플릿 `GET /v1/contents/qnas/templates` → `[{subject, content, questionType}]`.
  - 권한 확인(2026-11-22, 아무것도 만들지 않는 호출): 템플릿 200(2개), 없는 번호 0에 빈 본문 → 상품문의 PUT 404
    「데이터를 찾을 수 없습니다」, 고객문의 POST 400 `ERR-NC-101004` 「문의 번호 0에 해당하는 문의가 없습니다」.
    둘 다 403이 아니라 권한은 있다.
  - 안전장치: 확인 창(기본 「취소」, 공개 범위·미리보기), 보내기 직전 `find_inquiry`로 다시 읽어 이미 답변됐거나
    못 찾으면 보내지 않음(여러 PC 동시 답변 방지), 2xx면 응답 본문이 비거나 JSON이 아니어도 성공(성공을 실패로
    보고 다시 보내 두 번 붙는 것 방지), 401이면 새 토큰으로 다시 확인부터, 보내는 중 연결이 끊기면 「등록됐는지
    알 수 없음」으로 알리고 새로고침해 확인하게 함. 실패하면 초안을 지킨다.
  - 등록한 답변 수정(고객문의 `PUT …/answer/{answerContentId}`, 상품문의 같은 PUT)은 아직 없다.

## 키 관리 / 설정 UI

- 모든 비밀키는 **비공개 「설정」 탭에만 저장**(코드/레포 금지 — 퍼블릭 레포). 우체국 regkey·네이버
  키와 동일 패턴.
- config keys: `naver_client_id`, `naver_client_secret`, `coupang_vendor_id`,
  `coupang_access_key`, `coupang_secret_key`.
- 「관리자」 탭 그룹 "미답변 문의 알림 (네이버·쿠팡 → 슬랙)":
  - **[네이버 키]** — client_id/secret 입력→토큰 발급 검증→저장.
  - **[쿠팡 키]** — vendorId/accessKey/secretKey 3종 입력→1일치 조회 검증→저장.
  - **[지금 확인]** — 토글과 무관하게 즉시 1회 조회.
  - 체크박스 "이 PC에서 자동 조회(5분)" — PC별 `app_settings.naver_inquiry_notify`.
- 한쪽(네이버 or 쿠팡) 키만 등록돼 있어도 동작. 둘 다 없을 때만 에러.

## 구현 위치 (코드)

- `integrations/coupang_commerce.py` — 쿠팡 HMAC·조회 순수함수(신규).
- `integrations/naver_commerce.py` — 네이버 토큰·조회 순수함수(기존).
- `easy-fulfill.py`:
  - `run_naver_inquiry_poll_worker()` — 네이버+쿠팡 통합 조회·리마인더·근무시간 게이트(이름은
    네이버지만 쿠팡도 처리. 향후 `monitoring.py`로 추출 예정 — TODO).
  - `NaverInquiryPollThread`, `CoupangCredsValidateThread`(신규), `NaverCredsValidateThread`.
  - 「관리자」 탭(`_build_admin_tracking_settings`)·`_on_coupang_creds_edit_clicked`/`_validated`.
  - 상수: `NAVER_INQUIRY_REMIND_MIN=60`, `LOOKBACK_DAYS=30`, 근무시간 10~19시, `CONFIG_KEY_COUPANG_*`.

## 결정사항 (사용자 확정)

- 알림 형태: **새 건 즉시 + 60분 주기 리마인더**. 단 **쿠팡 CS문의(KC)는 하루 1회 리마인더**.
- 시간대: **평일 10:00~19:00** 코드 제한.
- 이번 범위: **미답변 문의만**(쿠팡 배송/주문 API는 다음 단계 — TODO).
- 통합 채널: 슬랙은 기존 `slack_webhook_url` 재사용(배송 위험 알림과 같은 채널).
