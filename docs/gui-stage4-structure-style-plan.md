# GUI 개선 4단계 계획: 구조·스타일

작성: 2026-11-17, 기준 커밋 `996cf3a`(앱 버전 `2026.11.17.12`). 1단계(레이아웃 버그, `98790a7`)와 2단계(주문·발송 탭, `9a322a7`~`996cf3a`)는 끝났다. 3단계(알림·단축키)는 이 문서 범위가 아니다.

> **2026-11-18 갱신:** 4-2·4-3·4-4·4-6은 [창 크기와 「설정」 탭 개편 계획](gui-window-settings-plan.md)으로 바뀌었다. 4-2의 ⚙ 표시는 새 「설정」 탭 하나에만 붙이고, 4-3은 B-4(진단 항목을 「설정 › 프로그램」으로 이동), 4-4는 폐기(설정 탭 분류 구조로 대체), 4-6은 B-5(「설정 › DB 동기화」 페이지)로 진행한다. 4-1과 4-5는 이 문서대로 하되, 순서는 그 계획의 권장 순서를 따른다. 진행 상황은 그 계획의 「진행 상황」 표가 기준이다(4-1·4-2·4-5 완료). 아래 줄 번호와 버전 표기는 작성 당시 값이다.

## 목표

탭과 메뉴는 숨기지 않고 **구분만 분명히** 하고, 버튼과 상태 표시의 생김새를 **한 규칙으로 통일**한다. 작업자가 관리 탭을 임의로 건드릴 일은 없으므로 숨기기·잠금(관리자 모드, PIN, 개발자 모드)은 쓰지 않는다.

## 작업 순서

공통 스타일(4-1)을 먼저 한다. 나머지 작업이 모두 그 규칙(버튼 역할, 상태 색)을 쓰기 때문이다. 하위 작업마다 따로 커밋한다.

| 순서 | 작업 | 주요 파일 |
|---|---|---|
| 4-1 | 공통 스타일(버튼 역할, 상태 색) | 새 `shared/ui_style.py`, `easy-fulfill.py` `main()`, `ui/order_ship_panel.py` |
| 4-2 | 관리 탭 구분 표시 | `easy-fulfill.py` `_apply_tab_order` |
| 4-3 | 「API(개발 중)」 메뉴 정리 | `easy-fulfill.py` `setup_menubar`, `_setup_menus` |
| 4-4 | 관리자·환경설정 탭 2열 배치와 상태 배지 | `ui/main_window.ui`, `easy-fulfill.py` |
| 4-5 | 상세페이지 탭 정리 | `easy-fulfill.py` `_mount_detail_page_tab` |
| 4-6 | DB동기화 탭 정리 | `ui/main_window.ui`, `easy-fulfill.py`, `sheets/db_sheet_sync.py` |

줄 번호는 `996cf3a` 기준이다. 작업 중 바뀌므로 함수 이름으로 찾는다.

## 4-1. 공통 스타일

**→ 완료:** `55f07d6`(앱 `0.0.11`, 브랜치 `feat/window-settings`).

**현재:** 색과 버튼 모양이 `setStyleSheet` 호출 곳곳에 흩어져 있다(`easy-fulfill.py` 19곳, `ui/` 8곳). 같은 의미의 색도 제각각이다(오류 빨강 `#b00020`·`#c0392b`, 성공 초록 `#237804`·`#2e7d32`). 2단계의 강조 버튼은 `ui/order_ship_panel.py`의 `PRIMARY_BUTTON_STYLE`(파랑 `#1a73e8`)을 버튼마다 직접 넣는다.

**변경:**

- 새 모듈 `shared/ui_style.py`에 앱 전체 스타일시트 문자열과 도우미 함수를 둔다.
  - `set_button_role(button, role)`: `role`은 `"primary"`(그 화면의 다음 동작), `"danger"`(되돌리기 어려운 실제 동작), `None`(기본). 동적 속성을 바꾼 뒤 `style().unpolish()`/`polish()`로 다시 그린다.
  - `set_status(label, kind, text, detail="")`: `kind`는 `"ok"`·`"warn"`·`"error"`·`"idle"`. 앞에 `✓`·`!`·`✗`를 붙이고 색을 바꾼다. `detail`은 툴팁.
- 스타일시트는 **속성 선택자만** 쓴다(`QPushButton[role="primary"]`, `QLabel[status="error"]` 등). `QPushButton { … }`처럼 종류 전체에 거는 규칙은 쓰지 않는다. 전체 규칙은 Windows 기본 모양을 통째로 바꿔 모든 화면을 다시 확인해야 하기 때문이다.
- 색은 한 곳에서 정한다: 강조 `#1a73e8`(2단계 값 유지), 위험 `#c0392b`, 성공 `#2e7d32`, 주의 `#b26a00`, 대기 `#888888`.
- `main()`에서 `QApplication` 생성 직후 `app.setStyleSheet(APP_STYLESHEET)`를 한 번 호출한다.
- 옮길 것:
  - `ui/order_ship_panel.py`의 `PRIMARY_BUTTON_STYLE` 직접 지정 → `set_button_role(…, "primary")`
  - `tracking/controller.py` `_set_tracking_count_text`의 오류 색 → `set_status(…, "error", …)`
  - `easy-fulfill.py` 2061·2064행(상세 편집 경고 문구 색) → `set_status`
- 옮기지 않을 것: 시작·작업 중 오버레이(4760~4907행), 상세 편집기 미리보기·칩(1875·1892행), 메뉴바·상태바 테두리(5430·5699·5714행), `_set_status_label`의 끌어다 놓기 점선 테두리. 특수 용도라 이번 규칙과 무관하다.

**완료 기준:** 2단계 강조 버튼이 수정 전과 같은 색·모양으로 보이고, `PRIMARY_BUTTON_STYLE` 직접 지정이 없어진다. 다른 탭의 모양은 바뀌지 않는다(수정 전후 캡처 비교).

## 4-2. 관리 탭 구분 표시

**→ 대체·완료:** 「⚙ 설정」 탭 하나에만 적용했다(B-1, `d249e97`).

**현재:** 탭 순서는 `_apply_tab_order`가 정한다(`tab`, `tab_tracking`, `tab_quick_invoice`, `tab_documents`, `tab_detail_page`, `tab_db_sheet_sync`, `tab_admin`, `tab_2`). 관리용 탭 3개가 작업자용 탭과 같은 모양이다.

**변경:** 순서는 그대로 둔다. `_apply_tab_order` 끝에서 `tab_db_sheet_sync`·`tab_admin`·`tab_2`의 탭 글자 앞에 `⚙ `를 붙인다(이미 붙어 있으면 다시 붙이지 않음). 탭을 숨기거나 비활성화하지 않는다.

**완료 기준:** 세 탭에만 `⚙`가 보인다. 탭 글자로 탭을 찾는 코드가 있으면 objectName으로 찾도록 바꾼다(`grep -n "tabText" easy-fulfill.py`로 확인).

## 4-3. 「API(개발 중)」 메뉴 정리

**→ 대체됨:** 창 크기·설정 탭 계획 B-4.

**현재:** `_setup_menus`의 `m_api`에 항목 16개가 한 줄로 이어져 있고, 구분선은 2번째 뒤에 하나뿐이다. 읽기 전용 확인 동작과 실제 동작이 섞여 있다. 실제 동작에는 이미 확인 창이 있다(`dispatch_naver_via_api`, `_run_epost_receipt` 실제 모드, `run_epost_oz_print_dialog(execute_print=True)`의 `QMessageBox.question`).

**변경:** 메뉴를 숨기지 않고 아래처럼 나눈다. 항목 이름·동작·확인 창은 그대로 둔다.

| 묶음 | 항목(QAction) |
|---|---|
| 맨 위 | `act_api_load`(API로 주문 불러오기), `act_epost_portal_login`(포털 로그인 연결), `act_epost_print_targets`(출력 대상 확인) |
| 하위 메뉴 「확인·진단 (실제 동작 없음)」 | `act_epost_test_receipt`, `act_epost_portal_diagnostic`, `act_epost_portal_lookup`, `act_epost_portal_print_popup`, `act_epost_portal_oz_viewer`, `act_epost_portal_output_confirm`, `act_epost_portal_reprint_popup`, `act_epost_portal_reprint_oz_viewer`, `act_epost_oz_print_dialog` |
| 구분선 아래 「실제 동작」 | `act_dispatch`(API 발송처리), `act_epost_real_receipt`(실제 접수 확인), `act_epost_actual_print`(단건 실제 인쇄), `act_epost_batch_actual_print`(다건 실제 인쇄) |

- 「실제 동작」 네 항목은 아이콘을 위험 표시로 바꾼다. `image/`에 맞는 아이콘이 없으면 `QStyle.StandardPixmap.SP_MessageBoxWarning`을 쓴다.
- 「실제 동작」 위에 비활성 제목 항목(「실제 동작 — 실행 전 확인 창이 뜹니다」)을 둔다.
- 메뉴 이름 「API(개발 중)」는 아래 「정할 것」이 정해질 때까지 그대로 둔다.

**완료 기준:** 네 실제 동작이 확인 창 없이는 실행되지 않음을 테스트로 고정한다(`QMessageBox.question`을 `No`로 바꿔치기하면 API·포털 호출이 일어나지 않음).

## 4-4. 관리자·환경설정 탭 2열 배치와 상태 배지

**→ 폐기:** 창 크기·설정 탭 계획 B단계의 분류 구조로 대체한다. 상태 배지는 B-2에서 한다.

**현재:** 두 탭 모두 내용 위젯(`widget_admin_content`, `widget_settings_content`)의 최대 폭이 470px이라 기본 창(880px)에서 오른쪽 절반이 비어 있다. 상태 문구는 회색 글자다(`label_google_auth_status`, `label_settings_api_status`, `label_index_sheet_sync`, `label_app_version`).

**변경:**

- 관리자: 왼쪽 열은 「연동 · Google Sheets」와 「연동 · API 키」, 오른쪽 열은 송장추적 설정(`widget_admin_tracking`, `_build_admin_tracking_settings`가 채움).
- 환경설정: 왼쪽 열은 「주문번호 시작값」, 오른쪽 열은 「구매 금액별 사은품」과 「프로그램 버전」.
- 최대 폭 470px 제한을 풀고, 열마다 최대 폭(약 460px)을 둔다. 창 최소 폭이 늘어나지 않게 한다(좁으면 1열로 내려오지 않아도 됨. 기본 창 880px에서 2열이 들어가면 충분).
- 위 네 상태 문구는 4-1의 `set_status`로 표시한다(성공 ✓ 초록, 실패 ✗ 빨강, 확인 중 회색). 여러 줄 오류는 첫 줄만 보이고 원문은 툴팁(1단계와 같은 방식). 관련 코드: 3519, 4236~4252, 5054, 5763, 6275~6277행.

**완료 기준:** 기본 창 크기에서 두 탭 모두 2열로 보이고, 창 크기(880×758)와 최소 폭이 수정 전과 같다.

## 4-5. 상세페이지 탭 정리

**→ 완료:** 창 크기·설정 탭 계획의 마지막 단계로 진행했다(앱 `0.0.17`). 상세페이지 탭은 A-2(`4bb4417`)에서 이미 스크롤 영역에 넣었다. 「대상 상품」 줄은 ①과 ② 사이에 두었다.

**현재:** `_mount_detail_page_tab`에서 「네이버 상품번호」 입력이 두 곳이다. 「상품 상세 HTML 만들기」의 `lineEdit_detail_naver_create`(5908행)와 「쿠팡 상품 HTML 적용」의 `lineEdit_detail_naver_replace`(5928행)이다. 작업 로그가 탭 아래쪽을 크게 차지한다.

**변경:**

- 「네이버 상품번호」를 하나로 합친다. 두 묶음 위에 「대상 상품」 줄을 두고, `_detail_product_no(...)`를 부르는 곳(6174·6216·6221·6233·6242행)이 모두 그 칸을 읽게 한다. 쿠팡 상품번호 칸은 적용 묶음에 그대로 둔다.
- 묶음 제목을 단계로 바꾼다: 「① 쿠팡 WING 연결」, 「② 상품 상세 HTML 만들기」, 「③ 쿠팡 상품 HTML 적용」. 단계가 끝나면(로그인 연결 성공, HTML 생성 성공, WING에 채우기 성공) 제목 끝에 `✓`를 붙이고, 입력 초기화 때 지운다.
- 작업 로그는 접을 수 있게 한다(「작업 로그 ▾」 토글). 기본은 접힘, 오류가 기록되면 자동으로 펼친다. 로그 지우기·복사·저장 버튼은 그대로 둔다.

**완료 기준:** 상품번호를 한 번만 입력해 생성과 적용이 모두 동작한다(관련 테스트 `tests/test_detail_*`가 통과하고, 입력 칸 통합 테스트 추가).

## 4-6. DB동기화 탭 정리

**→ 대체됨:** 창 크기·설정 탭 계획 B-5(「설정 › DB 동기화」 페이지).

**현재:** 「실행」과 「새로고침」 버튼이 같은 모양이다. 「테스트 모드 (적용 안함)」(`checkBox_db_sync_preview`)가 켜져 있는지 눈에 잘 띄지 않는다. 파일 표시가 `네이버 DB [ — ] : (없음)` 형태다(`sheets/db_sheet_sync.py` `format_db_sync_label_line`).

**변경:**

- `pushButton_db_sync_run`을 `set_button_role(…, "primary")`로 강조한다.
- 테스트 모드가 켜져 있으면 버튼 줄 위에 주의 색 띠(「테스트 모드: 시트에 적용하지 않습니다」)를 보여 주고, 실행 버튼 글자를 「테스트 실행」으로 바꾼다. 꺼지면 띠를 숨기고 「실행」으로 되돌린다.
- 파일 표시를 `네이버 DB · Product_20261117.csv · 11-17 (오늘)`처럼 파일 이름, 수정 날짜, 며칠 전인지로 바꾼다. 파일이 없으면 `네이버 DB · 파일 없음`을 주의 색으로 보여 준다. `format_db_sync_label_line`을 고치고 테스트를 추가한다.

**완료 기준:** 테스트 모드를 켜고 끌 때 띠와 버튼 글자가 바뀐다. 파일 표시 함수 테스트가 오늘·n일 전·파일 없음을 다룬다.

## 하지 않을 것

- 탭이나 메뉴를 숨기거나 잠그는 것(관리자 모드, PIN, 개발자 모드).
- 단축키 변경, 팝업 줄이기, 상태바 고정 표시(3단계).
- 공유 시트 쓰기 방식 변경. 따라서 `MIN_COMPATIBLE_VERSION`은 올리지 않는다.
- 우체국 실제 접수·인쇄, 네이버 발송처리 API, 쿠팡 WING 저장 실행, 운영 Google 시트 쓰기.

## 검증

- 하위 작업마다 `QT_QPA_PLATFORM=offscreen`으로 `MainWindow`를 띄워 바뀐 탭을 수정 전후로 캡처(`grab()`)해 비교한다. 8개 탭을 오가도 창 크기가 880×758로 유지되는지 확인한다.
- 전체 테스트 `python -m unittest discover`가 수정 전과 같은 결과인지 확인한다.
- Windows 실제 화면(배율 100%·125%·150%)에서 `⚙` 글자, 위험 아이콘, 2열 배치를 확인한다.

## 진행 규칙

`AGENTS.md`를 따른다. 하위 작업 커밋마다 `APP_VERSION`을 올리고, `docs/known-risks.md` 「작업별 기록」에 「E. 구조·스타일」 절을 만들어 커밋·내용을 적고, 커밋 후 푸시한다. 2단계처럼 작업 브랜치(`feat/structure-style`)에서 진행하고, 한 PC에서 확인한 뒤 `main`에 합친다.

## 정할 것

- **「API(개발 중)」 메뉴 이름:** 실제 접수·인쇄를 지금 업무에 쓰고 있으면 「우체국·API」 같은 정식 이름으로 바꾼다. 정해지기 전에는 이름을 바꾸지 않는다.
