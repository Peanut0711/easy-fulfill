# easy-fulfill.py 기능 분리 계획서

## 목적

`easy-fulfill.py`에 집중된 화면 제어, 도메인 규칙, 파일 처리, Google Sheets I/O를 기능 단위로 단계적으로 분리한다.

전면 재작성은 하지 않는다. 각 단계에서 현재 UI, Sheet 스키마, Excel 출력, 외부 API 동작을 보존하고 검증한 뒤 다음 단계로 진행한다.

## 현재 기준선

- 대상 파일: `easy-fulfill.py`
- 전체: 약 10,500줄
- `MainWindow`: 약 7,470줄, 메서드 267개
- 포함 기능: 배송추적, 주문번호 인덱스, Quick Excel, 마켓별 주문 Excel import, 송장 처리, 상품분류, 상세페이지, DB-Sheet 동기화, OAuth 및 공통 UI

## 목표 구조

```text
easy-fulfill.py                 # 실행 진입점과 의존성 조립
app/
  main_window.py                # 공통 화면 조립, 메뉴, 상태바
  settings.py                   # PC별 앱 설정
shared/
  google_sheets.py              # OAuth client 및 공통 Sheet 접근
  workers.py                    # 공통 비동기 작업 도구
tracking/
  service.py                    # 배송 상태, 위험 판정, 알림 정책
  repository.py                 # 배송추적/설정/문의 Sheet I/O
  workers.py                    # 배송추적 비동기 작업
  controller.py                 # 배송추적 탭 UI 제어
order_index/
  service.py                    # 주문번호 JSON, Sheet 동기화, 되돌리기
  workers.py
quick_excel/
  parsers.py                    # 스토어 감지 및 클립보드 파싱
  service.py                    # 송장용 Excel 생성
  controller.py
orders/
  importers.py                  # 네이버, 쿠팡, G마켓, 11번가 Excel 처리
  mappings.py                   # 상품코드 매핑
  controller.py
courier_invoice/
  service.py                    # 송장 원본 처리 및 출력
  controller.py
product_detail/
  editor_dialog.py              # DetailHtmlEditorDialog
  controller.py                 # 상세페이지 탭과 프로세스 제어
product_category/
  service.py
  image_dialog.py
```

`MainWindow`는 버튼 연결, 입력 수집, 진행 상태와 결과 표시만 담당한다. 파일 처리, 계산, 외부 API 및 Sheet 접근은 기능 모듈에 둔다.

## 호환성 보존 항목

분리 과정에서 다음 계약은 변경하지 않는다.

- `ui/main_window.ui`의 object name, 탭 순서, signal 연결 대상
- Google Sheets 시트명, 헤더, 열 의미
- `database/` 아래 주문번호 및 중복 로드 기록 파일 형식
- 기존 Excel 출력의 열, 서식, 파일명 규칙
- 외부 스크립트의 인자 계약
  - `naver_to_coupang_html.py`
  - `coupang_detail_replace.py`
- OAuth 토큰 위치와 재인증 흐름
- 배송추적 수동중지/재개, 메모 비동기 저장, Slack 알림의 동작
- 기존 실행 방식: `python easy-fulfill.py`

## 0단계: 기준선 고정

### 작업

1. 현재 실행 절차, OAuth 상태, 출력 경로, 필요한 설정을 기록한다.
2. 다음 기준 테스트를 실행하고 결과를 남긴다.

   ```powershell
   python test_tracking_manual_stop.py
   python test_quick_excel_parser.py
   python -m unittest test_quotation_statement.py
   python -m py_compile easy-fulfill.py
   git diff --check
   ```

3. 기능별 수동 점검표를 작성한다.
   - 앱 기동과 모든 탭 진입
   - 주문 Excel 4종 읽기
   - Quick Excel 자동 및 수동 생성
   - 송장 출력
   - 배송추적 조회, 중지, 재개, 메모
   - OAuth 재인증 화면
4. UI object name, Sheet 열, 출력 파일 형식을 호환성 기준으로 기록한다.

### 완료 기준

기준 테스트 결과와 수동 점검표가 남아 있으며, 이후 모든 분리 단계에서 동일하게 재검증할 수 있다.

## 1단계: 공통 기반과 테스트 방식 정리

### 작업

1. 경로, 로깅, 예외 처리, Google Sheets client 획득을 공통 모듈로 정리한다.
2. 새 모듈은 전역 `MainWindow`나 widget에 직접 의존하지 않고 입력값 또는 명시적 의존성을 받게 한다.
3. `run_path("easy-fulfill.py")`로 전체 앱을 읽는 테스트를 기능 모듈 직접 import 방식으로 점진 전환한다.
4. 순수 함수를 UI에서 먼저 분리한다.
   - 배송 상태와 위험도 판정
   - 날짜 및 영업시간 계산
   - 우편번호 추출
   - 클립보드 텍스트 파싱
   - 주문번호 값 정규화

### 완료 기준

공통 모듈과 순수 함수 모듈은 Qt UI 없이 import할 수 있고, 단위 테스트가 앱 전체 초기화에 의존하지 않는다.

## 2단계: 배송추적 도메인 분리

### 범위

- 배송 상태 계산, 위험도 판정, 영업시간 계산
- 배송추적 Sheet 읽기와 쓰기
- 개별 및 일괄 송장 조회
- 추적 등록, 메모 저장, 수동중지/재개
- Slack 테스트, 일일 요약, 문의 알림
- 관련 `QThread` 클래스
- 배송추적 탭의 표 채우기, 필터, 검색, 컨텍스트 메뉴

### 권장 경계

- `tracking/service.py`: 상태 판정, 위험도, 알림 본문 등 순수 로직
- `tracking/repository.py`: Google Sheets와 KPOST API 접근
- `tracking/workers.py`: repository/service를 호출하는 비동기 작업
- `tracking/controller.py`: UI widget 접근과 화면 갱신

### 주의사항

- 메모는 현재 Sheet K열, 관리상태는 M열이라는 계약을 상수와 테스트로 고정한다.
- 메모의 pending/inflight 직렬화는 그대로 유지한다.
- 수동중지 상태는 자동 조회와 재분류의 제외 상태로 유지한다.
- 실제 Sheet 쓰기 검증은 헤드리스 테스트와 별도로 인증된 운영 계정에서 수행한다.

### 완료 기준

수동중지/재개, 메모, 조회, 알림이 기존 결과와 동일하며, 배송추적 관련 유지보수를 `easy-fulfill.py` 밖에서 할 수 있다.

## 3단계: 주문번호 인덱스 분리

### 범위

- `order_index.json` 읽기와 쓰기
- 당일 중복 로드 기록
- 공용 Sheet 읽기, 쓰기, 폴링
- 변경 이력과 되돌리기
- debounce 타이머와 동기화 상태

### 완료 기준

주문 처리 UI는 현재 번호 조회와 번호 변경 API만 호출하고, JSON 및 Sheet 구현 세부사항을 직접 알지 않는다.

## 4단계: Quick Excel 분리

### 범위

- 스토어 자동 감지
- 네이버, 쿠팡, G마켓, 일반 양식 파서
- 우편번호 추출
- 수동 입력 주소 검색
- 송장용 Excel 저장
- Quick Excel 탭 UI 제어

### 주의사항

- 현재 스토어별 파서 우선순위와, 감지 실패 시 일반 파서로 진입하는 규칙을 유지한다.
- 제품명, 모델명 기본값 및 출력 열을 변경하지 않는다.

### 완료 기준

`test_quick_excel_parser.py`가 새 모듈을 직접 검증하며, 자동 및 수동 생성 결과가 기존 양식과 일치한다.

## 5단계: 주문 Excel import 분리

### 범위

- 네이버, 쿠팡, G마켓, 11번가 Excel 읽기와 컬럼 검증
- 데이터프레임 변환
- 상품코드 매핑
- 주문 행 생성
- 중복 주문 파일 검사
- 주문 목록 반영 전 데이터 검증

### 구성 원칙

- `orders/importers.py`는 입력 파일, 매핑, 현재 인덱스를 받아 결과 객체를 반환한다.
- `orders/controller.py`가 파일 선택, 진행창, 메시지 박스, UI 표 반영을 담당한다.
- 마켓별 차이는 importer 내부 전략 함수로 두고, 공통 처리만 공유한다.

### 완료 기준

동일 샘플 처리 시 주문 수, 주문번호, 상품명, 링크, 배송추적 등록 데이터가 기존 결과와 일치한다.

## 6단계: 송장 처리와 상품분류 분리

### 범위

- 송장 파일 로드 및 검증
- 네이버/쿠팡 송장 데이터 변환
- 저장 파일 생성 및 실패 진단
- 배송추적 등록용 데이터 반환
- 상품 카테고리 계산 및 Excel export
- `ImageDialog`

### 주의사항

- Excel COM 종료 실패와 파일 생성 성공을 별도 상태로 기록한다.
- 출력 파일 생성 여부, 열, 행 수를 확인하고 성공을 판정한다.

### 완료 기준

`courier_invoice/service.py`는 Qt widget 없이 동작하고, UI는 결과와 오류 표시만 담당한다.

## 7단계: 상세페이지, DB 동기화, 인증 UI 분리

### 범위

- `DetailHtmlEditorDialog`
- 상세페이지 생성, 업로드, 적용 `QProcess` 제어
- DB-Sheet 동기화 탭
- Google OAuth 상태 표시, 연결 해제, 재인증
- 주소 검색 다이얼로그
- 상품 이미지 다이얼로그

### 완료 기준

각 탭과 다이얼로그는 자체 controller 또는 dialog 모듈에서 유지되며, `MainWindow`는 탭 장착과 공통 메뉴 연결만 담당한다.

## 8단계: 앱 셸 정리와 문서화

### 작업

1. `easy-fulfill.py`를 실행 진입점과 의존성 조립으로 축소한다.
2. `app/main_window.py`에는 탭 조립, 상태바, 메뉴, 공통 오류 표시만 남긴다.
3. `docs/architecture.md`에 모듈 책임과 공개 API를 문서화한다.
4. 사용하지 않는 이전 코드는 즉시 삭제하지 않고 `legacy/` 이동 여부를 별도 검토한다.
5. 전체 회귀 점검과 실제 운영 계정 기반 Sheet 쓰기 점검을 수행한다.

### 완료 기준

- `easy-fulfill.py`는 실행과 의존성 조립에 집중한다.
- 기능별 테스트는 대상 모듈을 직접 import한다.
- `MainWindow`는 기능별 대형 처리 메서드를 직접 가지지 않는다.
- UI, Sheet 스키마, 파일 출력 형식, 실행 방식이 이전과 호환된다.

## 커밋과 검증 원칙

- 한 커밋에는 한 도메인만 포함한다.
- 각 단계마다 `git diff --check`, 관련 테스트, 전체 `py_compile`을 실행한다.
- `.ui` 수정 시 XML 파싱과 실제 `QUiLoader` 로드를 검증한다.
- Sheet 쓰기, 외부 API, Excel COM, PDF는 단위 테스트와 구분하여 실제 환경에서 확인한다.
- 커밋 전 `git status --short`, `git diff --cached`, 최근 `git log`를 확인한다.
- 커밋 제목과 본문은 저장소 규칙에 맞게 한글로 작성한다.

## 권장 실행 순서

1. 0단계 기준선 고정
2. 1단계 공통 기반과 테스트 정리
3. 2단계 배송추적 분리
4. 3단계 주문번호 인덱스 분리
5. 4단계 Quick Excel 분리
6. 5단계 주문 Excel import 분리
7. 6단계 송장 처리와 상품분류 분리
8. 7단계 상세페이지, DB 동기화, 인증 UI 분리
9. 8단계 앱 셸 정리와 최종 회귀 검증

분리의 핵심은 파일 수를 늘리는 것이 아니라 UI, 도메인 규칙, 외부 I/O의 책임을 분리하고 각 단계에서 운영 동작을 증명하는 것이다.
