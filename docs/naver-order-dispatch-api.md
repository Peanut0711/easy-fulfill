# 네이버 주문 조회·발송처리 API 연동

작성: 2026-06-29 / 관련 커밋: `f877693`(주문 불러오기), `13bb0d7`(발송처리)

기존 수작업(네이버에서 주문 엑셀 다운 → 노션 정리 → 우체국 접수 → 송장번호 받아
네이버 업로드 양식으로 변환)에서 **네이버 쪽 두 구간을 API로 대체**했다. 우체국
소포접수(송장 자동발급)는 아직 수작업 — 다음 단계.

```
[①주문조회 API]  → 노션 정리(수동) → [우체국 접수(수동)] → [⑤발송처리 API]
   ✅ 구현                                  ⬜ ④ 미구현            ✅ 구현(실호출 미검증)
```

전 구간을 잇는 키는 **상품주문번호(productOrderId)**. 주문번호(orderId)는 결제 단위,
발주확인·발송·클레임은 모두 상품주문 단위로 처리된다.

---

## ① 주문 직접 불러오기 (엑셀 다운로드 대체)

**설계 원칙: 모드 토글이 아니라 "입력 소스 추가".**
- 기존 「파일 불러오기」(엑셀)는 그대로 둠 → 키 미설정자/장애 시 수동 폴백 가능.
- 「API로 불러오기」 버튼 신설 → 키 설정자만 사용.
- 두 경로가 **동일한 컬럼명의 DataFrame** 으로 합류 → 다운스트림(매핑·노션 MD·송장양식)
  코드는 **무변경**. 합류 메서드: `_build_naver_orders_from_df(df, product_mapping)`
  (엑셀 처리 함수에서 추출, 양 경로 공유).

**API 흐름** (`integrations/naver_commerce.py`)
1. `get_access_token` (bcrypt 전자서명 — 기존 문의 기능과 공용)
2. `fetch_changed_product_order_ids` — `GET …/product-orders/last-changed-statuses`
   - ⚠️ **한 번에 최대 24시간**만 허용(초과 시 `104140 조회 날짜가 유효하지 않습니다`).
     → from~to 를 **24시간 창으로 분할** 순회. 각 창 내 300건 초과는 `more`
     (moreFrom/moreSequence) 페이징.
   - 날짜 형식: `yyyy-MM-dd'T'HH:mm:ss.SSS+09:00` (requests 가 `+`→`%2B` 자동 인코딩).
3. `fetch_product_order_details` — `POST …/product-orders/query` (300건씩 청크,
   `quantityClaimCompatibility: true`)
4. 상태 필터: **`PAYED`** = 신규주문(결제완료) + 발주확인(상품준비중). 발송하면
   `DELIVERING` 으로 바뀌어 자동 제외.

**필드 매핑** (`order_detail_to_row`) — 2026-06-29 실응답으로 전부 확인됨:

| 엑셀 컬럼 | API 필드 |
|---|---|
| 주문번호 | `order.orderId` (13자리) |
| 수취인명/연락처1/우편번호 | `productOrder.shippingAddress.name / tel1 / zipCode` |
| 통합배송지 | `shippingAddress.baseAddress` + `detailedAddress` |
| 구매자연락처 | `order.ordererTel` |
| 상품명/옵션/수량 | `productOrder.productName / productOption / quantity` |
| 상품번호 | `productOrder.productId` (스마트스토어 상품번호, 상품코드 매핑 키) |
| 배송방법(구매자 요청) | `productOrder.expectedDeliveryMethod` (`DELIVERY`→`택배,등기,소포`) |
| 최종 상품별 총 주문금액 | `productOrder.totalPaymentAmount` |
| (보조)_상품주문번호/_상태 | `productOrder.productOrderId / productOrderStatus` |

- 배송메세지(`productOrder.shippingMemo`): 2026-11-22 실응답(1일치 발송 전 7건 중 메모 있는 1건, 3일치 12주문 중
  3건)으로 채워짐을 확인했다. 네이버는 메모가 없으면 필드 자체를 생략한다.
- 중첩 필드명은 공식 문서가 "OAS 참조"로 생략 → `_first(...)` 후보키 폴백으로 방어.

**중복 로드 방지:** 처리한 상품주문번호를 `database/loaded_order_ids.json` 에 당일 기록,
재불러오기 시 제외 → 공유 주문번호 인덱스 중복 상승 방지. (단 **PC별 로컬** — 한계 아래 참고)

UI: 「API로 불러오기」 버튼(최근 3일 고정), 백그라운드 `NaverOrderFetchThread`,
키는 공유 「설정」 시트에서 읽음.

---

## ⑤ 발송처리 (송장 등록, 업로드 엑셀 대체)

**송장번호 소스: 기존 우체국 송장 엑셀 재사용.** 기존 일괄발송 매칭
(`_process_naver_invoice`, 주문번호↔등기번호)을 그대로 쓰고, 그 과정에서
**전체 orderId + 등기번호**를 `self._naver_dispatch_records` 에 저장.

**흐름** (`naver_commerce.dispatch_orders_by_tracking`)
1. 각 orderId → `GET …/orders/{orderId}/product-order-ids` 로 상품주문번호 해석
2. `POST …/product-orders/dispatch` — `dispatchProductOrders[]` 에
   `{productOrderId, deliveryMethod: DELIVERY, deliveryCompanyCode: EPOST(우체국),
   trackingNumber, dispatchDate}` 묶어 다건 전송
3. 응답 `successProductOrderIds` / `failProductOrderInfos`(사유) 분리 → 결과창+콘솔

UI: 「API 발송처리」 버튼 → 확인 다이얼로그(배송중 전환 경고) → `NaverDispatchThread`
→ 성공/실패 요약.

---

## 검증 상태

- ✅ 주문 불러오기: 실키로 실제 동작 확인(5주문/6상품, 필드·그룹핑·노션 MD·상품코드 매핑 정상)
- ✅ 발송처리: 컴파일 + 오케스트레이션 모의테스트(중복제거·다상품 동일송장·실패파싱) 통과
- ⚠️ **발송처리 실호출 미검증** — 주문 상태를 실제로 바꾸는(배송중) 쓰기 호출이라
  되돌리기 어려움. **진짜 발송할 1건으로 먼저 테스트할 것.**

## 알려진 한계 (다음에 보완)

1. **중복 방지가 PC별 로컬** → 직원 여럿이 각자 API 불러오기 시 겹친 주문을 따로
   불러와 공유 인덱스가 중복 상승 가능. 해결안: 불러온 상품주문번호도 구글시트에 공유 기록.
2. **한 주문 다상품 → 전부 같은 등기번호로 발송**(한 박스 가정). 분할배송 미대응.
3. 발송처리 **택배사 우체국(EPOST) 고정**, `dispatchDate` 오늘 자동, 송장 정정은 별도 API 필요.
4. 주문 조회 범위 **최근 3일 고정**. 당일 처리가 원칙이라 금·토·일을 넘기는 일이 드물어 3일이면 충분하다(사용자 확인,
   2026-11-22). 「최근 3일 안에 상태가 바뀐 주문」만 보므로 그보다 오래 변화 없는 주문은 잡히지 않는다.
5. ~~호출 한도~~ 2026-11-22 점검 중 3·7·14일을 연달아 조회하자 `HTTP 429 GW.RATE_LIMIT`. `0.0.43`부터 모든 GET·POST가
   429면 1·2·4초 기다렸다 다시 보내고, 24시간 창 조회 사이에 0.3초 쉰다(3일 = 쉼 2번, 실측 5.7초, 429 없음).

## 「발주」 탭 — 발송 전 주문 조회(`0.0.43`, 자동화 1단계)

목표 흐름은 「주문 조회 → 바로 노션 정리 → 우체국 송장」 자동화이고, 이번은 첫 단계(조회 화면)만 만들었다.

- 측면 툴바 「발주」(탭 이름 「네이버 발주」, Ctrl+8). `ui/main/naver_orders.py` `NaverOrdersMixin`,
  `NaverPendingOrdersThread` → `run_naver_pending_orders_worker`(최근 `NAVER_ORDER_VIEW_DAYS`=3일) →
  `fetch_changed_product_order_ids` → `fetch_product_order_details` → `naver_commerce.group_pending_orders`.
- 신규주문·발주확인은 둘 다 `productOrderStatus == PAYED`이고 `placeOrderStatus`(`NOT_YET`/`OK`)만 다르다.
  주문번호별로 묶고, 하나라도 미확인이면 그 주문은 「신규주문」. 발송기한(가장 이른 것) → 결제 시각 순.
- 목록: 구분(클레임이면 「· 클레임」, 빨강)·결제·수취인·상품(「외 N건」)·금액. 필터는 다시 부르지 않고 거른다.
- 상세: 구분·클레임·배송방법 배지, 수취인 제목, 주문번호·발송기한, 상품 카드(옵션·수량·금액), 배송메모 카드(주황),
  받는 분 표(수령인·연락처·주소·배송방법). 카드 도우미는 `ui/main/rich_cards.py`(문의 탭과 공유).
- 조회만 한다. 주문번호 인덱스·공유 시트·로컬 `loaded_order_ids.json`에 쓰지 않는다.
- 다음 단계용: 주문 dict마다 `rows`(엑셀 처리 경로와 같은 컬럼, `order_detail_to_row`)를 들고 있다.
  「이 주문으로 처리 시작」은 이 행을 `_build_naver_orders_from_df`에 넘기면 된다(기존 API 불러오기와 같은 합류점).
  그때는 위 「알려진 한계」 1(중복 방지 공유화)을 먼저 풀어야 한다.
- 실응답(2026-11-22, 읽기 전용, 건수·칸 채움만): 3일치 12주문·상품주문 16건, 모두 발주확인, 클레임 1, 배송메모 3,
  여러 상품 주문 2. 수취인·연락처·주소·우편번호·결제·발송기한 빈 칸 0.

## 참고

- 전체 사용가능 엔드포인트: `docs/naver_api_list.txt.txt`
- LLM용 스펙 인덱스: `docs/naver_commerce_api_for_llm.txt` (상세는 외부 .md URL)
- 의존성: `bcrypt`(토큰 전자서명) — requirements.txt 에 있으나 .venv 미설치였어서
  2026-06-29 설치함. 배포 PC들도 `pip install -r requirements.txt` 확인 필요.
