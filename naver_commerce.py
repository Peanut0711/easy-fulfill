"""네이버 커머스API 문의 조회 래퍼 (UI/Qt 비의존 순수 함수 모듈).

조회 대상:
  - 상품문의(상품 Q&A): GET /v1/contents/qnas
  - 고객문의(네이버페이): GET /v1/pay-user/inquiries

인증 흐름(POST /v1/oauth2/token):
  client_id 와 client_secret(=bcrypt salt 형태의 문자열)으로 전자서명을 만들어
  액세스 토큰을 발급받는다. 전자서명은
    password = f"{client_id}_{timestamp_ms}"
    sign = base64( bcrypt.hashpw(password, client_secret) )
  로 생성한다. timestamp 는 발급 시각 ±5분 내에서만 유효하다.

키(client_id/secret)는 비밀값이므로 코드/레포에 두지 말 것(공유 「설정」 탭에만 저장).
"""

import base64
import time
from datetime import datetime, timedelta

try:
    import requests
except ImportError:  # requests 는 이미 의존성이지만 방어적으로 처리
    requests = None

try:
    import bcrypt
except ImportError:
    bcrypt = None

API_BASE = "https://api.commerce.naver.com/external"
TOKEN_URL = API_BASE + "/v1/oauth2/token"
QNAS_URL = API_BASE + "/v1/contents/qnas"            # 상품문의
INQUIRIES_URL = API_BASE + "/v1/pay-user/inquiries"  # 고객문의
PRODUCT_SEARCH_URL = API_BASE + "/v1/products/search"  # 판매 상품 목록
DEFAULT_TIMEOUT = 15
MAX_PAGES = 20  # 과도 호출 방지용 안전장치
KST = "+09:00"


def _require():
    if requests is None:
        raise RuntimeError("requests 패키지가 필요합니다. (pip install requests)")
    if bcrypt is None:
        raise RuntimeError("bcrypt 패키지가 필요합니다. (pip install bcrypt)")


def _raise_for_status_with_body(resp):
    """HTTP 오류 시 네이버가 본문에 담아 보내는 에러코드/메시지까지 예외에 포함시킨다.
    (403 등은 본문의 code/message 가 원인 파악의 핵심이다.)"""
    if resp.status_code < 400:
        return
    body = ""
    try:
        js = resp.json()
        # 네이버 게이트웨이/도메인 에러: code, message, invalidInputs 등
        code = js.get("code") or js.get("error") or ""
        msg = js.get("message") or js.get("error_description") or ""
        body = " ".join(x for x in (str(code), str(msg)) if x).strip() or str(js)
    except Exception:
        body = (resp.text or "").strip()
    if len(body) > 300:
        body = body[:300] + "…"
    raise RuntimeError(f"HTTP {resp.status_code}: {body}")


def _make_signature(client_id, client_secret, timestamp_ms):
    """client_id_timestamp 문자열을 client_secret(bcrypt salt)로 해시 후 base64."""
    password = f"{client_id}_{timestamp_ms}"
    hashed = bcrypt.hashpw(password.encode("utf-8"), client_secret.encode("utf-8"))
    return base64.b64encode(hashed).decode("utf-8")


def get_access_token(client_id, client_secret, account_type="SELF"):
    """액세스 토큰 문자열을 반환. 네트워크/인증 실패 시 예외 발생."""
    _require()
    cid = str(client_id or "").strip()
    csec = str(client_secret or "").strip()
    ts = int(time.time() * 1000)
    sign = _make_signature(cid, csec, ts)
    data = {
        "client_id": cid,
        "timestamp": ts,
        "grant_type": "client_credentials",
        "client_secret_sign": sign,
        "type": account_type,
    }
    resp = requests.post(TOKEN_URL, data=data, timeout=DEFAULT_TIMEOUT)
    _raise_for_status_with_body(resp)
    js = resp.json()
    token = js.get("access_token")
    if not token:
        raise RuntimeError(f"토큰 응답에 access_token 이 없습니다: {js}")
    return token


def validate_credentials(client_id, client_secret):
    """토큰 발급을 시도해 키 유효성만 확인. 반환 {ok, valid, error}."""
    cid = str(client_id or "").strip()
    csec = str(client_secret or "").strip()
    if not cid or not csec:
        return {"ok": True, "valid": False,
                "error": "client_id 또는 client_secret 이 비어 있습니다."}
    try:
        get_access_token(cid, csec)
        return {"ok": True, "valid": True, "error": ""}
    except Exception as e:
        return {"ok": True, "valid": False, "error": str(e)}


def _get_json(url, token, params):
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, params=params, timeout=DEFAULT_TIMEOUT)
    _raise_for_status_with_body(resp)
    return resp.json()


def fetch_product_qnas(token, from_dt, to_dt, answered=None):
    """상품문의 목록(contents 항목 리스트)을 반환.
    from_dt/to_dt: datetime. answered=False 면 미답변만 조회."""
    _require()
    items = []
    page = 1
    froms = from_dt.strftime("%Y-%m-%dT%H:%M:%S.000") + KST
    tos = to_dt.strftime("%Y-%m-%dT%H:%M:%S.000") + KST
    while page <= MAX_PAGES:
        params = {"fromDate": froms, "toDate": tos, "page": page, "size": 100}
        if answered is not None:
            params["answered"] = "true" if answered else "false"
        js = _get_json(QNAS_URL, token, params)
        contents = js.get("contents") or []
        items.extend(contents)
        if not contents or js.get("last", True):
            break
        page += 1
    return items


def fetch_customer_inquiries(token, start_date, end_date, answered=None):
    """고객문의 목록(content 항목 리스트)을 반환.
    start_date/end_date: date 또는 datetime(yyyy-MM-dd 로 변환). answered=False 면 미답변만."""
    _require()
    items = []
    page = 1
    sd = start_date.strftime("%Y-%m-%d")
    ed = end_date.strftime("%Y-%m-%d")
    while page <= MAX_PAGES:
        params = {"startSearchDate": sd, "endSearchDate": ed, "page": page, "size": 200}
        if answered is not None:
            params["answered"] = "true" if answered else "false"
        js = _get_json(INQUIRIES_URL, token, params)
        content = js.get("content") or []
        items.extend(content)
        if not content or js.get("last", True):
            break
        page += 1
    return items


# ── 주문(발주) 조회 ─────────────────────────────────────────────
# 폴링 파이프라인: last-changed-statuses(변경 식별자만) → query(상세 풀세트).
PRODUCT_ORDERS_BASE = API_BASE + "/v1/pay-order/seller/product-orders"
LAST_CHANGED_URL = PRODUCT_ORDERS_BASE + "/last-changed-statuses"
ORDER_QUERY_URL = PRODUCT_ORDERS_BASE + "/query"
DISPATCH_URL = PRODUCT_ORDERS_BASE + "/dispatch"
QUERY_CHUNK = 300  # query API 1회 최대 상품주문 수(상한)

# 거래명세서 자동 발급 대상에서 제외할 거래 종료 상태. 부분 취소 등으로 한 주문에
# 정상 상품과 아래 상태 상품이 섞인 경우에도 자동 발급은 막아 원거래와 다른 문서가
# 만들어지지 않게 한다.
TRANSACTION_STATEMENT_BLOCKED_STATUSES = {
    "CANCELED",
    "CANCELED_BY_NOPAYMENT",
    "RETURNED",
    "EXCHANGED",
}
TRANSACTION_STATEMENT_BLOCKED_CLAIM_STATUSES = {
    "CANCEL_DONE",
    "RETURN_DONE",
    "EXCHANGE_DONE",
    "ADMIN_CANCEL_DONE",
}

# 우체국 택배사 코드(발송처리 deliveryCompanyCode). 우체국택배=EPOST,
# 우편등기=REGISTPOST, 일반우편=GENERALPOST.
KPOST_COMPANY_CODE = "EPOST"

# 발송 전(아직 미발송) 상태. '신규주문(결제완료)'과 '발주확인(상품준비중)'은
# 모두 productOrderStatus == 'PAYED' 로 내려오고, 발송하면 DELIVERING 으로 바뀐다.
SHIPPABLE_STATUSES = ("PAYED",)

# 배송방법 enum → 엑셀 '배송방법(구매자 요청)' 라벨. 다운스트림은 '택배,등기,소포'
# 문자열만 특별 처리(배송비 중복 결제 방지)하므로 DELIVERY 매핑이 핵심이다.
DELIVERY_METHOD_LABELS = {
    "DELIVERY": "택배,등기,소포",
    "GDFW_ISSUE_SVC": "퀵서비스",
    "VISIT_RECEIPT": "방문수령",
    "DIRECT_DELIVERY": "직접배송(화물배달)",
    "NOTHING": "배송없음",
}


def _post_json(url, token, payload):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=payload, timeout=DEFAULT_TIMEOUT)
    _raise_for_status_with_body(resp)
    return resp.json()


def fetch_sale_products(token, max_pages=MAX_PAGES):
    """판매 중 스마트스토어 상품을 문서 검색용의 단순한 dict 목록으로 반환한다.

    상품 목록 API는 상품명 자유검색을 제공하지 않으므로 최대 500개씩 조회하고,
    이름 키워드 필터는 호출 측에서 수행한다. 쇼핑윈도 채널은 제외한다.
    """
    _require()
    products = []
    page = 1
    size = 500
    while page <= max_pages:
        js = _post_json(PRODUCT_SEARCH_URL, token, {
            "productStatusTypes": ["SALE"],
            "page": page,
            "size": size,
            "orderType": "NAME",
        })
        contents = js.get("contents") or []
        for group in contents:
            for product in (group.get("channelProducts") or []):
                if product.get("channelServiceType") != "STOREFARM":
                    continue
                sale_price = int(product.get("salePrice") or 0)
                discounted_price = int(product.get("discountedPrice") or 0)
                products.append({
                    "product_no": str(product.get("channelProductNo") or ""),
                    "name": str(product.get("name") or "").strip(),
                    "sale_price": sale_price,
                    "discounted_price": discounted_price,
                    "price": discounted_price if discounted_price > 0 else sale_price,
                })
        total = int(js.get("totalElements") or 0)
        if not contents or len(contents) < size or (total and page * size >= total):
            break
        page += 1
    return products


def _first(d, *keys):
    """dict d 에서 keys 를 순서대로 보고 비어있지 않은 첫 값을 반환(없으면 '')."""
    if not isinstance(d, dict):
        return ""
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return ""


def _fmt_dt(dt):
    """네이버 규격(yyyy-MM-dd'T'HH:mm:ss.SSS+09:00)으로 포맷."""
    return dt.strftime("%Y-%m-%dT%H:%M:%S.000") + KST


def fetch_changed_product_order_ids(token, from_dt, to_dt=None,
                                    changed_type=None, max_pages=MAX_PAGES):
    """from_dt~to_dt 사이 변경된 상품주문번호 목록(중복 제거, 시간순)을 반환.

    네이버 변경조회는 한 번에 '최대 24시간' 범위만 허용하므로(초과 시
    104140 '조회 날짜가 유효하지 않습니다'), 범위를 24시간 창으로 쪼개 순회한다.
    각 창 안에서 응답이 300건을 넘으면 more(moreFrom/moreSequence)로 이어 받는다.
    """
    _require()
    if to_dt is None:
        to_dt = from_dt + timedelta(hours=24)
    ids = []
    seen = set()
    window = timedelta(hours=24)
    pages = 0
    w_start = from_dt
    while w_start < to_dt and pages < max_pages:
        w_end = min(w_start + window, to_dt)
        cur_from = _fmt_dt(w_start)
        to_s = _fmt_dt(w_end)
        more_seq = None
        while pages < max_pages:
            pages += 1
            params = {"lastChangedFrom": cur_from, "lastChangedTo": to_s,
                      "limitCount": 300}
            if changed_type:
                params["lastChangedType"] = changed_type
            if more_seq:
                params["moreSequence"] = more_seq
            js = _get_json(LAST_CHANGED_URL, token, params)
            data = js.get("data") or {}
            for it in (data.get("lastChangeStatuses") or []):
                poid = it.get("productOrderId")
                if poid is None:
                    continue
                poid = str(poid)
                if poid not in seen:
                    seen.add(poid)
                    ids.append(poid)
            more = data.get("more")
            if not more:
                break
            nxt = more.get("moreFrom")
            more_seq = more.get("moreSequence")
            if not nxt:
                break
            cur_from = nxt
        w_start = w_end
    return ids


def fetch_product_order_details(token, product_order_ids):
    """상품주문번호들을 300개씩 묶어 상세(data 항목 리스트)를 반환한다."""
    _require()
    out = []
    ids = [str(x) for x in product_order_ids if x not in (None, "")]
    for i in range(0, len(ids), QUERY_CHUNK):
        chunk = ids[i:i + QUERY_CHUNK]
        js = _post_json(ORDER_QUERY_URL, token,
                        {"productOrderIds": chunk, "quantityClaimCompatibility": True})
        out.extend(js.get("data") or [])
    return out


def order_detail_to_row(item):
    """query 응답의 data 항목 1개 → 주문 엑셀과 동일한 컬럼명의 dict 1행.

    네이버 문서가 order/productOrder/shippingAddress 의 중첩 필드명을 'OAS 참조'로
    생략하므로, 알려진 필드명 + 후보키 폴백으로 방어적으로 추출한다. 실제 응답과
    어긋나는 항목이 있으면 아래 _first(...) 후보 목록만 손보면 된다.
    """
    order = item.get("order") or {}
    po = item.get("productOrder") or {}
    addr = po.get("shippingAddress") or {}

    base = _first(addr, "baseAddress", "roadNameAddress", "address", "addressName")
    detail = _first(addr, "detailedAddress", "detailAddress")
    full_addr = (str(base) + " " + str(detail)).strip()

    method_enum = _first(po, "expectedDeliveryMethod", "deliveryMethod", "deliveryPolicyType")
    method = DELIVERY_METHOD_LABELS.get(str(method_enum).upper(), str(method_enum) or "")

    return {
        "주문번호": str(_first(order, "orderId") or _first(po, "orderId")),
        "수취인명": str(_first(addr, "name", "receiverName", "ordererName")),
        "수취인연락처1": str(_first(addr, "tel1", "tel2", "receiverTel")),
        "통합배송지": full_addr,
        "구매자연락처": str(_first(order, "ordererTel", "ordererPhoneNumber",
                                     "ordererCellPhoneNumber")),
        "배송메세지": str(_first(po, "shippingMemo") or _first(addr, "shippingMemo")),
        "상품명": str(_first(po, "productName")),
        "옵션정보": str(_first(po, "productOption", "optionInfo")),
        "수량": _first(po, "quantity") or 1,
        "우편번호": str(_first(addr, "zipCode", "zipcode")),
        "상품번호": str(_first(po, "productId", "channelProductNo",
                                 "merchantChannelId", "originProductNo")),
        "배송방법(구매자 요청)": method,
        "최종 상품별 총 주문금액": _first(po, "totalPaymentAmount",
                                          "totalProductAmount") or 0,
        # 보조 컬럼(다운스트림은 무시; 발송처리/중복방지용으로 보관)
        "_상품주문번호": str(_first(po, "productOrderId")),
        "_상태": str(_first(po, "productOrderStatus")),
    }


def dispatch_product_orders(token, items, dispatch_dt=None,
                            default_company=KPOST_COMPANY_CODE):
    """발송 처리(송장 등록). 상태를 발주확인→배송중으로 전이시킨다.

    items: [{productOrderId, trackingNumber, deliveryCompanyCode(옵션),
             deliveryMethod(옵션, 기본 DELIVERY), dispatchDate(옵션)}].
    반환: {ok, success:[productOrderId...], fail:[{productOrderId, reason}...], raw}.
    같은 productOrderId+동일 송장 재호출은 보통 무시되지만, 다른 송장으로의
    정정은 이 API가 아닌 별도 송장수정 API를 써야 한다.
    """
    _require()
    default_date = _fmt_dt(dispatch_dt) if dispatch_dt else _fmt_dt(datetime.now())
    body = {"dispatchProductOrders": []}
    for it in items:
        poid = str(it.get("productOrderId") or "").strip()
        tracking = str(it.get("trackingNumber") or "").strip()
        if not poid or not tracking:
            continue
        body["dispatchProductOrders"].append({
            "productOrderId": poid,
            "deliveryMethod": it.get("deliveryMethod") or "DELIVERY",
            "deliveryCompanyCode": it.get("deliveryCompanyCode") or default_company,
            "trackingNumber": tracking,
            "dispatchDate": it.get("dispatchDate") or default_date,
        })
    if not body["dispatchProductOrders"]:
        return {"ok": True, "success": [], "fail": [], "raw": {}}
    js = _post_json(DISPATCH_URL, token, body)
    data = js.get("data") or {}
    success = [str(x) for x in (data.get("successProductOrderIds") or [])]
    fail = []
    for f in (data.get("failProductOrderInfos") or []):
        if isinstance(f, dict):
            fail.append({
                "productOrderId": str(f.get("productOrderId") or ""),
                "reason": str(f.get("message") or f.get("reason")
                              or f.get("code") or f),
            })
        else:
            fail.append({"productOrderId": "", "reason": str(f)})
    return {"ok": True, "success": success, "fail": fail, "raw": data}


def fetch_product_order_ids_of_order(token, order_id):
    """주문번호(orderId) 1개에 속한 상품주문번호(productOrderId) 목록을 반환."""
    _require()
    url = API_BASE + f"/v1/pay-order/seller/orders/{order_id}/product-order-ids"
    js = _get_json(url, token, None)
    return [str(x) for x in (js.get("data") or [])]


def _as_positive_int(value, default=0):
    """네이버 금액/수량 응답을 양의 정수로 안전하게 변환한다."""
    try:
        converted = int(value)
    except (TypeError, ValueError):
        return default
    return converted if converted > 0 else default


def build_transaction_statement_order(order_id, details):
    """상품주문 상세 응답을 거래명세서 화면용 주문 1건으로 정규화한다.

    네이버 주문 자체에 적용된 할인 결과(totalPaymentAmount)를 품목별 금액으로
    보존한다. 수량으로 나누어 떨어지지 않는 금액도 문서 계산 모듈의
    gross_amount_override로 정확한 행 합계를 유지할 수 있도록 함께 반환한다.
    """
    requested_id = str(order_id or "").strip()
    if not requested_id:
        raise ValueError("네이버 주문번호를 입력해주세요.")
    if not details:
        raise ValueError("입력한 주문번호의 상품 주문 정보를 찾지 못했습니다.")

    items = []
    orderer_name = ""
    receiver_name = ""
    payment_date = ""
    blocked = []
    for detail in details:
        order = detail.get("order") or {}
        product_order = detail.get("productOrder") or {}
        response_order_id = str(
            _first(order, "orderId") or _first(product_order, "orderId")
        ).strip()
        if response_order_id and response_order_id != requested_id:
            continue

        status = str(_first(product_order, "productOrderStatus")).upper()
        claim_status = str(_first(product_order, "claimStatus")).upper()
        if (status in TRANSACTION_STATEMENT_BLOCKED_STATUSES
                or claim_status in TRANSACTION_STATEMENT_BLOCKED_CLAIM_STATUSES):
            blocked.append(claim_status or status)
            continue

        quantity = _as_positive_int(_first(product_order, "quantity"), default=1)
        gross_amount = _as_positive_int(
            _first(product_order, "totalPaymentAmount", "totalProductAmount")
        )
        if not gross_amount:
            unit_price = _as_positive_int(_first(product_order, "unitPrice"))
            option_price = _as_positive_int(_first(product_order, "optionPrice"))
            gross_amount = (unit_price + option_price) * quantity
        if not gross_amount:
            raise ValueError("주문 상품의 결제금액을 확인할 수 없습니다.")

        name = str(_first(product_order, "productName")).strip()
        if not name:
            raise ValueError("주문 상품명을 확인할 수 없습니다.")
        orderer_name = orderer_name or str(_first(order, "ordererName")).strip()
        address = product_order.get("shippingAddress") or {}
        receiver_name = receiver_name or str(_first(address, "name", "receiverName")).strip()
        payment_date = payment_date or str(_first(order, "paymentDate", "orderDate")).strip()
        # 단가는 화면 표시용 반올림값이고, 문서 합계는 gross_amount로 계산한다.
        display_unit_price = int((gross_amount + quantity // 2) // quantity)
        items.append({
            "product_order_id": str(_first(product_order, "productOrderId")).strip(),
            "name": name,
            "specification": str(_first(product_order, "productOption", "optionInfo")).strip(),
            "quantity": quantity,
            "gross_unit_price": display_unit_price,
            "gross_amount": gross_amount,
            "status": status,
        })

    if blocked:
        labels = ", ".join(sorted(set(blocked)))
        raise ValueError(
            "취소·반품·교환 완료 상품이 포함된 주문은 자동으로 거래명세서를 "
            f"발급할 수 없습니다. (상태: {labels})"
        )
    if not items:
        raise ValueError("발급할 수 있는 주문 상품을 찾지 못했습니다.")
    return {
        "order_id": requested_id,
        "orderer_name": orderer_name,
        "receiver_name": receiver_name,
        "customer_name": orderer_name or receiver_name,
        "payment_date": payment_date,
        "items": items,
    }


def fetch_order_for_transaction_statement(token, order_id):
    """주문번호 입력 → 상품주문번호 → 상세 조회를 한 번에 수행한다."""
    normalized_order_id = str(order_id or "").strip()
    if not normalized_order_id:
        raise ValueError("네이버 주문번호를 입력해주세요.")
    product_order_ids = fetch_product_order_ids_of_order(token, normalized_order_id)
    if not product_order_ids:
        raise ValueError("입력한 주문번호의 상품주문번호를 찾지 못했습니다.")
    details = fetch_product_order_details(token, product_order_ids)
    return build_transaction_statement_order(normalized_order_id, details)


def dispatch_orders_by_tracking(token, records, company=KPOST_COMPANY_CODE,
                                dispatch_dt=None):
    """주문번호↔송장번호 쌍을 받아 상품주문번호로 풀어 발송처리한다.

    records: [{orderId, trackingNumber, ...}]. 각 orderId 의 productOrderId 를
    조회(fetch_product_order_ids_of_order)해 같은 송장번호로 dispatch 한다.
    한 주문에 상품이 여럿이면 모두 같은 등기번호로 발송 처리된다(한 박스 가정).
    반환: {ok, success, fail, resolved, errors}.
    """
    _require()
    items = []
    resolved = []
    errors = []
    seen = set()
    for r in records:
        oid = str(r.get("orderId") or "").strip()
        tracking = str(r.get("trackingNumber") or "").strip()
        if not oid or not tracking or oid in seen:
            continue
        seen.add(oid)
        try:
            poids = fetch_product_order_ids_of_order(token, oid)
        except Exception as e:
            errors.append({"orderId": oid, "error": str(e)})
            continue
        if not poids:
            errors.append({"orderId": oid, "error": "상품주문번호를 찾지 못함"})
            continue
        resolved.append({"orderId": oid, "trackingNumber": tracking,
                         "productOrderIds": poids})
        for poid in poids:
            items.append({"productOrderId": poid, "trackingNumber": tracking,
                          "deliveryCompanyCode": company})
    res = dispatch_product_orders(token, items, dispatch_dt=dispatch_dt,
                                  default_company=company)
    res["resolved"] = resolved
    res["errors"] = errors
    return res


def fetch_orders_for_shipping(client_id, client_secret, from_dt, to_dt=None,
                              statuses=SHIPPABLE_STATUSES, account_type="SELF",
                              debug=False):
    """발송 전 주문을 '주문 엑셀과 동일 컬럼'의 행(dict) 리스트로 반환한다.

    흐름: 토큰 발급 → last-changed-statuses(변경 식별자) → query(상세) → 상태 필터.
    statuses 가 비어 있으면 상태 필터를 적용하지 않는다.
    """
    token = get_access_token(client_id, client_secret, account_type=account_type)
    ids = fetch_changed_product_order_ids(token, from_dt, to_dt)
    details = fetch_product_order_details(token, ids)
    if debug and details:
        po0 = details[0].get("productOrder") or {}
        print("[naver order sample] order keys:",
              list((details[0].get("order") or {}).keys()))
        print("[naver order sample] productOrder keys:", list(po0.keys()))
        print("[naver order sample] shippingAddress keys:",
              list((po0.get("shippingAddress") or {}).keys()))
    rows = [order_detail_to_row(it) for it in details]
    if statuses:
        sset = set(statuses)
        rows = [r for r in rows if r.get("_상태") in sset]
    return rows
