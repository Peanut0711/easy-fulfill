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
QUERY_CHUNK = 300  # query API 1회 최대 상품주문 수(상한)

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
