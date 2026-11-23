"""쿠팡 WING OpenAPI 문의 조회 래퍼 (UI/Qt 비의존 순수 함수 모듈).

조회 대상(미답변 위주):
  - 온라인 고객문의(상품문의): GET …/api/v5/vendors/{vendorId}/onlineInquiries
  - 콜센터(CS) 문의:        GET …/api/v5/vendors/{vendorId}/callCenterInquiries

인증(네이버와 다름): 토큰이 아니라 '요청마다 HMAC 서명'을 만든다(CEA HmacSHA256).
  signed_date = 현재 GMT, 포맷 yyMMdd'T'HHmmss'Z'
  message     = signed_date + method + path + query   (query 는 '?' 없이)
  signature   = hex( HMAC-SHA256(secret_key, message) )
  헤더 Authorization:
    "CEA algorithm=HmacSHA256, access-key={access}, signed-date={signed_date}, signature={signature}"
표준 라이브러리(hmac/hashlib)만 사용 → 추가 의존성 없음.

키(access_key/secret_key/vendor_id)는 비밀값이므로 코드/레포에 두지 말 것(공유 「설정」 탭에만 저장).
주의: 서명에 쓰는 query 문자열과 실제 요청 URL 의 query 가 '완전히 동일'해야 한다
(인코딩·순서까지). 그래서 params 를 직접 urlencode 해 URL 에 붙이고 같은 문자열로 서명한다.
"""

import hashlib
import hmac
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from urllib.parse import urlencode

try:
    import requests
except ImportError:  # requests 는 이미 의존성이지만 방어적으로 처리
    requests = None

API_GATEWAY = "https://api-gateway.coupang.com"
PATH_PREFIX = "/v2/providers/openapi/apis/api/v5/vendors"
DEFAULT_TIMEOUT = 15
MAX_PAGES = 20          # 과도 호출 방지
MAX_RANGE_DAYS = 7      # 쿠팡 문의 조회는 한 번에 최대 약 7일 → 그 단위로 끊어 호출
PAGE_SIZE = 50


def _require():
    if requests is None:
        raise RuntimeError("requests 패키지가 필요합니다. (pip install requests)")


def _raise_for_status_with_body(resp):
    """HTTP 오류 시 쿠팡이 본문에 담아 보내는 code/message 까지 예외에 포함시킨다."""
    if resp.status_code < 400:
        return
    body = ""
    try:
        js = resp.json()
        code = js.get("code") or js.get("error") or ""
        msg = js.get("message") or js.get("errorMessage") or ""
        body = " ".join(x for x in (str(code), str(msg)) if x).strip() or str(js)
    except Exception:
        body = (resp.text or "").strip()
    if len(body) > 300:
        body = body[:300] + "…"
    raise RuntimeError(f"HTTP {resp.status_code}: {body}")


def _signed_date():
    """GMT 기준 yyMMdd'T'HHmmss'Z'."""
    return time.strftime("%y%m%dT%H%M%SZ", time.gmtime())


def _authorization(method, path, query, access_key, secret_key):
    signed = _signed_date()
    message = signed + method + path + query
    signature = hmac.new(
        secret_key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256
    ).hexdigest()
    return (f"CEA algorithm=HmacSHA256, access-key={access_key}, "
            f"signed-date={signed}, signature={signature}")


def _get(path, params, access_key, secret_key):
    """서명을 만들어 GET 요청 후 파싱된 JSON 을 반환."""
    _require()
    query = urlencode(params)  # dict 입력 순서를 유지(서명·요청 동일 문자열 보장)
    auth = _authorization("GET", path, query, access_key, secret_key)
    url = API_GATEWAY + path + ("?" + query if query else "")
    headers = {
        "Authorization": auth,
        "Content-Type": "application/json;charset=UTF-8",
    }
    resp = requests.get(url, headers=headers, timeout=DEFAULT_TIMEOUT)
    _raise_for_status_with_body(resp)
    return resp.json()


def _extract_items(js):
    """응답에서 항목 리스트를 꺼낸다(data.content 또는 data 가 리스트인 두 형태 모두 대응)."""
    data = js.get("data")
    if isinstance(data, dict):
        return data.get("content") or data.get("inquiries") or []
    if isinstance(data, list):
        return data
    return []


def _date_windows(from_dt, to_dt, days=MAX_RANGE_DAYS):
    """[from_dt, to_dt] 를 '겹치지 않고' 날짜 기준 최대 days 일(양끝 포함) 구간들로 쪼갠다.
    쿠팡은 inquiryStartAt~inquiryEndAt 를 양끝 포함으로 세므로(예 06-01~06-07 = 7일),
    한 구간의 끝은 시작+ (days-1) 일로 잡아 7일을 넘기지 않게 한다. 반환은 date 객체."""
    windows = []
    start = from_dt.date()
    end_date = to_dt.date()
    span = timedelta(days=max(1, days) - 1)  # 양끝 포함 days 일 = 차이 (days-1)
    one = timedelta(days=1)
    cur = start
    while cur <= end_date:
        w_end = min(cur + span, end_date)
        windows.append((cur, w_end))
        cur = w_end + one
    if not windows:
        windows.append((start, end_date))
    return windows


def _page_count(js):
    """응답의 data.pagination.totalPages(없으면 None)."""
    data = js.get("data")
    pagination = data.get("pagination") if isinstance(data, dict) else None
    try:
        return int(pagination.get("totalPages"))
    except (AttributeError, TypeError, ValueError):
        return None


def _paged(path, base_params, access_key, secret_key):
    """pageNum 을 늘려가며 모든 페이지 항목을 모은다.

    콜센터문의는 pageSize 를 50으로 보내도 한 페이지에 30건만 준다(2026-11-23 실응답 countPerPage 30).
    그래서 받은 건수가 아니라 응답의 totalPages 로 끝을 판단하고, totalPages 가 없을 때만 건수로 본다.
    """
    items = []
    page = 1
    while page <= MAX_PAGES:
        params = dict(base_params)
        params["pageNum"] = page
        params["pageSize"] = PAGE_SIZE
        js = _get(path, params, access_key, secret_key)
        page_items = _extract_items(js)
        items.extend(page_items)
        total_pages = _page_count(js)
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(page_items) < PAGE_SIZE:
            break
        page += 1
    return items


def fetch_online_inquiries(vendor_id, access_key, secret_key, from_dt, to_dt,
                           answered=None):
    """온라인 고객문의(상품문의) 목록을 반환. answered=False → 미답변만(NOANSWER)."""
    path = f"{PATH_PREFIX}/{vendor_id}/onlineInquiries"
    answered_type = "ALL" if answered is None else ("ANSWERED" if answered else "NOANSWER")
    out = []
    for s, e in _date_windows(from_dt, to_dt):
        base = {
            "vendorId": vendor_id,
            "answeredType": answered_type,
            "inquiryStartAt": s.strftime("%Y-%m-%d"),
            "inquiryEndAt": e.strftime("%Y-%m-%d"),
        }
        out.extend(_paged(path, base, access_key, secret_key))
    return out


def fetch_callcenter_inquiries(vendor_id, access_key, secret_key, from_dt, to_dt,
                               answered=None):
    """콜센터(CS) 문의 목록을 반환. answered=False → 미답변만(NO_ANSWER), None → 전체(NONE).

    partnerCounselingStatus 를 빼면 0건이 온다(2026-11-23 실응답). 그래서 전체도 NONE 으로 보낸다.
    """
    path = f"{PATH_PREFIX}/{vendor_id}/callCenterInquiries"
    status = "NONE" if answered is None else ("ANSWER" if answered else "NO_ANSWER")
    out = []
    for s, e in _date_windows(from_dt, to_dt):
        base = {
            "vendorId": vendor_id,
            "inquiryStartAt": s.strftime("%Y-%m-%d"),
            "inquiryEndAt": e.strftime("%Y-%m-%d"),
            "partnerCounselingStatus": status,
        }
        out.extend(_paged(path, base, access_key, secret_key))
    return out


# ── 「문의」 탭: 쿠팡 문의 조회(읽기 전용) ─────────────────────────────
# 실응답(2026-11-23, 최근 30일 읽기 전용, 키 이름·상태 값·건수만 확인):
#   온라인문의 45건: inquiryId, productId, sellerProductId, sellerItemId, vendorItemId, content, inquiryAt,
#     orderIds[], commentDtoList[{inquiryCommentId, inquiryId, content, inquiryCommentAt}]. 상품명·작성자 칸은 없다.
#     미답변(NOANSWER)은 0건이었고 45건 모두 답변(commentDtoList)이 1~2개 있었다.
#   콜센터문의 103건: inquiryId, csPartnerCounselingStatus, inquiryStatus, vendorItemId[], itemName, content,
#     answeredAt, replies[{receptionistName, replyAt, content, answerType, partnerTransferStatus, needAnswer, …}],
#     inquiryAt, buyerEmail, buyerPhone(가린 번호), orderId, orderDate, receiptCategory.
#     상태 조합: answered+complete/progress 95건(ANSWER), requestAnswer+progress 7건(NO_ANSWER),
#     requestAnswer+complete 1건(TRANSFER). replies 의 answerType 은 csAgent(쿠팡 상담원)·vendor(판매자).
INQUIRY_KIND_ONLINE = "쿠팡 상품문의"
INQUIRY_KIND_CALLCENTER = "쿠팡 CS문의"
# 온라인문의에는 상품명이 없어 등록 상품 상세(약 0.7초/건)로 찾는다. 한 번 찾은 이름은 앱을 끌 때까지 기억하고,
# 한 번 조회에 새로 찾는 상품 수를 제한한다(30일 전체 조회 때 수십 개를 찾느라 오래 걸리지 않게).
PRODUCT_NAME_LOOKUP_LIMIT = 20
PRODUCT_NAME_LOOKUP_WORKERS = 4
_product_names = {}
_product_names_lock = threading.Lock()


def short_datetime(value):
    """'2026-11-20T13:22:11+09:00' → '2026-11-20 13:22'(모르는 형식은 그대로)."""
    s = str(value or "").strip()
    if len(s) >= 16 and s[4] == "-" and s[7] == "-" and s[10] in "T " and s[13] == ":":
        return s[:10] + " " + s[11:16]
    return s


def _text(value):
    return "" if value is None else str(value).strip()


def _category(value):
    """receiptCategory('배송>배송>배송', '>>') → 겹치는 단계를 뺀 '배송'."""
    parts = []
    for part in _text(value).split(">"):
        part = part.strip()
        if part and part not in parts:
            parts.append(part)
    return " > ".join(parts)


def normalize_online_inquiry(item, product_names=None):
    """온라인문의(상품문의) 1개 → 문의 탭 행 dict. 답변은 commentDtoList(판매자 답변) 전부를 이어 보인다."""
    comments = [c for c in (item.get("commentDtoList") or []) if isinstance(c, dict)]
    seller_product_id = _text(item.get("sellerProductId"))
    name = (product_names or {}).get(seller_product_id, "")
    order_ids = [_text(x) for x in (item.get("orderIds") or []) if _text(x)]
    return {
        "kind": INQUIRY_KIND_ONLINE,
        "store": "쿠팡",
        "id": _text(item.get("inquiryId")),
        "created": short_datetime(item.get("inquiryAt")),
        "writer": "",
        "product": name,
        "product_no": seller_product_id,
        "order_id": ", ".join(order_ids),
        "product_order_ids": [],
        "category": "",
        "title": "",
        "content": _text(item.get("content")),
        "answered": bool(comments),
        "answer": "\n\n".join(_text(c.get("content")) for c in comments if _text(c.get("content"))),
        "answered_at": short_datetime(comments[-1].get("inquiryCommentAt")) if comments else "",
        "replies": [],
    }


CALLCENTER_REPLY_WHO = {"csAgent": "쿠팡 상담원", "vendor": "판매자"}


def normalize_callcenter_inquiry(item):
    """콜센터(CS)문의 1개 → 문의 탭 행 dict. 연락처(buyerPhone·buyerEmail)는 넣지 않는다.

    답변 여부: csPartnerCounselingStatus 가 answered 면 답변, requestAnswer 이면서 진행 중(progress)이면 미답변,
    requestAnswer 인데 종료(complete, 쿠팡 TRANSFER)면 판매자 답변 없이 끝난 것이라 None(알 수 없음)으로 둔다.
    """
    status = _text(item.get("csPartnerCounselingStatus"))
    progress = _text(item.get("inquiryStatus"))
    if status == "answered":
        answered, status_label = True, ""
    elif status == "requestAnswer" and progress == "progress":
        answered, status_label = False, ""
    elif status == "requestAnswer":
        answered, status_label = None, "판매자 답변 없이 종료"
    else:
        answered, status_label = None, ""
    replies = []
    for reply in item.get("replies") or []:
        if not isinstance(reply, dict):
            continue
        answer_type = _text(reply.get("answerType"))
        who = CALLCENTER_REPLY_WHO.get(answer_type, answer_type or "답변")
        if answer_type == "csAgent" and reply.get("needAnswer"):
            who += " · 판매자 답변 요청"
        replies.append({
            "who": who,
            "vendor": answer_type == "vendor",
            "at": short_datetime(reply.get("replyAt")),
            "content": _text(reply.get("content")),
        })
    vendor_replies = [r for r in replies if r["vendor"]]
    return {
        "kind": INQUIRY_KIND_CALLCENTER,
        "store": "쿠팡",
        "id": _text(item.get("inquiryId")),
        "created": short_datetime(item.get("inquiryAt")),
        "writer": "",
        "product": _text(item.get("itemName")),
        "product_no": "",
        "order_id": _text(item.get("orderId")),
        "product_order_ids": [],
        "category": _category(item.get("receiptCategory")),
        "title": "",
        "content": _text(item.get("content")),
        "answered": answered,
        "status_label": status_label,
        "answer": vendor_replies[-1]["content"] if vendor_replies else "",
        "answered_at": short_datetime(item.get("answeredAt")) if answered else "",
        "replies": replies,
    }


def _remember_product_names(found):
    with _product_names_lock:
        _product_names.update(found)


def product_names_for(seller_product_ids, access_key, secret_key, limit=PRODUCT_NAME_LOOKUP_LIMIT):
    """등록 상품 번호 → 상품명. 기억한 이름은 다시 조회하지 않고, 새로 찾는 것은 limit 개까지(앞쪽부터).

    상세 조회가 실패한 상품은 이름 없이 둔다(목록에는 상품 번호만 보인다). 반환: {번호: 이름}.
    """
    wanted = []
    for pid in seller_product_ids:
        pid = _text(pid)
        if pid and pid not in wanted:
            wanted.append(pid)
    with _product_names_lock:
        known = {pid: _product_names[pid] for pid in wanted if pid in _product_names}
    missing = [pid for pid in wanted if pid not in known][:max(0, int(limit))]
    if missing:
        def lookup(pid):
            try:
                data = fetch_seller_product(pid, access_key, secret_key)
            except Exception:
                return pid, None
            return pid, _text(data.get("sellerProductName") or data.get("displayProductName"))

        with ThreadPoolExecutor(max_workers=PRODUCT_NAME_LOOKUP_WORKERS) as pool:
            found = {pid: name for pid, name in pool.map(lookup, missing) if name}
        _remember_product_names(found)
        known.update(found)
    return known


def fetch_inquiry_list(vendor_id, access_key, secret_key, from_dt, to_dt, unanswered_only=True):
    """온라인문의·콜센터문의를 함께(동시에) 조회해 문의 탭 행 목록으로 돌려준다.

    한 종류가 실패해도 다른 종류는 보여 주도록 오류는 errors 에 모은다.
    반환: {rows, errors, sample_keys: {종류: 첫 항목의 키 목록}}.
    unanswered_only 면 미답변만 조회하고, 답변 여부를 알 수 없는 건은 미답변으로 본다.
    """
    answered = False if unanswered_only else None
    sources = (
        (INQUIRY_KIND_ONLINE, fetch_online_inquiries),
        (INQUIRY_KIND_CALLCENTER, fetch_callcenter_inquiries),
    )
    with ThreadPoolExecutor(max_workers=len(sources)) as pool:
        futures = [(kind, pool.submit(fetch, vendor_id, access_key, secret_key, from_dt, to_dt, answered=answered))
                   for kind, fetch in sources]
    fetched = {}
    errors = []
    for kind, future in futures:
        try:
            fetched[kind] = [it for it in future.result() if isinstance(it, dict)]
        except Exception as e:
            errors.append(f"{kind} 조회 실패: {e}")
            fetched[kind] = []
    online = fetched[INQUIRY_KIND_ONLINE]
    names = product_names_for([it.get("sellerProductId") for it in online], access_key, secret_key) if online else {}
    rows = []
    sample_keys = {}
    for kind, normalize in ((INQUIRY_KIND_ONLINE, lambda it: normalize_online_inquiry(it, names)),
                            (INQUIRY_KIND_CALLCENTER, normalize_callcenter_inquiry)):
        items = fetched[kind]
        if items:
            sample_keys[kind] = sorted(str(k) for k in items[0].keys())
        for item in items:
            row = normalize(item)
            if not row["id"]:
                continue
            if unanswered_only and row["answered"] is None:
                row["answered"] = False
            rows.append(row)
    rows.sort(key=lambda r: r["created"], reverse=True)
    return {"rows": rows, "errors": errors, "sample_keys": sample_keys}


def _money_units(value, field_name):
    """쿠팡 Money 객체를 원 단위 정수로 변환한다.

    거래명세서는 원 단위 문서이므로 KRW가 아니거나 소수점이 있으면 조용히
    반올림하지 않고 발급을 중단한다. 금액을 추정해 문서를 만드는 일을 막는다.
    """
    if not isinstance(value, dict):
        raise ValueError(f"{field_name} 금액 정보를 확인할 수 없습니다.")
    currency = str(value.get("currencyCode") or "").upper()
    if currency != "KRW":
        raise ValueError(f"{field_name} 통화가 KRW가 아니어서 거래명세서를 발급할 수 없습니다.")
    try:
        units = int(value.get("units"))
        nanos = int(value.get("nanos") or 0)
    except (TypeError, ValueError):
        raise ValueError(f"{field_name} 금액 형식이 올바르지 않습니다.")
    if nanos:
        raise ValueError(f"{field_name}에 원 단위가 아닌 금액이 있어 거래명세서를 발급할 수 없습니다.")
    if units < 0:
        raise ValueError(f"{field_name} 금액이 음수여서 거래명세서를 발급할 수 없습니다.")
    return units


def build_transaction_statement_order(order_id, order_sheets):
    """쿠팡 발주서 응답을 거래명세서용 주문 데이터로 엄격하게 변환한다.

    품목 금액은 API가 제공하는 주문 상품금액(orderPrice)에서 실제 적용 할인
    (discountPrice)을 뺀 값만 사용한다. 배송비는 묶음배송번호당 한 번만 더한다.
    취소·취소대기·부분취소 주문은 최종 결제액을 추정하지 않도록 발급 대상에서 제외한다.
    """
    requested_id = str(order_id or "").strip()
    if not requested_id:
        raise ValueError("쿠팡 주문번호를 입력해주세요.")
    if not isinstance(order_sheets, list) or not order_sheets:
        raise ValueError("입력한 주문번호의 발주서 정보를 찾지 못했습니다.")

    items = []
    orderer_name = ""
    receiver_name = ""
    payment_date = ""
    shipping_by_box = {}
    for sheet in order_sheets:
        if not isinstance(sheet, dict):
            raise ValueError("쿠팡 발주서 응답 형식이 올바르지 않습니다.")
        response_id = str(sheet.get("orderId") or "").strip()
        if response_id and response_id != requested_id:
            continue
        status = str(sheet.get("status") or "").upper()
        if status in {"CANCEL", "CANCELLED", "RETURN", "RETURNED", "EXCHANGE", "EXCHANGED"}:
            raise ValueError("취소·반품·교환 완료 주문은 거래명세서를 발급할 수 없습니다.")

        orderer = sheet.get("orderer") or {}
        receiver = sheet.get("receiver") or {}
        orderer_name = orderer_name or str(orderer.get("name") or "").strip()
        receiver_name = receiver_name or str(receiver.get("name") or "").strip()
        payment_date = payment_date or str(sheet.get("paidAt") or sheet.get("orderedAt") or "").strip()

        shipment_box_id = str(sheet.get("shipmentBoxId") or "").strip()
        if not shipment_box_id:
            raise ValueError("묶음배송번호를 확인할 수 없어 배송비를 정확히 계산할 수 없습니다.")
        shipping = _money_units(sheet.get("shippingPrice"), "기본 배송비")
        remote = _money_units(sheet.get("remotePrice"), "도서산간 배송비")
        shipping_by_box[shipment_box_id] = max(shipping_by_box.get(shipment_box_id, 0), shipping + remote)

        order_items = sheet.get("orderItems")
        if not isinstance(order_items, list) or not order_items:
            raise ValueError("발주서에서 주문 상품을 찾지 못했습니다.")
        for source_item in order_items:
            if not isinstance(source_item, dict):
                raise ValueError("주문 상품 정보 형식이 올바르지 않습니다.")
            cancelled = bool(source_item.get("canceled"))
            try:
                cancel_count = int(source_item.get("cancelCount") or 0)
                hold_cancel_count = int(source_item.get("holdCountForCancel") or 0)
                quantity = int(source_item.get("shippingCount") or 0)
            except (TypeError, ValueError):
                raise ValueError("주문 상품 수량 정보를 확인할 수 없습니다.")
            if cancelled or cancel_count or hold_cancel_count:
                raise ValueError("부분취소 또는 취소 처리 중인 주문은 거래명세서를 발급할 수 없습니다.")
            if quantity <= 0:
                raise ValueError("주문 상품 수량을 확인할 수 없습니다.")
            order_price = _money_units(source_item.get("orderPrice"), "상품 주문금액")
            discount_price = _money_units(source_item.get("discountPrice"), "상품 할인금액")
            if discount_price > order_price:
                raise ValueError("상품 할인금액이 주문금액보다 커서 거래명세서를 발급할 수 없습니다.")
            paid_amount = order_price - discount_price
            if paid_amount <= 0:
                raise ValueError("실제 결제 상품금액이 0원이어서 거래명세서를 발급할 수 없습니다.")
            name = str(source_item.get("sellerProductName") or source_item.get("vendorItemName") or "").strip()
            if not name:
                raise ValueError("주문 상품명을 확인할 수 없습니다.")
            specification = str(source_item.get("sellerProductItemName") or "").strip()
            items.append({
                "name": name,
                "specification": specification,
                "quantity": quantity,
                "gross_unit_price": (paid_amount + quantity // 2) // quantity,
                "gross_amount": paid_amount,
                "status": status,
            })

    if not items:
        raise ValueError("입력한 주문번호의 발급 가능한 상품을 찾지 못했습니다.")
    if not (orderer_name or receiver_name):
        raise ValueError("주문자명과 수취인명을 확인할 수 없어 거래명세서를 발급할 수 없습니다.")
    if not payment_date:
        raise ValueError("결제일시를 확인할 수 없어 거래명세서를 발급할 수 없습니다.")
    return {
        "order_id": requested_id,
        "orderer_name": orderer_name,
        "receiver_name": receiver_name,
        "customer_name": orderer_name or receiver_name,
        "payment_date": payment_date,
        "shipping_gross": sum(shipping_by_box.values()),
        "items": items,
    }


def fetch_order_for_transaction_statement(vendor_id, access_key, secret_key, order_id):
    """쿠팡 주문번호로 발주서 전체를 조회해 거래명세서용 데이터로 변환한다."""
    normalized_order_id = str(order_id or "").strip()
    if not normalized_order_id:
        raise ValueError("쿠팡 주문번호를 입력해주세요.")
    path = f"{PATH_PREFIX}/{vendor_id}/{normalized_order_id}/ordersheets"
    response = _get(path, {}, access_key, secret_key)
    data = response.get("data") if isinstance(response, dict) else None
    return build_transaction_statement_order(normalized_order_id, data)


SELLER_PRODUCTS_PATH = "/v2/providers/seller_api/apis/api/v1/marketplace/seller-products"
SELLER_PRODUCT_PAGE_SIZE = 100   # 상품 목록 API 최대값
SELLER_PRODUCT_MAX_PAGES = 200   # 2만 개까지(과도 호출 방지)


def fetch_seller_products(vendor_id, access_key, secret_key, max_pages=SELLER_PRODUCT_MAX_PAGES):
    """등록 상품 목록(최근 등록 순)을 반환한다. 항목에는 sellerProductId·productId·상품명·상태만 있고
    옵션 ID(vendorItemId)는 없다 — 옵션은 fetch_seller_product로 상품마다 조회한다."""
    products = []
    next_token = ""
    for _ in range(max_pages):
        params = {"vendorId": vendor_id, "maxPerPage": SELLER_PRODUCT_PAGE_SIZE}
        if next_token:
            params["nextToken"] = next_token
        js = _get(SELLER_PRODUCTS_PATH, params, access_key, secret_key)
        page = js.get("data") or []
        products.extend(page)
        next_token = str(js.get("nextToken") or "")
        if not next_token or not page:
            break
    return products


def fetch_seller_product(seller_product_id, access_key, secret_key):
    """등록 상품 하나의 상세(items[].vendorItemId 옵션 포함)를 반환한다. 한 번에 약 0.7초 걸린다."""
    js = _get(f"{SELLER_PRODUCTS_PATH}/{seller_product_id}", {}, access_key, secret_key)
    return js.get("data") or {}


def validate_credentials(vendor_id, access_key, secret_key):
    """키 3종으로 최근 1일 온라인문의를 1회 조회해 유효성만 확인. 반환 {ok, valid, error}."""
    vid = str(vendor_id or "").strip()
    ak = str(access_key or "").strip()
    sk = str(secret_key or "").strip()
    if not (vid and ak and sk):
        return {"ok": True, "valid": False,
                "error": "vendorId / accessKey / secretKey 중 빈 값이 있습니다."}
    try:
        now = datetime.now()
        fetch_online_inquiries(vid, ak, sk, now - timedelta(days=1), now, answered=False)
        return {"ok": True, "valid": True, "error": ""}
    except Exception as e:
        return {"ok": True, "valid": False, "error": str(e)}
