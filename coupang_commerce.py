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
import time
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


def _paged(path, base_params, access_key, secret_key):
    """pageNum 을 늘려가며 모든 페이지 항목을 모은다."""
    items = []
    page = 1
    while page <= MAX_PAGES:
        params = dict(base_params)
        params["pageNum"] = page
        params["pageSize"] = PAGE_SIZE
        js = _get(path, params, access_key, secret_key)
        page_items = _extract_items(js)
        items.extend(page_items)
        if len(page_items) < PAGE_SIZE:
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
    """콜센터(CS) 문의 목록을 반환. answered=False → 미답변만(NO_ANSWER)."""
    path = f"{PATH_PREFIX}/{vendor_id}/callCenterInquiries"
    status = None if answered is None else ("ANSWER" if answered else "NO_ANSWER")
    out = []
    for s, e in _date_windows(from_dt, to_dt):
        base = {
            "vendorId": vendor_id,
            "inquiryStartAt": s.strftime("%Y-%m-%d"),
            "inquiryEndAt": e.strftime("%Y-%m-%d"),
        }
        if status:
            base["partnerCounselingStatus"] = status
        out.extend(_paged(path, base, access_key, secret_key))
    return out


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
