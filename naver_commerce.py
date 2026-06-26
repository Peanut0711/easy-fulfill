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
