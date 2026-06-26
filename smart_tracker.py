"""스마트택배(sweettracker) 배송조회 API 래퍼.

UI/Qt에 의존하지 않는 순수 함수 모듈. easy-fulfill 의 백그라운드 스레드에서 호출한다.

- companylist : 택배사 코드 목록(우체국 t_code 확인용)
- trackingInfo: 단일 등기번호 배송 상태 조회

주의: 엔드포인트·응답 필드는 발급받은 API 문서로 최종 확인이 필요하다.
일반적으로 trackingInfo 응답은 completeYN("Y"/"N"), level(1~6, 6=배송완료),
trackingDetails[](단계별 내역), lastDetail(최근 내역)을 포함한다.
"""

import time

try:
    import requests
except ImportError:  # requests 는 이미 의존성이지만 방어적으로 처리
    requests = None

API_BASE = "https://info.sweettracker.co.kr/api/v1"
DEFAULT_TIMEOUT = 10
# 우체국택배 기본 코드(대부분 "04"). companylist 로 최종 확정 권장.
DEFAULT_KPOST_T_CODE = "04"


def _require_requests():
    if requests is None:
        raise RuntimeError("requests 패키지가 필요합니다. (pip install requests)")


def fetch_company_list(t_key):
    """택배사 코드 목록을 반환. 반환: {"ok": bool, "companies": [{"Code","Name"}...], "error"}"""
    _require_requests()
    try:
        resp = requests.get(
            f"{API_BASE}/companylist",
            params={"t_key": t_key},
            timeout=DEFAULT_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
        companies = data.get("Company", []) if isinstance(data, dict) else []
        return {"ok": True, "companies": companies}
    except Exception as e:
        return {"ok": False, "error": str(e), "companies": []}


def find_kpost_t_code(t_key):
    """companylist 에서 우체국 택배사 코드를 찾습니다. 못 찾으면 기본값."""
    result = fetch_company_list(t_key)
    if not result.get("ok"):
        return DEFAULT_KPOST_T_CODE
    for c in result.get("companies", []):
        name = str(c.get("Name", ""))
        if "우체국" in name:
            return str(c.get("Code", DEFAULT_KPOST_T_CODE))
    return DEFAULT_KPOST_T_CODE


def fetch_tracking_info(t_key, t_code, t_invoice):
    """단일 등기번호 배송 상태 조회. 원본 응답 dict 를 반환(에러 시 status=False 포함)."""
    _require_requests()
    try:
        resp = requests.get(
            f"{API_BASE}/trackingInfo",
            params={"t_key": t_key, "t_code": t_code, "t_invoice": t_invoice},
            timeout=DEFAULT_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        return {"status": False, "error": str(e)}


def is_delivery_complete(payload):
    """배송완료 여부 판정. completeYN=="Y" 우선, 보조로 level>=6."""
    if not isinstance(payload, dict):
        return False
    complete = str(payload.get("completeYN", "")).strip().upper()
    if complete == "Y":
        return True
    if complete == "N":
        return False
    try:
        return int(payload.get("level", 0)) >= 6
    except (TypeError, ValueError):
        return False


def _last_detail(payload):
    """가장 최근 배송 내역 1건을 추출."""
    if not isinstance(payload, dict):
        return {}
    last = payload.get("lastDetail")
    if isinstance(last, dict) and last:
        return last
    details = payload.get("trackingDetails")
    if isinstance(details, list) and details:
        tail = details[-1]
        if isinstance(tail, dict):
            return tail
    return {}


def summarize_tracking(payload):
    """시트 기록용 요약. 반환: {ok, complete, status, where, time, error}."""
    if not isinstance(payload, dict):
        return {"ok": False, "complete": False, "status": "", "where": "",
                "time": "", "error": "응답 형식 오류"}
    # API 레벨 오류(status=False 또는 잘못된 키/송장)
    if payload.get("status") is False or payload.get("error"):
        err = payload.get("error") or payload.get("msg") or "조회 실패"
        return {"ok": False, "complete": False, "status": "", "where": "",
                "time": "", "error": str(err)}
    last = _last_detail(payload)
    # 상태 텍스트: kind(예: '배달완료') 우선, 없으면 level 기반
    status = str(last.get("kind", "") or "").strip()
    if not status:
        status = "배송완료" if is_delivery_complete(payload) else "배송중"
    where = str(last.get("where", "") or "").strip()
    when = str(last.get("timeString", "") or last.get("time", "") or "").strip()
    return {
        "ok": True,
        "complete": is_delivery_complete(payload),
        "status": status,
        "where": where,
        "time": when,
        "error": "",
    }


def run_tracking_batch_worker(t_key, t_code, invoices, sleep_sec=0.25):
    """여러 등기번호를 순차 조회해 요약 리스트를 반환(시트 미참조).

    invoices: [등기번호(str) ...]
    반환: {"ok": bool, "results": {등기번호: summary}, "error"}
    """
    if requests is None:
        return {"ok": False, "results": {}, "error": "requests 패키지가 필요합니다."}
    if not t_key:
        return {"ok": False, "results": {}, "error": "스마트택배 API 키가 설정되지 않았습니다."}
    code = t_code or DEFAULT_KPOST_T_CODE
    results = {}
    for inv in invoices:
        inv = str(inv).strip()
        if not inv:
            continue
        payload = fetch_tracking_info(t_key, code, inv)
        results[inv] = summarize_tracking(payload)
        if sleep_sec:
            time.sleep(sleep_sec)
    return {"ok": True, "results": results, "error": ""}
