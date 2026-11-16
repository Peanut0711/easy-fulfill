"""우체국 OpenAPI(KpostPortal) 국내우편물 종추적조회 래퍼.

UI/Qt에 비의존하는 순수 함수 모듈. 응답은 XML이다.
요청 예:
  http://biz.epost.go.kr/KpostPortal/openapi?regkey=<30자리>&target=trace&query=<등기번호 13자리>&showRec=Y

요청 변수: regkey(인증키), target=trace(국내우편물), query=등기번호(13자리), showRec=Y(접수정보 포함)
주요 응답 필드(종추적): regino(등기번호), recevnm(수취인), eventnm(배달결과-상태),
  eventregiponm(현재 위치=우체국명), eventymd/eventhms(이벤트 날짜·시간),
  tracestatus(처리현황), delivrsltnm(배달결과 상세), nondelivreasnnm(미배달사유)
오류 XML: <error><error_code>ERR-xxx</error_code><message>..</message></error>
  ERR-001 조회결과 없음 / ERR-123 미등록·무효 인증키 / ERR-131 시스템 부하 차단 /
  ERR-321 등기번호 13자리 / ERR-125 등기번호 없음·무효
"""

import xml.etree.ElementTree as ET

try:
    import requests
except ImportError:  # requests 는 이미 의존성이지만 방어적으로 처리
    requests = None

KPOST_TRACE_URL = "http://biz.epost.go.kr/KpostPortal/openapi"
DEFAULT_TIMEOUT = 12
COMPLETE_KEYWORDS = ("배달완료", "배달 완료")


def _local(tag):
    return tag.split("}", 1)[-1].lower() if tag else ""


def _require_requests():
    if requests is None:
        raise RuntimeError("requests 패키지가 필요합니다. (pip install requests)")


def fetch_tracking_text(regkey, regino, show_rec=True):
    """종추적조회 XML 원문을 반환. 네트워크/HTTP 실패 시 예외 발생."""
    _require_requests()
    params = {"regkey": regkey, "target": "trace", "query": str(regino).strip()}
    if show_rec:
        params["showRec"] = "Y"
    resp = requests.get(KPOST_TRACE_URL, params=params, timeout=DEFAULT_TIMEOUT)
    resp.raise_for_status()
    return resp.content.decode("utf-8", errors="replace")


def _parse_error(root):
    for el in root.iter():
        if _local(el.tag) != "error":
            continue
        code, msg = "", ""
        for ch in el:
            ln = _local(ch.tag)
            if ln == "error_code":
                code = (ch.text or "").strip()
            elif ln == "message":
                msg = (ch.text or "").strip()
        return code or "오류", msg or "알 수 없는 오류"
    return None


def _is_admin_receipt(ev):
    """선구분-후접수 등 '행정상 접수' 이벤트인지. (우체국 웹도 현재상태에서 제외)"""
    return ev.get("status") == "접수" and (
        "선구분" in ev.get("reason", "") or "후접수" in ev.get("reason", "")
    )


def parse_tracking(xml_text):
    """종추적 XML → 요약 dict.

    실제 구조: <trace> ... <itemlist><item> 안에 sortingdate/eventhms/
    eventregiponm/tracestatus/nondelivreasnnm 가 단계별로 반복. 최상위 eventnm/
    eventymd 는 '배달결과' 요약(배달완료 시에만 채워짐).

    반환: {ok, error_code, error, complete, status, where, time, recipient, events}
      events: [{date, time, where, status, reason} ...] 시간순
    """
    base = {"ok": False, "error_code": "", "error": "", "complete": False,
            "status": "", "where": "", "time": "", "recipient": "", "events": []}
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        base.update(error_code="XML", error="응답을 XML로 해석할 수 없습니다.")
        return base
    err = _parse_error(root)
    if err:
        base.update(error_code=err[0], error=err[1])
        return base

    recipient = ""
    top_eventnm = ""
    events = []
    for el in root.iter():
        ln = _local(el.tag)
        if ln == "recevnm" and not recipient:
            recipient = (el.text or "").strip()
        elif ln == "eventnm" and not top_eventnm:
            top_eventnm = (el.text or "").strip()
        elif ln == "item":
            ev = {"date": "", "time": "", "where": "", "status": "", "reason": ""}
            for ch in el:
                cln = _local(ch.tag)
                t = (ch.text or "").strip()
                if cln == "sortingdate":
                    ev["date"] = t
                elif cln == "eventhms":
                    ev["time"] = t
                elif cln == "eventregiponm":
                    ev["where"] = t
                elif cln == "tracestatus":
                    ev["status"] = t
                elif cln == "nondelivreasnnm":
                    ev["reason"] = t
            if any(ev.values()):
                events.append(ev)

    events.sort(key=lambda e: (e.get("date", ""), e.get("time", "")))
    statuses = [e["status"] for e in events]
    complete = any(k in top_eventnm for k in COMPLETE_KEYWORDS) or any(
        any(k in s for k in COMPLETE_KEYWORDS) for s in statuses
    )

    if complete:
        done = next((e for e in reversed(events)
                     if any(k in e["status"] for k in COMPLETE_KEYWORDS)), None)
        cur = done or (events[-1] if events else None)
        status = "배달완료"
    else:
        # 현재상태: 선구분-후접수 같은 행정 접수는 제외하고 가장 최신(우체국 웹과 동일)
        meaningful = [e for e in events if not _is_admin_receipt(e)]
        cur = meaningful[-1] if meaningful else (events[-1] if events else None)
        status = cur["status"] if cur else ""

    where = cur["where"] if cur else ""
    when = ((cur["date"] + " " + cur["time"]).strip()) if cur else ""
    base.update(ok=True, complete=complete, status=status, where=where,
                time=when, recipient=recipient, events=events)
    return base


def summarize_tracking(regkey, regino):
    """단건 조회 후 요약 반환(네트워크 예외도 dict로 변환)."""
    try:
        text = fetch_tracking_text(regkey, regino)
    except Exception as e:
        out = parse_tracking("")  # base 형태
        out.update(error_code="HTTP", error=str(e))
        return out
    return parse_tracking(text)


def validate_key(regkey):
    """regkey 유효성 검증. 샘플 등기번호로 호출해 ERR-123(미등록·무효 키)이면 무효,
    그 외(ERR-001/ERR-321/ERR-125/정상)는 인증 통과로 본다.
    반환: {ok, valid, error}."""
    key = str(regkey or "").strip()
    if not key:
        return {"ok": True, "valid": False, "error": "키가 비어 있습니다."}
    try:
        text = fetch_tracking_text(key, "1234567890123")
    except Exception as e:
        return {"ok": False, "valid": False, "error": str(e)}
    res = parse_tracking(text)
    code = (res.get("error_code") or "").upper()
    if code == "ERR-123":
        return {"ok": True, "valid": False,
                "error": "등록되지 않았거나 유효하지 않은 인증키입니다(종추적 서비스 미신청일 수 있음)."}
    return {"ok": True, "valid": True, "error": ""}
