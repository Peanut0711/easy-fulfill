"""슬랙 Incoming Webhook 전송 (UI 비의존 순수 함수).

webhook_url 1개 = 채널 1개 고정. JSON {"text": ...} 를 POST 하면 그 채널에 글이 올라온다.
webhook_url 은 비밀값이므로 코드/레포에 두지 말 것(공유 「설정」 탭·로컬 설정에만 저장).
"""

try:
    import requests
except ImportError:
    requests = None

DEFAULT_TIMEOUT = 10


def send_slack(webhook_url, text):
    """슬랙 채널에 텍스트 메시지를 전송. 반환 {ok} 또는 {ok False, error}."""
    if requests is None:
        return {"ok": False, "error": "requests 패키지가 필요합니다."}
    url = str(webhook_url or "").strip()
    if not url:
        return {"ok": False, "error": "슬랙 웹훅 URL이 비어 있습니다."}
    try:
        resp = requests.post(url, json={"text": text}, timeout=DEFAULT_TIMEOUT)
        resp.raise_for_status()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}
