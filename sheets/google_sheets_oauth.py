"""
Google Sheets용 OAuth(gspread) 공통 모듈.

easy-fulfill.py, database-sync.py에서 동일한 google-oauth/ 경로와 token.json을 사용합니다.
database-sync는 시트 쓰기가 필요하므로 스코프는 spreadsheets(전체)입니다.
기존 token.json이 spreadsheets.readonly로만 발급된 경우, token.json을 삭제한 뒤 재로그인하세요.
"""

from __future__ import annotations

import threading

from google.oauth2.credentials import Credentials

from shared.paths import APP_ROOT

# gspread·google_auth_oauthlib(요청·로컬 웹서버 포함)은 무거워(약 0.3초) 실제 로그인·연결할 때 불러온다.
# 앱 시작 시에는 연동 상태 표시(get_oauth_status)만 쓰므로 Credentials만 미리 불러온다.

GOOGLE_AUTH_DIR = APP_ROOT / "google-oauth"
OAUTH_CREDENTIAL_PATH = GOOGLE_AUTH_DIR / "credentials.json"
TOKEN_PATH = GOOGLE_AUTH_DIR / "token.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
LOGIN_TIMEOUT_SECONDS = 300

# 브라우저 로그인 허용 기본값. GUI는 시작할 때 False로 바꿔 백그라운드 작업이 로그인 창을
# 열지 않게 하고(재인증 버튼만 interactive=True), CLI 스크립트(database-sync 등)는 그대로 연다.
_interactive_login_default = True
_login_lock = threading.Lock()  # 브라우저 로그인은 한 번에 하나만


class GoogleLoginRequiredError(RuntimeError):
    """저장된 토큰으로 로그인할 수 없어 사용자의 Google 재인증이 필요하다."""

    def __init__(self, message: str | None = None):
        super().__init__(message or (
            "Google 로그인이 필요합니다. 저장된 로그인(token.json)이 없거나 만료되었습니다. "
            "「⚙ 설정」→「연동」에서 「재인증」을 눌러 주세요."))


def set_interactive_login_default(enabled: bool) -> bool:
    """interactive 인자를 생략한 호출의 브라우저 로그인 허용 여부를 바꾸고 이전 값을 반환한다."""
    global _interactive_login_default
    previous = _interactive_login_default
    _interactive_login_default = bool(enabled)
    return previous


def _is_invalid_grant_error(error: BaseException) -> bool:
    """리프레시 토큰 폐기/만료(invalid_grant) 류 오류인지 문자열 기반 판별."""
    text = str(error).lower()
    return (
        "invalid_grant" in text
        or "token has been expired or revoked" in text
        or "token has expired or revoked" in text
    )


def get_authorized_gspread_client(interactive: bool | None = None,
                                  timeout_seconds: int = LOGIN_TIMEOUT_SECONDS):
    """OAuth로 로그인한 gspread Client를 반환합니다.

    저장된 토큰을 쓸 수 없으면 interactive일 때만 브라우저 로그인을 열고(최대 timeout_seconds,
    동시에 하나만), 아니면 GoogleLoginRequiredError를 냅니다.
    """
    import gspread
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow

    if interactive is None:
        interactive = _interactive_login_default
    GOOGLE_AUTH_DIR.mkdir(parents=True, exist_ok=True)

    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                # 리프레시 토큰이 만료/폐기되면 기존 token.json을 지우고 재인증으로 복구한다.
                if not _is_invalid_grant_error(e):
                    raise
                delete_oauth_token_file()
                creds = None
        if not creds or not creds.valid:
            if not OAUTH_CREDENTIAL_PATH.exists():
                raise FileNotFoundError(
                    "OAuth 클라이언트 파일이 없습니다.\n"
                    f"다음 경로에 credentials.json을 배치하세요:\n{OAUTH_CREDENTIAL_PATH}"
                )
            if not interactive:
                raise GoogleLoginRequiredError()
            if not _login_lock.acquire(blocking=False):
                raise GoogleLoginRequiredError(
                    "다른 Google 로그인이 진행 중입니다. 열린 브라우저에서 로그인을 마친 뒤 다시 시도해 주세요.")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(OAUTH_CREDENTIAL_PATH), SCOPES
                )
                try:
                    creds = flow.run_local_server(port=0, timeout_seconds=timeout_seconds)
                except AttributeError as error:
                    # 제한 시간 안에 브라우저 응답이 없으면 라이브러리가 None을 다루다 실패한다.
                    raise GoogleLoginRequiredError(
                        f"Google 로그인 대기 시간({timeout_seconds}초)이 지났습니다. 다시 재인증해 주세요.",
                    ) from error
            finally:
                _login_lock.release()
            if creds is None:
                raise GoogleLoginRequiredError("Google 로그인이 완료되지 않았습니다. 다시 재인증해 주세요.")

        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")

    # 공유 시트의 최소 버전보다 낮은 앱은 모든 쓰기 요청을 보내기 전에 막는다(sheets.version_gate).
    from sheets.version_gate import install_write_guard
    return install_write_guard(gspread.authorize(creds))


def delete_oauth_token_file() -> bool:
    """로컬 token.json만 삭제합니다. 있어서 지웠으면 True, 원래 없으면 False."""
    if TOKEN_PATH.exists():
        TOKEN_PATH.unlink()
        return True
    return False


def get_oauth_status() -> tuple[str, str, str]:
    """설정 탭에 표시할 Google Sheets OAuth 상태 ``(kind, text, detail)``.

    ``kind``는 ``shared.ui_style.set_status``의 상태(ok·warn·error)다. 파일만 읽고 네트워크는 쓰지 않는다.
    액세스 토큰은 한 시간마다 만료되고 사용할 때 자동 갱신되므로, 갱신 토큰이 있으면 연결됨으로 본다.
    """
    if not OAUTH_CREDENTIAL_PATH.exists():
        return ("error", "OAuth 클라이언트 파일(credentials.json)이 없습니다.",
                f"경로: {OAUTH_CREDENTIAL_PATH}")
    if not TOKEN_PATH.exists():
        return ("warn", "이 PC에 저장된 Google 로그인이 없습니다.",
                "「재인증」으로 로그인하거나, 주문 엑셀을 불러올 때 브라우저 로그인이 열릴 수 있습니다.")
    try:
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    except Exception:
        return ("error", "token.json을 읽을 수 없습니다.", "「연결 해제」 후 「재인증」을 시도하세요.")

    if creds.valid:
        return ("ok", "연결됨", "저장된 토큰이 유효합니다.")
    if creds.expired and creds.refresh_token:
        return ("ok", "연결됨 (사용할 때 토큰 자동 갱신)",
                "저장된 토큰이 만료되어 사용할 때 자동 갱신합니다.\n"
                "문제가 계속되면 「연결 해제」 후 「재인증」하세요.")
    return ("error", "로그인 정보가 불완전합니다.", "「재인증」이 필요합니다.")


def get_oauth_status_description() -> str:
    """Google Sheets OAuth 상태 설명(짧은 단락). 설정 탭은 ``get_oauth_status``를 쓴다."""
    _kind, text, detail = get_oauth_status()
    return f"{text}\n{detail}" if detail else text
