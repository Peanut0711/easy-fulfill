"""
Google Sheets용 OAuth(gspread) 공통 모듈.

easy-fulfill.py, database-sync.py에서 동일한 google-oauth/ 경로와 token.json을 사용합니다.
database-sync는 시트 쓰기가 필요하므로 스코프는 spreadsheets(전체)입니다.
기존 token.json이 spreadsheets.readonly로만 발급된 경우, token.json을 삭제한 뒤 재로그인하세요.
"""

from __future__ import annotations

from pathlib import Path

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

APP_ROOT = Path(__file__).resolve().parent
GOOGLE_AUTH_DIR = APP_ROOT / "google-oauth"
OAUTH_CREDENTIAL_PATH = GOOGLE_AUTH_DIR / "credentials.json"
TOKEN_PATH = GOOGLE_AUTH_DIR / "token.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _is_invalid_grant_error(error: BaseException) -> bool:
    """리프레시 토큰 폐기/만료(invalid_grant) 류 오류인지 문자열 기반 판별."""
    text = str(error).lower()
    return (
        "invalid_grant" in text
        or "token has been expired or revoked" in text
        or "token has expired or revoked" in text
    )


def get_authorized_gspread_client():
    """OAuth로 로그인한 gspread Client를 반환합니다. 브라우저 로그인이 필요할 수 있습니다."""
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
            flow = InstalledAppFlow.from_client_secrets_file(
                str(OAUTH_CREDENTIAL_PATH), SCOPES
            )
            creds = flow.run_local_server(port=0)

        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")

    return gspread.authorize(creds)


def delete_oauth_token_file() -> bool:
    """로컬 token.json만 삭제합니다. 있어서 지웠으면 True, 원래 없으면 False."""
    if TOKEN_PATH.exists():
        TOKEN_PATH.unlink()
        return True
    return False


def get_oauth_status_description() -> str:
    """환경설정 등에 표시할 Google Sheets OAuth 상태 설명(짧은 단락)."""
    if not OAUTH_CREDENTIAL_PATH.exists():
        return (
            "OAuth 클라이언트 파일(credentials.json)이 없습니다.\n"
            f"경로: {OAUTH_CREDENTIAL_PATH}"
        )
    if not TOKEN_PATH.exists():
        return (
            "저장된 Google 로그인이 없습니다.\n"
            "「재인증」으로 로그인하거나, 주문 엑셀을 불러올 때 브라우저 로그인이 열릴 수 있습니다."
        )
    try:
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    except Exception:
        return "token.json을 읽을 수 없습니다.\n「연결 해제」 후 「재인증」을 시도하세요."

    if creds.valid:
        return "상태: 연결됨 (저장된 토큰이 유효합니다)."
    if creds.expired and creds.refresh_token:
        return (
            "상태: 토큰 만료됨 — 사용 시 자동 갱신을 시도합니다.\n"
            "문제가 계속되면 「연결 해제」 후 「재인증」하세요."
        )
    return "상태: 로그인 정보가 불완전합니다.\n「재인증」이 필요합니다."
