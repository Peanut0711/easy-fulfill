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


def get_authorized_gspread_client():
    """OAuth로 로그인한 gspread Client를 반환합니다. 브라우저 로그인이 필요할 수 있습니다."""
    GOOGLE_AUTH_DIR.mkdir(parents=True, exist_ok=True)

    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
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
