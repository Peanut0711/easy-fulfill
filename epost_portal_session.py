"""우체국 계약고객전용시스템 전용 Chromium 로그인 세션 연결 도구."""

from __future__ import annotations

import argparse
import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from ctypes import wintypes
from dataclasses import dataclass
from typing import Mapping
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parent
PROFILE_DIR = ROOT / "output" / "epost-browser-profile"
AUTH_STATE_PATH = PROFILE_DIR / "easy-fulfill-auth-state.bin"
PORTAL_URL = "https://biz.epost.go.kr/ui/index.jsp"
LOGIN_TIMEOUT_SECONDS = 600
SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"
CONFIG_SHEET_TITLE = "설정"
CONFIG_SHEET_HEADERS = ["키", "값"]
CONFIG_KEY_PORTAL_MEMBER_ID = "epost_portal_member_id"
CONFIG_KEY_PORTAL_PASSWORD = "epost_portal_password"
LOGIN_MEMBER_ID_SELECTOR = "#mainframe_VFrameSet_exFrame_form_divMain_divLogin_edtUserid_input"
LOGIN_PASSWORD_SELECTOR = "#mainframe_VFrameSet_exFrame_form_divMain_divLogin_edtUserpw_input"
LOGIN_BUTTON_SELECTOR = "#mainframe_VFrameSet_exFrame_form_divMain_divLogin_btnLogin"


class PortalCredentialError(RuntimeError):
    """공유 설정 탭의 우체국 포털 로그인 정보가 없거나 읽히지 않을 때의 오류."""


@dataclass(frozen=True)
class PortalCredentials:
    member_id: str
    password: str


def portal_credentials_from_settings(settings: Mapping[str, object]) -> PortalCredentials:
    """설정값에서 포털 로그인에 필요한 두 값만 꺼낸다.

    호출자와 예외 메시지 모두 실제 ID·비밀번호를 출력하지 않는다.
    """
    member_id = str(settings.get(CONFIG_KEY_PORTAL_MEMBER_ID, "") or "").strip()
    password = str(settings.get(CONFIG_KEY_PORTAL_PASSWORD, "") or "").strip()
    missing = []
    if not member_id:
        missing.append(CONFIG_KEY_PORTAL_MEMBER_ID)
    if not password:
        missing.append(CONFIG_KEY_PORTAL_PASSWORD)
    if missing:
        raise PortalCredentialError(
            "설정 탭의 우체국 포털 로그인 정보가 비어 있습니다: " + ", ".join(missing),
        )
    return PortalCredentials(member_id=member_id, password=password)


def load_portal_credentials() -> PortalCredentials:
    """접근 제어된 기존 설정 탭에서만 포털 로그인 정보를 읽는다."""
    try:
        from google_sheets_oauth import get_authorized_gspread_client
        from tracking.repository import open_config_worksheet, read_config_values_map

        worksheet = open_config_worksheet(
            get_authorized_gspread_client(),
            SPREADSHEET_ID,
            CONFIG_SHEET_TITLE,
            CONFIG_SHEET_HEADERS,
        )
        return portal_credentials_from_settings(read_config_values_map(worksheet))
    except PortalCredentialError:
        raise
    except Exception as error:
        raise PortalCredentialError(
            "우체국 포털 로그인 정보를 설정 탭에서 읽지 못했습니다. "
            "Google Sheets 연결과 설정 탭 접근 권한을 확인해 주세요.",
        ) from error


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _data_blob(data: bytes):
    buffer = (ctypes.c_byte * len(data)).from_buffer_copy(data)
    return _DataBlob(len(data), buffer), buffer


def _protect_for_current_windows_user(data: bytes) -> bytes:
    source, source_buffer = _data_blob(data)
    protected = _DataBlob()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(source), None, None, None, None, 1, ctypes.byref(protected)
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(protected.pbData, protected.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(protected.pbData)


def _unprotect_for_current_windows_user(data: bytes) -> bytes:
    source, source_buffer = _data_blob(data)
    plain = _DataBlob()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(source), None, None, None, None, 1, ctypes.byref(plain)
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(plain.pbData, plain.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(plain.pbData)


def page_has_logged_in_state(page) -> bool:
    """포털 화면에 로그인 완료 표식이 보이는지만 확인한다."""
    try:
        return "로그아웃" in page.locator("body").inner_text(timeout=3_000)
    except Exception:
        return False


def wait_for_logged_in_state(page, timeout_seconds: int) -> bool:
    """SPA 화면이 그려질 시간을 주며 로그인 완료 표식을 기다린다."""
    deadline = time.monotonic() + max(0, timeout_seconds)
    while time.monotonic() < deadline:
        if page_has_logged_in_state(page):
            return True
        time.sleep(0.25)
    return page_has_logged_in_state(page)


def login_with_config_credentials(page, timeout_seconds: int) -> str:
    """저장된 세션이 없을 때 설정 탭의 자격증명으로만 로그인한다."""
    credentials = load_portal_credentials()
    try:
        page.locator(LOGIN_MEMBER_ID_SELECTOR).wait_for(state="visible", timeout=10_000)
        page.locator(LOGIN_MEMBER_ID_SELECTOR).fill(credentials.member_id)
        page.locator(LOGIN_PASSWORD_SELECTOR).fill(credentials.password)
        page.locator(LOGIN_BUTTON_SELECTOR).click()
    except Exception as error:
        raise RuntimeError(
            "우체국 포털 로그인 화면의 입력칸 또는 로그인 버튼을 찾지 못했습니다. "
            "포털 로그인 화면이 변경됐는지 확인해 주세요.",
        ) from error

    if wait_for_logged_in_state(page, timeout_seconds):
        return "config_credentials"
    raise RuntimeError(
        "우체국 포털 로그인 완료를 확인하지 못했습니다. "
        "설정 탭의 ID·비밀번호와 추가 인증 필요 여부를 확인해 주세요.",
    )


def ensure_portal_login(page, timeout_seconds: int) -> str:
    """유효한 세션을 우선 쓰고, 없으면 설정 탭 자격증명으로 로그인한다."""
    # 계약고객전용시스템은 초기 HTML 로드 뒤 화면을 비동기로 구성한다. 즉시 판정하면
    # 이미 유효한 세션도 로그인 화면으로 오인할 수 있으므로, 먼저 짧게 기다린다.
    if wait_for_logged_in_state(page, min(timeout_seconds, 12)):
        return "saved_session"
    return login_with_config_credentials(page, timeout_seconds)


def save_epost_session(context, page) -> None:
    """쿠키·세션 저장소를 현재 Windows 사용자만 읽을 수 있게 보관한다."""
    if os.name != "nt":
        raise RuntimeError("우체국 포털 로그인 세션 저장은 Windows에서만 지원합니다.")
    state = context.storage_state()
    origin = urlparse(page.url)
    if origin.scheme == "https" and origin.hostname and origin.hostname.endswith("epost.go.kr"):
        state["easyFulfillSessionStorage"] = {
            f"{origin.scheme}://{origin.netloc}": page.evaluate(
                "Object.fromEntries(Object.entries(sessionStorage))",
            ),
        }
    AUTH_STATE_PATH.write_bytes(
        _protect_for_current_windows_user(json.dumps(state, ensure_ascii=False).encode("utf-8")),
    )


def restore_epost_session(context) -> bool:
    """새 Chromium 실행 전에 보호된 쿠키·로컬/세션 저장소를 복원한다."""
    if not AUTH_STATE_PATH.exists():
        return False
    try:
        state = json.loads(_unprotect_for_current_windows_user(AUTH_STATE_PATH.read_bytes()).decode("utf-8"))
        cookies = state.get("cookies", [])
        for cookie in cookies:
            if cookie.get("expires", -1) < 0:
                cookie.pop("expires", None)
        if cookies:
            context.add_cookies(cookies)
        for origin_state in state.get("origins", []):
            origin = origin_state.get("origin", "")
            local_storage = origin_state.get("localStorage", [])
            if local_storage:
                values = [(item["name"], item["value"]) for item in local_storage]
                context.add_init_script(
                    "if (location.origin === "
                    f"{json.dumps(origin)}) for (const [key, value] of {json.dumps(values)}) "
                    "localStorage.setItem(key, value);",
                )
        for origin, values in state.get("easyFulfillSessionStorage", {}).items():
            if values:
                context.add_init_script(
                    "if (location.origin === "
                    f"{json.dumps(origin)}) for (const [key, value] of {json.dumps(list(values.items()))}) "
                    "sessionStorage.setItem(key, value);",
                )
        return bool(cookies)
    except Exception:
        AUTH_STATE_PATH.unlink(missing_ok=True)
        print("저장된 우체국 포털 로그인 세션을 복원하지 못했습니다. 새 로그인이 필요합니다.", flush=True)
        return False


def launch_epost_context(playwright):
    """현재 Windows 사용자 전용 Chromium 프로필을 연다.

    로그인 세션은 Chromium 프로필과 Windows DPAPI로 보호한 백업에만 보관한다.
    세션이 유효하지 않을 때만 기존 접근 제어된 설정 탭의 로그인 정보를 읽는다.
    """
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    options = {
        "headless": False,
        "args": ["--window-size=1440,900"],
        "no_viewport": True,
    }
    try:
        return playwright.chromium.launch_persistent_context(str(PROFILE_DIR), **options)
    except Exception as error:
        if "Executable doesn't exist" not in str(error):
            raise
        print("우체국 포털용 Chromium을 처음 설치합니다. 잠시 기다려 주세요.", flush=True)
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
        return playwright.chromium.launch_persistent_context(str(PROFILE_DIR), **options)


def connect_login(timeout_seconds: int = LOGIN_TIMEOUT_SECONDS) -> dict[str, object]:
    """전용 창에서 로그인 세션을 연결하고, 세션은 Chromium 프로필에만 남긴다."""
    try:
        from playwright.sync_api import sync_playwright
    except ModuleNotFoundError as error:
        if error.name != "playwright":
            raise
        raise RuntimeError(
            "Playwright가 설치돼 있지 않습니다. easy-fulfill 폴더의 run.bat으로 프로그램을 다시 실행해 주세요.",
        ) from None

    with sync_playwright() as playwright:
        context = launch_epost_context(playwright)
        restore_epost_session(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30_000)
            login_source = ensure_portal_login(page, timeout_seconds)
            save_epost_session(context, page)
            return {"connected": True, "loginSource": login_source}
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--login", action="store_true", help="전용 Chromium 창에서 우체국 로그인 세션을 연결합니다.")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=LOGIN_TIMEOUT_SECONDS)
    args = parser.parse_args()

    if args.self_test:
        assert page_has_logged_in_state.__name__ == "page_has_logged_in_state"
        assert wait_for_logged_in_state.__name__ == "wait_for_logged_in_state"
        assert portal_credentials_from_settings({
            CONFIG_KEY_PORTAL_MEMBER_ID: "member",
            CONFIG_KEY_PORTAL_PASSWORD: "password",
        }).member_id == "member"
        print("self-test: ok")
        return
    if not args.login:
        parser.error("--login을 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")

    result = connect_login(args.timeout_seconds)
    print(json.dumps(result, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
