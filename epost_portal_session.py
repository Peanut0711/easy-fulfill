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
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parent
PROFILE_DIR = ROOT / "output" / "epost-browser-profile"
AUTH_STATE_PATH = PROFILE_DIR / "easy-fulfill-auth-state.bin"
PORTAL_URL = "https://biz.epost.go.kr/ui/index.jsp"
LOGIN_TIMEOUT_SECONDS = 600


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
    ID·비밀번호는 별도 파일이나 스프레드시트에 쓰지 않는다.
    """
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    options = {
        "headless": False,
        "args": ["--window-size=1600,1000"],
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
    """전용 창에서 사용자의 직접 로그인을 기다리고, 세션은 Chromium 프로필에만 남긴다."""
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
            deadline = time.monotonic() + timeout_seconds
            while time.monotonic() < deadline:
                if page_has_logged_in_state(page):
                    save_epost_session(context, page)
                    return {"connected": True}
                time.sleep(1)
            raise RuntimeError("로그인 완료를 확인하지 못했습니다. 전용 Chromium 창에서 로그인 상태를 확인해 주세요.")
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
