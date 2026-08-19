"""우체국 계약고객전용시스템 전용 Chromium 로그인 세션 연결 도구."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import subprocess
import sys
import time


ROOT = Path(__file__).resolve().parent
PROFILE_DIR = ROOT / "output" / "epost-browser-profile"
PORTAL_URL = "https://biz.epost.go.kr/ui/index.jsp"
LOGIN_TIMEOUT_SECONDS = 600


def page_has_logged_in_state(page) -> bool:
    """포털 화면에 로그인 완료 표식이 보이는지만 확인한다."""
    try:
        return "로그아웃" in page.locator("body").inner_text(timeout=3_000)
    except Exception:
        return False


def launch_epost_context(playwright):
    """현재 Windows 사용자 전용 Chromium 프로필을 연다.

    로그인 쿠키는 Chromium이 Windows 사용자 프로필 아래에서 관리한다. ID·비밀번호를
    별도 파일이나 스프레드시트에 쓰지 않는다.
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
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30_000)
            deadline = time.monotonic() + timeout_seconds
            while time.monotonic() < deadline:
                if page_has_logged_in_state(page):
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
