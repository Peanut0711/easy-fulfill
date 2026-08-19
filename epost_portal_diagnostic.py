"""우체국 운송장출력 화면의 자동화 가능 여부를 읽기 전용으로 진단한다."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import time
from urllib.parse import urlparse

from epost_portal_session import (
    LOGIN_TIMEOUT_SECONDS,
    PORTAL_URL,
    launch_epost_context,
    page_has_logged_in_state,
)


DIAGNOSTIC_PATH = Path(__file__).resolve().parent / "output" / "epost-portal-diagnostic.json"


def control_snapshot(page) -> list[dict[str, str]]:
    """개인정보가 담길 수 있는 표·입력값은 제외하고 조작 컨트롤 정보만 읽는다."""
    return page.locator("button, input, select, textarea, [role='button']").evaluate_all(
        """
        controls => controls.map((element, index) => ({
            index: String(index),
            tag: element.tagName.toLowerCase(),
            type: element.getAttribute('type') || '',
            id: element.id || '',
            name: element.getAttribute('name') || '',
            role: element.getAttribute('role') || '',
            ariaLabel: element.getAttribute('aria-label') || '',
            title: element.getAttribute('title') || '',
            placeholder: element.getAttribute('placeholder') || '',
            text: (element.tagName === 'BUTTON' || element.getAttribute('role') === 'button')
                ? (element.innerText || '').trim().slice(0, 80)
                : ''
        }))
        """,
    )


def looks_like_print_page(page) -> bool:
    """운송장출력 화면의 제목·조회 버튼·입력 컨트롤이 함께 보이는지 확인한다."""
    try:
        body_text = page.locator("body").inner_text(timeout=3_000)
        snapshot = control_snapshot(page)
    except Exception:
        return False
    has_query = any(item["text"] == "조회" for item in snapshot)
    input_count = sum(item["tag"] == "input" for item in snapshot)
    return "운송장출력" in body_text and "검색일자" in body_text and has_query and input_count >= 4


def sanitized_page_path(url: str) -> str:
    """쿼리 문자열을 남기지 않는 페이지 경로만 기록한다."""
    parsed = urlparse(url)
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}" if parsed.scheme and parsed.netloc else ""


def diagnose(timeout_seconds: int = LOGIN_TIMEOUT_SECONDS) -> dict[str, object]:
    """전용 Chromium에서 사용자가 이동한 운송장출력 화면을 읽기 전용으로 수집한다."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = launch_epost_context(playwright)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30_000)
            if not page_has_logged_in_state(page):
                raise RuntimeError("저장된 우체국 포털 로그인 세션이 없습니다. 먼저 로그인 연결을 실행해 주세요.")

            deadline = time.monotonic() + timeout_seconds
            while time.monotonic() < deadline:
                if looks_like_print_page(page):
                    snapshot = control_snapshot(page)
                    payload = {
                        "diagnosedAt": time.strftime("%Y-%m-%dT%H:%M:%S"),
                        "pagePath": sanitized_page_path(page.url),
                        "controlCount": len(snapshot),
                        "controls": snapshot,
                    }
                    DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
                    DIAGNOSTIC_PATH.write_text(
                        json.dumps(payload, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                    return {"diagnosed": True, "controlCount": len(snapshot)}
                time.sleep(1)
            raise RuntimeError(
                "운송장출력 화면을 확인하지 못했습니다. 전용 Chromium 창에서 계약소포 > 운송장출력으로 이동해 주세요.",
            )
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--diagnose", action="store_true", help="운송장출력 화면의 컨트롤을 읽기 전용으로 진단합니다.")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=LOGIN_TIMEOUT_SECONDS)
    args = parser.parse_args()

    if args.self_test:
        assert sanitized_page_path("https://example.test/path?a=b") == "https://example.test/path"
        print("self-test: ok")
        return
    if not args.diagnose:
        parser.error("--diagnose를 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")

    print(json.dumps(diagnose(args.timeout_seconds), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
