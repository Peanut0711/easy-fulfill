"""기존 쿠팡 상품의 옵션별 이미지 상세를 기본 HTML 상세로 전환한다.

기본은 화면에 HTML만 채우고 저장하지 않는다. --apply는 사용자가 화면을 확인한 뒤에만 쓴다.
"""

from __future__ import annotations

import argparse
import json
import queue
import re
import sys
import threading
from collections import deque
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parent
DETAIL_URL = "https://wing.coupang.com/tenants/seller-web/vendor-inventory/modify?vendorInventoryId={vendor_inventory_id}"
WING_CLOSED_MARKER = "[WING 창 닫힘 감지]"


def diagnostic_url(url):
    parsed = urlparse(url)
    return parsed._replace(query="", fragment="").geturl()


def diagnostic_text(value):
    value = re.sub(r'https?://[^\s<>"\']+', lambda match: diagnostic_url(match.group()), str(value))
    return re.sub(r'(?i)(authorization|cookie|token|password|session)(["\']?\s*[:=]\s*["\']?)([^"\'\s;&<]+)',
                  r'\1\2[REDACTED]', value)[:4000]


def start_detail_diagnostics(page):
    """상품 수정 화면의 최근 오류만 보관한다. 요청 본문·헤더는 수집하지 않는다."""
    events = deque(maxlen=40)

    def response_error(response):
        if response.status >= 400 and response.request.resource_type in {"document", "xhr", "fetch", "script"}:
            events.append({"kind": "http", "status": response.status, "url": diagnostic_url(response.url)})

    page.on("response", response_error)
    page.on("requestfailed", lambda request: events.append({
        "kind": "requestfailed", "url": diagnostic_url(request.url), "error": diagnostic_text(request.failure),
    }))
    page.on("pageerror", lambda error: events.append({"kind": "pageerror", "error": diagnostic_text(error)}))
    page.on("console", lambda message: events.append({"kind": "console", "error": diagnostic_text(message.text)})
            if message.type == "error" else None)
    return events


def save_detail_diagnostics(page, vendor_inventory_id, stage, events):
    """진단 저장에 실패해도 원래 탭 조작 오류를 가리지 않는다."""
    record = {"time": datetime.now().isoformat(), "vendorInventoryId": vendor_inventory_id,
              "stage": stage, "events": list(events or []), "captureErrors": []}
    try:
        record["url"] = diagnostic_url(page.url)
        record["dom"] = page.evaluate("""() => ({
            title: document.title, readyState: document.readyState,
            text: (document.body?.innerText || '').slice(0, 4000),
            tabs: Array.from(document.querySelectorAll('label[for^="tab-content"], [role="tab"]')).slice(0, 60).map(el => ({
                for: el.getAttribute('for'), text: (el.innerText || '').slice(0, 100),
                visible: !!el.getClientRects().length && getComputedStyle(el).visibility !== 'hidden',
                checked: document.getElementById(el.htmlFor)?.checked ?? null,
                disabled: document.getElementById(el.htmlFor)?.disabled ?? null
            })),
            dialogs: Array.from(document.querySelectorAll('[role="dialog"], .sweet-alert, .modal'))
                .filter(el => el.getClientRects().length).slice(0, 10).map(el => (el.innerText || '').slice(0, 500))
        })""")
        record["dom"]["title"] = diagnostic_text(record["dom"]["title"])
        record["dom"]["text"] = diagnostic_text(record["dom"]["text"])
        record["dom"]["dialogs"] = [diagnostic_text(text) for text in record["dom"]["dialogs"]]
        for tab in record["dom"]["tabs"]:
            tab["text"] = diagnostic_text(tab["text"])
        record["frames"] = [diagnostic_url(frame.url) for frame in page.frames]
    except Exception as error:
        record["captureErrors"].append(diagnostic_text(error))
    print(f"[WING 화면 진단] 단계={stage} · 등록상품 ID={vendor_inventory_id} · URL={record.get('url', '(확인 불가)')}", flush=True)
    try:
        directory = ROOT / "output" / "wing-diagnostics"
        directory.mkdir(parents=True, exist_ok=True)
        prefix = directory / f"detail-{datetime.now():%Y%m%d-%H%M%S-%f}"
        try:
            screenshot = prefix.with_suffix(".png")
            page.screenshot(path=str(screenshot), full_page=True, timeout=5_000)
            record["screenshot"] = str(screenshot)
        except Exception as error:
            record["captureErrors"].append(diagnostic_text(error))
        path = prefix.with_suffix(".json")
        path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[WING 진단 파일] {path}", flush=True)
        return path
    except Exception as error:
        print(f"[WING 진단 저장 실패] {diagnostic_text(error)}", flush=True)
        return None


def load_coupang_modules():
    """실행 Python에 Playwright가 없을 때 작업 로그에 복구 방법을 남긴다."""
    try:
        from playwright.sync_api import Error as playwright_error, sync_playwright
        import coupang_cdn_upload
    except ModuleNotFoundError as error:
        if error.name != "playwright":
            raise
        raise SystemExit(
            "Playwright가 현재 실행 환경에 설치되어 있지 않습니다. "
            "easy-fulfill 폴더의 run.bat으로 프로그램을 다시 실행해 설치를 완료하세요."
        ) from None
    return playwright_error, sync_playwright, coupang_cdn_upload


def load_html(naver_product_no: str):
    path = ROOT / "output" / "detail-preview" / naver_product_no / "coupang-paste.html"
    if not path.exists():
        raise RuntimeError(f"붙여넣기 HTML이 없습니다: {path}. 먼저 1번 변환의 --upload를 실행하세요.")
    return path, path.read_text(encoding="utf-8")


def editor_textarea(page):
    candidates = page.locator("textarea:not([disabled])")
    count = candidates.count()
    for index in range(count):
        candidate = candidates.nth(index)
        if candidate.get_attribute("placeholder") not in {"의견을 적어주세요.", "현재 페이지에 대한 피드백을 적어주세요."}:
            return candidate
    debug_path = ROOT / "output" / "coupang-detail-replace-debug.html"
    debug_path.parent.mkdir(parents=True, exist_ok=True)
    debug_path.write_text(page.content(), encoding="utf-8")
    page.screenshot(path=str(debug_path.with_suffix(".png")), full_page=True)
    raise RuntimeError(f"쿠팡 HTML 입력칸을 찾지 못했습니다. 진단 파일: {debug_path}")


def confirm_detail_level_change(page):
    alert = page.locator(".sweet-alert.showSweetAlert.visible")
    try:
        alert.wait_for(state="visible", timeout=3_000)
    except Exception:
        return
    alert.locator("button.confirm").click(timeout=10_000)


def detail_level_missing_message(vendor_inventory_id: str):
    return (
        f"쿠팡 등록상품 ID {vendor_inventory_id}의 수정 화면에서 상세설명 탭을 찾지 못했습니다. "
        "등록상품 ID가 올바른지, 쿠팡 WING에서 해당 상품을 직접 열 수 있는지 확인하세요."
    )


def select_basic_detail_level(page, vendor_inventory_id: str, playwright_error, events=None):
    try:
        page.locator('label[for="tab-content-level-0"]').click(timeout=15_000)
    except playwright_error as error:
        save_detail_diagnostics(page, vendor_inventory_id, "기본 상세 선택", events)
        message = detail_level_missing_message(vendor_inventory_id)
        print(f"[오류] {message}")
        raise RuntimeError(message) from error


def select_html_detail_tab(page, vendor_inventory_id, playwright_error, events=None):
    try:
        page.locator('label[for="tab-content-2"]').click(timeout=15_000)
    except playwright_error as error:
        path = save_detail_diagnostics(page, vendor_inventory_id, "HTML 작성 탭 선택", events)
        message = (f"쿠팡 등록상품 ID {vendor_inventory_id}의 HTML 작성 탭을 선택하지 못했습니다. "
                   "화면 로딩·팝업·보안 안내·탭 구성 여부를 진단 파일에서 확인하세요. "
                   f"진단 파일: {path if path else '저장 실패 (작업 로그 확인)'}")
        print(f"[오류] {message}", flush=True)
        raise RuntimeError(message) from error


def close_context(context, playwright_error):
    """사용자가 WING 창을 먼저 닫은 경우에도 종료를 정상 처리한다."""
    try:
        context.close()
    except playwright_error as error:
        if "has been closed" not in str(error):
            raise


def watch_completion_popup_close(page):
    """완료 팝업의 닫기를 사용자가 누른 경우에만 종료 요청을 전달한다."""
    requested = threading.Event()
    page.expose_function("__easyFulfillCompletionClosed", requested.set)
    script = r"""
        (() => {
            if (window.__easyFulfillCompletionWatcher) return;
            window.__easyFulfillCompletionWatcher = true;
            const normalize = text => (text || '').replace(/\s+/g, '');
            document.addEventListener('click', event => {
                if (!event.isTrusted || !(event.target instanceof Element)) return;
                const button = event.target.closest('button, a, [role="button"], input[type="button"]');
                if (!button || normalize(button.innerText || button.value) !== '닫기') return;
                // 상위 페이지에 남은 완료 문구와 다른 닫기 버튼을 혼동하지 않는다.
                for (let panel = button.parentElement; panel && panel !== document.body; panel = panel.parentElement) {
                    if (!panel.getClientRects().length) continue;
                    const text = normalize(panel.innerText);
                    if (!/^수정요청이완료되었습니다[.!。]?/.test(text)) continue;
                    const controls = text.replace(/^수정요청이완료되었습니다[.!。]?/, '');
                    if (!['닫기', '닫기상품목록', '상품목록닫기'].includes(controls)) continue;
                    window.__easyFulfillCompletionClosed().catch(() => {});
                    return;
                }
            }, true);
        })();
    """
    page.add_init_script(script)
    page.evaluate(script)
    return requested


def wait_for_wing_close_or_finish(page, context, playwright_error):
    """WING 창 닫힘 또는 GUI의 작업 종료 요청까지 기다린다."""
    finish_requests = queue.Queue()
    wing_closed = threading.Event()

    page.on("close", lambda _page: wing_closed.set())
    context.on("close", lambda _context: wing_closed.set())
    completion_closed = watch_completion_popup_close(page)

    def read_finish_request():
        try:
            finish_requests.put(sys.stdin.readline())
        except OSError:
            # GUI가 종료돼 표준 입력이 닫힌 경우에도 아래 대기 루프를 끝낸다.
            finish_requests.put("")

    threading.Thread(target=read_finish_request, daemon=True).start()
    print("WING 수정요청 완료 팝업에서 닫기를 누르거나 창을 닫으면 작업이 자동으로 종료됩니다.")
    while not wing_closed.is_set():
        if completion_closed.is_set():
            print("[WING 수정요청 완료] 완료 팝업의 닫기를 눌러 WING 창을 종료합니다.")
            close_context(context, playwright_error)
            wing_closed.set()
            break
        try:
            finish_requests.get_nowait()
        except queue.Empty:
            try:
                # Playwright 동기 이벤트 루프가 창 닫힘 이벤트를 수신하도록 한다.
                page.wait_for_timeout(200)
            except playwright_error:
                wing_closed.set()
        else:
            print("WING 작업 종료 요청을 받아 저장 없이 종료합니다.")
            return
    print(f"{WING_CLOSED_MARKER} WING 창이 닫혀 작업을 종료합니다.")


def main():
    parser = argparse.ArgumentParser(description="기존 쿠팡 상품의 상세설명만 HTML로 교체합니다.")
    parser.add_argument("naver_product_no", nargs="?")
    parser.add_argument("vendor_inventory_id", nargs="?")
    parser.add_argument("--apply", action="store_true", help="수정 및 검수 요청 버튼까지 누릅니다.")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        assert DETAIL_URL.format(vendor_inventory_id="123").endswith("vendorInventoryId=123")
        assert "등록상품 ID 123" in detail_level_missing_message("123")
        assert WING_CLOSED_MARKER.startswith("[WING")
        print("self-test: ok")
        return
    if not args.naver_product_no or not args.vendor_inventory_id or not args.naver_product_no.isdigit() or not args.vendor_inventory_id.isdigit():
        parser.error("네이버 상품번호와 vendorInventoryId는 숫자여야 합니다.")

    playwright_error, sync_playwright, coupang_cdn_upload = load_coupang_modules()
    html_path, html = load_html(args.naver_product_no)
    with sync_playwright() as playwright:
        context = coupang_cdn_upload.launch_coupang_context(playwright)
        coupang_cdn_upload.restore_coupang_session(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            coupang_cdn_upload.wait_for_login(page)
            coupang_cdn_upload.save_coupang_session(context, page)
            events = start_detail_diagnostics(page)
            page.goto(DETAIL_URL.format(vendor_inventory_id=args.vendor_inventory_id))
            select_basic_detail_level(page, args.vendor_inventory_id, playwright_error, events)
            confirm_detail_level_change(page)
            page.wait_for_timeout(1_000)
            select_html_detail_tab(page, args.vendor_inventory_id, playwright_error, events)
            confirm_detail_level_change(page)
            page.wait_for_timeout(1_000)
            textarea = editor_textarea(page)
            textarea.fill(html, timeout=30_000)
            print(f"준비 완료: 쿠팡 상품 {args.vendor_inventory_id}에 {html_path.name}을 채웠습니다.")
            if not args.apply:
                wait_for_wing_close_or_finish(page, context, playwright_error)
                return
            if input("'APPLY'를 입력하면 수정 및 검수 요청을 보냅니다: ") != "APPLY":
                print("저장하지 않았습니다.")
                return
            page.get_by_role("button", name="수정 및 검수 요청").click(timeout=15_000)
            print("수정 및 검수 요청을 보냈습니다.")
        finally:
            close_context(context, playwright_error)


if __name__ == "__main__":
    main()
