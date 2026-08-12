"""기존 쿠팡 상품의 옵션별 이미지 상세를 기본 HTML 상세로 전환한다.

기본은 화면에 HTML만 채우고 저장하지 않는다. --apply는 사용자가 화면을 확인한 뒤에만 쓴다.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from playwright.sync_api import sync_playwright

import coupang_cdn_upload


ROOT = Path(__file__).resolve().parent
DETAIL_URL = "https://wing.coupang.com/tenants/seller-web/vendor-inventory/modify?vendorInventoryId={vendor_inventory_id}"


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


def main():
    parser = argparse.ArgumentParser(description="기존 쿠팡 상품의 상세설명만 HTML로 교체합니다.")
    parser.add_argument("naver_product_no", nargs="?")
    parser.add_argument("vendor_inventory_id", nargs="?")
    parser.add_argument("--apply", action="store_true", help="수정 및 검수 요청 버튼까지 누릅니다.")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        assert DETAIL_URL.format(vendor_inventory_id="123").endswith("vendorInventoryId=123")
        print("self-test: ok")
        return
    if not args.naver_product_no or not args.vendor_inventory_id or not args.naver_product_no.isdigit() or not args.vendor_inventory_id.isdigit():
        parser.error("네이버 상품번호와 vendorInventoryId는 숫자여야 합니다.")

    html_path, html = load_html(args.naver_product_no)
    with sync_playwright() as playwright:
        context = coupang_cdn_upload.launch_coupang_context(playwright)
        coupang_cdn_upload.restore_coupang_session(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            coupang_cdn_upload.wait_for_login(page)
            coupang_cdn_upload.save_coupang_session(context, page)
            page.goto(DETAIL_URL.format(vendor_inventory_id=args.vendor_inventory_id))
            page.locator('label[for="tab-content-level-0"]').click(timeout=15_000)
            confirm_detail_level_change(page)
            page.wait_for_timeout(1_000)
            page.locator('label[for="tab-content-2"]').click(timeout=15_000)
            confirm_detail_level_change(page)
            page.wait_for_timeout(1_000)
            textarea = editor_textarea(page)
            textarea.fill(html, timeout=30_000)
            print(f"준비 완료: 쿠팡 상품 {args.vendor_inventory_id}에 {html_path.name}을 채웠습니다.")
            if not args.apply:
                input("화면을 확인한 뒤 Enter를 누르면 저장 없이 종료합니다. ")
                return
            if input("'APPLY'를 입력하면 수정 및 검수 요청을 보냅니다: ") != "APPLY":
                print("저장하지 않았습니다.")
                return
            page.get_by_role("button", name="수정 및 검수 요청").click(timeout=15_000)
            print("수정 및 검수 요청을 보냈습니다.")
        finally:
            context.close()


if __name__ == "__main__":
    main()
