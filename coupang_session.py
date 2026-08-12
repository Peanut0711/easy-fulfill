"""Easy Fulfill용 쿠팡 WING 로그인 세션을 연결하거나 삭제한다."""

from __future__ import annotations

import argparse
import json
import shutil

from playwright.sync_api import sync_playwright

import coupang_cdn_upload


def seller_name(page):
    names = [name.strip() for name in page.locator("em").all_text_contents() if name.strip()]
    return names[0] if names else ""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--login", action="store_true", help="WING 로그인 창을 열고 세션을 저장합니다.")
    parser.add_argument("--clear", action="store_true", help="이 PC의 WING 로그인 세션을 삭제합니다.")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        assert seller_name.__name__ == "seller_name"
        print("self-test: ok")
        return
    if args.login == args.clear:
        parser.error("--login 또는 --clear 중 하나를 지정하세요.")
    if args.clear:
        if coupang_cdn_upload.PROFILE_DIR.exists():
            shutil.rmtree(coupang_cdn_upload.PROFILE_DIR)
        print(json.dumps({"connected": False, "cleared": True}, ensure_ascii=False))
        return

    with sync_playwright() as playwright:
        context = coupang_cdn_upload.launch_coupang_context(playwright)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            coupang_cdn_upload.wait_for_login(page)
            coupang_cdn_upload.save_coupang_session(context, page)
            print(json.dumps({"connected": True, "seller": seller_name(page)}, ensure_ascii=False))
        finally:
            context.close()


if __name__ == "__main__":
    main()
