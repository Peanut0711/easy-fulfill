"""우체국 팝업의 인쇄 요청 후 OZ Report Viewer 열림만 확인한다.

이 단계는 포털의 빨간 `인쇄` 버튼을 누르므로 신규출력 상태가 바뀔 수 있다.
OZ Viewer의 프린터 아이콘, Windows 인쇄 대화상자, 실제 인쇄 명령은 수행하지 않는다.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from epost_desktop_windows import (
    oz_viewer_windows,
    wait_for_new_single_oz_viewer,
    wait_for_window_close,
)
from epost_portal_lookup import PAGE_READY_TIMEOUT_SECONDS, korea_today, pending_candidates_for_date, wait_for_print_page
from epost_portal_print_popup import preferred_outer_control_id, prepare_verified_print_popup
from epost_portal_reprint_popup import prepare_verified_reprint_popup
from epost_portal_session import (
    LOGIN_TIMEOUT_SECONDS,
    PORTAL_URL,
    ensure_portal_login,
    launch_epost_context,
    restore_epost_session,
)
from post_parcel_receipt_store import ParcelReceiptStore, ReceiptStoreError


OZ_VIEWER_TIMEOUT_SECONDS = 20
POPUP_PRINT_BUTTON_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-portal-popup-print-button-controls.json"
)


def _visible_print_button_controls(page) -> list[dict[str, object]]:
    """현재 화면에서 보이는 텍스트 일치 인쇄 버튼 후보의 메타데이터만 읽는다."""
    return page.locator("[id]").evaluate_all(
        """
        elements => elements.filter(element => {
            const style = getComputedStyle(element);
            return /btn/i.test(element.id)
                && (element.innerText || '').trim() === '인쇄'
                && style.display !== 'none'
                && style.visibility !== 'hidden'
                && element.getClientRects().length > 0;
        }).map(element => {
            const rect = element.getBoundingClientRect();
            return {
                id: element.id,
                tag: element.tagName.toLowerCase(),
                left: Math.round(rect.left),
                top: Math.round(rect.top),
                width: Math.round(rect.width),
                height: Math.round(rect.height)
            };
        })
        """,
    )


def _save_popup_print_button_diagnostic(controls: list[dict[str, object]]) -> None:
    POPUP_PRINT_BUTTON_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    POPUP_PRINT_BUTTON_DIAGNOSTIC_PATH.write_text(
        json.dumps({"popupPrintButtonControls": controls}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def preferred_popup_print_button_id(control_ids: list[str]) -> str:
    """현재 운송장출력 팝업 상단의 실제 인쇄 버튼만 고른다.

    Nexacro는 팝업 본체 외에 뒤쪽 화면의 인쇄 버튼과 팝업 하단의 동명 버튼도
    동시에 DOM에 남긴다. 실제 상단 빨간 버튼은 현재 포털에서 `form_btnPrint2`
    식별자로 렌더링되므로, 이 규칙에 하나만 맞을 때만 진행한다.
    """
    matches = [control_id for control_id in control_ids if control_id.endswith("_form_btnPrint2")]
    return matches[0] if len(matches) == 1 else ""


def click_popup_print_button(page) -> None:
    """운송장출력 팝업에서 유일하게 식별된 빨간 인쇄 버튼만 누른다."""
    controls = _visible_print_button_controls(page)
    button_id = preferred_popup_print_button_id([str(control["id"]) for control in controls])
    if not button_id:
        _save_popup_print_button_diagnostic(controls)
        raise RuntimeError("운송장출력 팝업의 인쇄 버튼을 하나로 식별하지 못했습니다.")
    page.locator(f'[id="{button_id}"]').click()


def open_oz_viewer(timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS) -> dict[str, object]:
    """검증된 행의 포털 인쇄 요청 후 OZ Viewer 하나가 열렸는지 확인한다."""
    lookup_date = korea_today()
    try:
        candidates = pending_candidates_for_date(
            ParcelReceiptStore().list_pending_prints(), lookup_date,
        )
    except ReceiptStoreError as error:
        raise RuntimeError("프로그램의 우체국 실제 접수 이력을 읽지 못했습니다.") from error
    if not candidates:
        raise RuntimeError(f"{lookup_date}에 프로그램이 실제 접수한 미출력 건이 없습니다.")

    existing_viewers = oz_viewer_windows()
    if existing_viewers:
        raise RuntimeError("기존 OZ Report Viewer 창이 열려 있습니다. 인쇄하지 않고 해당 창을 먼저 닫아 주세요.")

    expected = [candidate.regi_no for candidate in candidates]
    store = ParcelReceiptStore()
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = launch_epost_context(playwright)
        restore_epost_session(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30_000)
            ensure_portal_login(page, LOGIN_TIMEOUT_SECONDS)
            wait_for_print_page(page, timeout_seconds)
            prepared = prepare_verified_print_popup(page, lookup_date, candidates)
            click_popup_print_button(page)
            store.mark_portal_print_requested(expected)
            viewer = wait_for_new_single_oz_viewer(set(), OZ_VIEWER_TIMEOUT_SECONDS)
            if not wait_for_window_close(viewer.handle, timeout_seconds):
                raise RuntimeError("OZ Report Viewer 검토 시간이 만료됐습니다. 프린터 아이콘을 누르지 말고 창을 닫은 뒤 다시 확인해 주세요.")
            return {
                "ozViewerOpened": True,
                "lookupDate": lookup_date,
                **prepared,
                "ozViewerReviewed": True,
                "printCommandExecuted": False,
            }
        finally:
            context.close()


def open_reprint_oz_viewer(timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS) -> dict[str, object]:
    """포털에서 이미 출력으로 확인한 단건만 재출력 인쇄 요청 후 OZ Viewer를 확인한다."""
    store = ParcelReceiptStore()
    try:
        candidates = store.list_portal_print_confirmed()
    except ReceiptStoreError as error:
        raise RuntimeError("프로그램의 포털 출력 확인 이력을 읽지 못했습니다.") from error
    if not candidates:
        raise RuntimeError("재출력 OZ Viewer를 열 포털 출력 확인 건이 없습니다.")
    if len(candidates) != 1:
        raise RuntimeError("단건 재출력은 포털 출력 확인 건이 정확히 1건일 때만 실행할 수 있습니다.")

    existing_viewers = oz_viewer_windows()
    if existing_viewers:
        raise RuntimeError("기존 OZ Report Viewer 창이 열려 있습니다. 인쇄하지 않고 해당 창을 먼저 닫아 주세요.")

    from playwright.sync_api import sync_playwright
    with sync_playwright() as playwright:
        context = launch_epost_context(playwright)
        restore_epost_session(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30_000)
            ensure_portal_login(page, LOGIN_TIMEOUT_SECONDS)
            wait_for_print_page(page, timeout_seconds)
            prepared = prepare_verified_reprint_popup(page, candidates)
            click_popup_print_button(page)
            viewer = wait_for_new_single_oz_viewer(set(), OZ_VIEWER_TIMEOUT_SECONDS)
            if not wait_for_window_close(viewer.handle, timeout_seconds):
                raise RuntimeError("OZ Report Viewer 검토 시간이 만료됐습니다. 프린터 아이콘을 누르지 말고 창을 닫은 뒤 다시 확인해 주세요.")
            return {
                "reprintOzViewerOpened": True,
                **prepared,
                "ozViewerReviewed": True,
                "printCommandExecuted": False,
            }
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--open-oz-viewer", action="store_true", help="포털 신규출력 인쇄 요청 후 OZ Viewer 열림만 확인합니다.")
    parser.add_argument("--open-reprint-oz-viewer", action="store_true", help="포털 출력 확인 단건을 재출력해 OZ Viewer 열림만 확인합니다.")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=PAGE_READY_TIMEOUT_SECONDS)
    args = parser.parse_args()

    if args.self_test:
        assert preferred_outer_control_id(["btn", "btnTextBoxElement"]) == "btn"
        print("self-test: ok")
        return
    if not (args.open_oz_viewer or args.open_reprint_oz_viewer):
        parser.error("--open-oz-viewer 또는 --open-reprint-oz-viewer를 지정하세요.")
    if args.open_oz_viewer and args.open_reprint_oz_viewer:
        parser.error("OZ Viewer 열기 방식은 하나만 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")
    opener = open_reprint_oz_viewer if args.open_reprint_oz_viewer else open_oz_viewer
    print(json.dumps(opener(args.timeout_seconds), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
