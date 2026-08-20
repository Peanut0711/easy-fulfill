"""포털 출력 확인 건만 정확히 선택해 재출력 팝업을 연다.

이 단계는 팝업을 열기만 하며 팝업의 인쇄 버튼, OZ Viewer, Windows 인쇄 명령은
수행하지 않는다.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import time

from epost_portal_lookup import (
    PAGE_READY_TIMEOUT_SECONDS,
    apply_print_target_query,
    normalize_registration_number,
    read_portal_grid_rows,
    wait_for_print_page,
    wait_for_query_result,
)
from epost_portal_print_popup import (
    open_print_popup,
    select_verified_target_rows,
    wait_for_popup,
    wait_for_popup_close,
)
from epost_portal_session import (
    LOGIN_TIMEOUT_SECONDS,
    PORTAL_URL,
    ensure_portal_login,
    launch_epost_context,
    restore_epost_session,
)
from post_parcel_receipt_store import ParcelReceiptStore, ReceiptStoreError


REPRINT_GRID_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-portal-reprint-grid-diagnostic.json"
)
MAX_REPRINT_GRID_SCROLL_STEPS = 40
REPRINT_GRID_SCROLL_CLICKS_PER_STEP = 8
REPRINT_GRID_SCROLL_SETTLE_SECONDS = 0.04


def target_registration_control_ids(page, work_prefix: str, regi_no: str) -> list[str]:
    """이미 알고 있는 대상 등기번호가 든 그리드 컨트롤 ID만 찾는다.

    수취인·주소·전화번호 등 다른 셀 값은 읽거나 저장하지 않는다.
    """
    normalized = normalize_registration_number(regi_no)
    return page.locator("[id]").evaluate_all(
        r"""
        (elements, options) => elements.filter(element => {
            if (!element.id.startsWith(options.prefix)) return false;
            const digits = (element.innerText || '').replace(/\D/g, '');
            return digits === options.regiNo;
        }).map(element => element.id)
        """,
        {"prefix": work_prefix, "regiNo": normalized},
    )


def save_reprint_grid_diagnostic(page, work_prefix: str, expected_regi_nos: list[str], portal_rows: list[dict[str, object]]) -> None:
    """재출력 대상의 발견 위치·행 수만 저장한다."""
    target_ids = []
    for regi_no in expected_regi_nos:
        target_ids.extend(target_registration_control_ids(page, work_prefix, regi_no))
    REPRINT_GRID_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPRINT_GRID_DIAGNOSTIC_PATH.write_text(
        json.dumps({
            "gridRowCount": len(portal_rows),
            "targetRegistrationControlIds": target_ids,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def verified_reprint_rows(expected_regi_nos: list[str], portal_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """출력 상태로 정확히 한 번씩 존재하고 체크 가능한 대상 행만 반환한다."""
    target_rows = []
    for regi_no in expected_regi_nos:
        normalized = normalize_registration_number(regi_no)
        matches = [
            row for row in portal_rows
            if normalize_registration_number(str(row.get("regiNo", ""))) == normalized
        ]
        if len(matches) != 1:
            raise RuntimeError("재출력 대상 등기번호를 포털 전체 목록에서 하나로 찾지 못했습니다.")
        row = matches[0]
        if str(row.get("printState", "")).strip() != "출력":
            raise RuntimeError("재출력 대상의 포털 출력여부가 출력이 아닙니다.")
        if not row.get("hasCheckbox"):
            raise RuntimeError("재출력 대상 행의 체크박스를 찾지 못했습니다.")
        target_rows.append(row)
    return target_rows


def _matching_target_rows(expected_regi_nos: list[str], portal_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """현재 보이는 행에서 대상 등기번호와 같은 행만 반환한다."""
    expected = {normalize_registration_number(regi_no) for regi_no in expected_regi_nos}
    return [
        row for row in portal_rows
        if normalize_registration_number(str(row.get("regiNo", ""))) in expected
    ]


def reprint_grid_scroll_down_button_id(work_prefix: str) -> str:
    """Nexacro 재출력 그리드의 아래쪽 세로 스크롤 화살표 ID를 만든다."""
    return f"{work_prefix}grdList_vscrollbar_incbutton"


def find_reprint_rows_by_scrolling(page, work_prefix: str, expected_regi_nos: list[str]) -> list[dict[str, object]]:
    """가상 그리드를 안전하게 스크롤해 대상 행이 보일 때만 반환한다."""
    grid = page.locator(f'[id="{work_prefix}grdList"]')
    if grid.count() != 1 or not grid.is_visible():
        raise RuntimeError("재출력 결과 그리드를 찾지 못했습니다.")
    for _ in range(MAX_REPRINT_GRID_SCROLL_STEPS):
        portal_rows = read_portal_grid_rows(page, work_prefix)
        matches = _matching_target_rows(expected_regi_nos, portal_rows)
        if matches:
            return matches
        scroll_down = page.locator(f'[id="{reprint_grid_scroll_down_button_id(work_prefix)}"]')
        if scroll_down.count() == 1 and scroll_down.is_visible():
            scroll_down.click(
                click_count=REPRINT_GRID_SCROLL_CLICKS_PER_STEP,
                delay=10,
            )
        else:
            grid.hover()
            page.mouse.wheel(0, 1800)
        time.sleep(REPRINT_GRID_SCROLL_SETTLE_SECONDS)
    raise RuntimeError("재출력 대상 등기번호를 포털 전체 목록에서 하나로 찾지 못했습니다.")


def prepare_verified_reprint_popup(page, candidates) -> dict[str, object]:
    """포털 출력 확인 단건만 검증·선택해 재출력 팝업을 열고 인쇄 전에서 멈춘다."""
    if len(candidates) != 1:
        raise RuntimeError("단건 재출력은 포털 출력 확인 건이 정확히 1건일 때만 실행할 수 있습니다.")
    expected = [candidates[0].regi_no]
    lookup_date = candidates[0].received_at[:10]
    work_prefix, total_before_query = apply_print_target_query(
        page, lookup_date, "전체", expected[0],
    )
    wait_for_query_result(page, candidates, total_before_query)
    portal_rows = read_portal_grid_rows(page, work_prefix)
    save_reprint_grid_diagnostic(page, work_prefix, expected, portal_rows)
    visible_target_rows = _matching_target_rows(expected, portal_rows)
    target_rows = verified_reprint_rows(
        expected,
        visible_target_rows or find_reprint_rows_by_scrolling(page, work_prefix, expected),
    )
    selected_count = select_verified_target_rows(page, work_prefix, target_rows)
    open_print_popup(page, work_prefix)
    if not wait_for_popup(page):
        raise RuntimeError("재출력 운송장출력 팝업이 열렸는지 확인하지 못했습니다.")
    return {
        "lookupDate": lookup_date,
        "selectedCount": selected_count,
        "selectedRegiNos": expected,
    }


def open_verified_reprint_popup(timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS) -> dict[str, object]:
    """현재 날짜의 포털 출력 확인 건만 선택해 재출력 팝업을 열고 닫힘을 기다린다."""
    store = ParcelReceiptStore()
    try:
        candidates = store.list_portal_print_confirmed()
    except ReceiptStoreError as error:
        raise RuntimeError("프로그램의 포털 출력 확인 이력을 읽지 못했습니다.") from error
    if not candidates:
        raise RuntimeError("재출력 팝업을 열 포털 출력 확인 건이 없습니다.")
    if len(candidates) != 1:
        raise RuntimeError("단건 재출력은 포털 출력 확인 건이 정확히 1건일 때만 실행할 수 있습니다.")

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
            if not wait_for_popup_close(page, timeout_seconds):
                raise RuntimeError("재출력 팝업 검토 시간이 만료됐습니다. 인쇄하지 말고 팝업을 닫은 뒤 다시 실행해 주세요.")
            return {
                "reprintPopupOpened": True,
                **prepared,
                "popupReviewed": True,
                "printExecuted": False,
            }
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--open-popup", action="store_true", help="확정된 단건만 재출력 팝업을 엽니다.")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=PAGE_READY_TIMEOUT_SECONDS)
    args = parser.parse_args()

    if args.self_test:
        assert len(verified_reprint_rows(["R1"], [{
            "regiNo": "R1", "printState": "출력", "hasCheckbox": True,
        }])) == 1
        print("self-test: ok")
        return
    if not args.open_popup:
        parser.error("--open-popup을 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")
    print(json.dumps(open_verified_reprint_popup(args.timeout_seconds), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
