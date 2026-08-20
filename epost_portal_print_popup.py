"""검증된 우체국 신규출력 행만 선택해 운송장출력 팝업 직전까지 연다.

이 모듈은 포털 행 선택과 운송장출력 팝업 열기까지만 수행한다. 팝업의 인쇄,
OZ Viewer 실행, Windows 인쇄 명령은 수행하지 않는다.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import time

from epost_portal_lookup import (
    PAGE_READY_TIMEOUT_SECONDS,
    apply_new_print_query,
    korea_today,
    pending_candidates_for_date,
    read_portal_grid_rows,
    verify_target_rows,
    wait_for_print_page,
    wait_for_query_result,
)
from epost_portal_session import (
    LOGIN_TIMEOUT_SECONDS,
    PORTAL_URL,
    ensure_portal_login,
    launch_epost_context,
    restore_epost_session,
)
from post_parcel_receipt_store import ParcelReceiptStore, ReceiptStoreError


POPUP_TIMEOUT_SECONDS = 15
PRINT_BUTTON_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-portal-print-button-controls.json"
)


def preferred_outer_control_id(control_ids: list[str]) -> str:
    """버튼 본체와 그 내부 텍스트 요소가 함께 잡힐 때 가장 바깥 본체만 고른다."""
    unique_ids = list(dict.fromkeys(control_ids))
    if len(unique_ids) == 1:
        return unique_ids[0]
    outer = [
        control_id for control_id in unique_ids
        if all(other_id == control_id or other_id.startswith(control_id) for other_id in unique_ids)
    ]
    return outer[0] if len(outer) == 1 else ""


def _visible_work_area_button_controls(page, work_prefix: str, text: str) -> list[dict[str, object]]:
    """현재 운송장출력 탭의 텍스트 일치 버튼 후보만 안전한 메타데이터로 찾는다."""
    return page.locator("[id]").evaluate_all(
        """
        (elements, options) => elements.filter(element => {
            const style = getComputedStyle(element);
            return element.id.startsWith(options.prefix)
                && /btn/i.test(element.id)
                && (element.innerText || '').trim() === options.text
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
        {"prefix": work_prefix, "text": text},
    )


def _save_print_button_diagnostic(work_prefix: str, controls: list[dict[str, object]]) -> None:
    PRINT_BUTTON_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    PRINT_BUTTON_DIAGNOSTIC_PATH.write_text(
        json.dumps({"workPrefix": work_prefix, "printButtonControls": controls}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def select_verified_target_rows(page, work_prefix: str, target_rows: list[dict[str, object]]) -> int:
    """검증된 각 행의 체크박스만 한 번 누른다."""
    if not target_rows:
        raise RuntimeError("선택할 검증 대상 행이 없습니다.")
    selected = 0
    for row in target_rows:
        row_index = int(row["rowIndex"])
        checkbox_id = (
            f"{work_prefix}grdList_body_gridrow_{row_index}"
            f"_cell_{row_index}_0_controlcheckbox"
        )
        checkbox = page.locator(f'[id="{checkbox_id}"]')
        if checkbox.count() != 1 or not checkbox.is_visible():
            raise RuntimeError("검증한 대상 행의 체크박스를 찾지 못했습니다.")
        checkbox.click()
        selected += 1
    return selected


def open_print_popup(page, work_prefix: str) -> None:
    """현재 탭의 운송장출력 버튼만 눌러 팝업을 연 뒤 인쇄 전에서 멈춘다."""
    controls = _visible_work_area_button_controls(page, work_prefix, "운송장출력")
    button_id = preferred_outer_control_id([str(item["id"]) for item in controls])
    if not button_id:
        _save_print_button_diagnostic(work_prefix, controls)
        raise RuntimeError("운송장출력 버튼을 하나로 식별하지 못했습니다.")
    page.locator(f'[id="{button_id}"]').click()


def print_popup_is_open(page) -> bool:
    """팝업의 제목과 인쇄 버튼이 동시에 보이는지만 판정한다."""
    try:
        body_text = page.locator("body").inner_text(timeout=3_000)
        print_controls = page.get_by_text("인쇄", exact=True)
        return "운송장출력(팝업)" in body_text and any(
            print_controls.nth(index).is_visible() for index in range(print_controls.count())
        )
    except Exception:
        return False


def wait_for_popup(page) -> bool:
    """인쇄 버튼을 누르지 않고 운송장출력 팝업의 표시만 확인한다."""
    deadline = time.monotonic() + POPUP_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        if print_popup_is_open(page):
            return True
        time.sleep(0.5)
    return False


def wait_for_popup_close(page, timeout_seconds: int) -> bool:
    """사용자가 팝업을 검토한 뒤 X로 닫을 때까지 인쇄하지 않고 대기한다."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if not print_popup_is_open(page):
            return True
        time.sleep(0.5)
    return False


def prepare_verified_print_popup(
    page, lookup_date: str, candidates,
) -> dict[str, object]:
    """신규출력의 검증된 행만 선택하고 운송장출력 팝업을 연다.

    팝업의 인쇄 버튼은 이 함수에서 누르지 않는다. OZ Viewer 단계도 같은 엄격한
    대상 검증을 재사용하도록, 브라우저·접수 이력 수집과 분리한다.
    """
    expected_regi_nos = [candidate.regi_no for candidate in candidates]
    work_prefix = apply_new_print_query(page, lookup_date)
    body_text, _ = wait_for_query_result(page, candidates)
    if any(regi_no not in body_text for regi_no in expected_regi_nos):
        raise RuntimeError("포털 조회 본문에서 프로그램 접수 등기번호를 모두 찾지 못했습니다.")
    verification = verify_target_rows(expected_regi_nos, read_portal_grid_rows(page, work_prefix))
    if not verification["verified"]:
        raise RuntimeError("포털 신규출력 대상 행이 안전 조건과 일치하지 않습니다.")
    selected_count = select_verified_target_rows(page, work_prefix, verification["targetRows"])
    open_print_popup(page, work_prefix)
    if not wait_for_popup(page):
        raise RuntimeError("운송장출력 팝업이 열렸는지 확인하지 못했습니다.")
    return {
        "selectedCount": selected_count,
        "selectedRegiNos": expected_regi_nos,
    }


def open_verified_print_popup(timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS) -> dict[str, object]:
    """오늘 신규출력에서 안전하게 검증된 행만 선택해 팝업을 연다."""
    lookup_date = korea_today()
    try:
        candidates = pending_candidates_for_date(
            ParcelReceiptStore().list_pending_prints(), lookup_date,
        )
    except ReceiptStoreError as error:
        raise RuntimeError("프로그램의 우체국 실제 접수 이력을 읽지 못했습니다.") from error
    if not candidates:
        raise RuntimeError(f"{lookup_date}에 프로그램이 실제 접수한 미출력 건이 없습니다.")

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
            if not wait_for_popup_close(page, timeout_seconds):
                raise RuntimeError("운송장출력 팝업 검토 시간이 만료됐습니다. 인쇄하지 말고 팝업을 닫은 뒤 다시 실행해 주세요.")
            return {
                "popupOpened": True,
                "lookupDate": lookup_date,
                **prepared,
                "popupReviewed": True,
                "printExecuted": False,
            }
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--open-popup", action="store_true", help="검증된 행만 선택해 운송장출력 팝업을 엽니다.")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=PAGE_READY_TIMEOUT_SECONDS)
    args = parser.parse_args()

    if args.self_test:
        assert preferred_outer_control_id(["btn", "btnTextBoxElement"]) == "btn"
        print("self-test: ok")
        return
    if not args.open_popup:
        parser.error("--open-popup을 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")
    print(json.dumps(open_verified_print_popup(args.timeout_seconds), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
