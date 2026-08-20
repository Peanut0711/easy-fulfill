"""우체국 포털의 출력여부를 읽기 전용으로 재확인한다."""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from epost_portal_lookup import (
    PAGE_READY_TIMEOUT_SECONDS,
    apply_print_target_query,
    korea_today,
    read_portal_grid_rows,
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


OUTPUT_CONFIRM_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-portal-output-confirm-diagnostic.json"
)


def confirmed_output_regi_nos(expected_regi_nos: list[str], portal_rows: list[dict[str, object]]) -> list[str]:
    """포털 그리드에서 정확히 한 번씩 `출력`으로 확인된 등기번호만 반환한다."""
    confirmed: list[str] = []
    for regi_no in expected_regi_nos:
        matches = [row for row in portal_rows if row.get("regiNo") == regi_no]
        if len(matches) == 1 and str(matches[0].get("printState", "")).strip() == "출력":
            confirmed.append(regi_no)
    return confirmed


def save_output_confirm_diagnostic(expected_regi_nos: list[str], portal_rows: list[dict[str, object]]) -> None:
    """대상 행의 개수·상태만 기록하고 등기번호나 수취인 정보는 저장하지 않는다."""
    target_states = []
    target_row_count = 0
    for regi_no in expected_regi_nos:
        matching_rows = [row for row in portal_rows if row.get("regiNo") == regi_no]
        target_row_count += len(matching_rows)
        target_states.extend(str(row.get("printState", "")).strip() for row in matching_rows)
    OUTPUT_CONFIRM_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_CONFIRM_DIAGNOSTIC_PATH.write_text(
        json.dumps({
            "diagnosedAt": datetime.now(ZoneInfo("Asia/Seoul")).isoformat(timespec="seconds"),
            "expectedTargetCount": len(expected_regi_nos),
            "gridRowCount": len(portal_rows),
            "matchedTargetRowCount": target_row_count,
            "matchedTargetPrintStates": target_states,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def confirm_portal_output(timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS) -> dict[str, object]:
    """오늘 포털 인쇄 요청 이력이 포털에서 실제 출력으로 바뀌었는지 읽기 전용 확인한다."""
    lookup_date = korea_today()
    store = ParcelReceiptStore()
    try:
        candidates = [
            candidate for candidate in store.list_portal_print_requests()
            if candidate.received_at[:10] == lookup_date
        ]
    except ReceiptStoreError as error:
        raise RuntimeError("프로그램의 포털 인쇄 요청 이력을 읽지 못했습니다.") from error
    if not candidates:
        raise RuntimeError(f"{lookup_date}에 포털 출력여부를 확인할 인쇄 요청 건이 없습니다.")

    expected = [candidate.regi_no for candidate in candidates]
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = launch_epost_context(playwright)
        restore_epost_session(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30_000)
            ensure_portal_login(page, LOGIN_TIMEOUT_SECONDS)
            wait_for_print_page(page, timeout_seconds)
            work_prefix, total_before_query = apply_print_target_query(page, lookup_date, "전체")
            wait_for_query_result(page, candidates, total_before_query)
            portal_rows = read_portal_grid_rows(page, work_prefix)
            save_output_confirm_diagnostic(expected, portal_rows)
            confirmed = confirmed_output_regi_nos(expected, portal_rows)
            if set(confirmed) != set(expected):
                raise RuntimeError("포털에서 모든 인쇄 요청 건의 출력여부를 출력으로 확인하지 못했습니다.")
            store.mark_portal_print_confirmed(confirmed)
            return {
                "portalOutputConfirmed": True,
                "lookupDate": lookup_date,
                "confirmedCount": len(confirmed),
                "printCommandExecuted": False,
            }
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--confirm", action="store_true", help="포털의 출력여부를 읽기 전용으로 확인합니다.")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=PAGE_READY_TIMEOUT_SECONDS)
    args = parser.parse_args()

    if args.self_test:
        assert confirmed_output_regi_nos(["R1"], [{"regiNo": "R1", "printState": "출력"}]) == ["R1"]
        print("self-test: ok")
        return
    if not args.confirm:
        parser.error("--confirm을 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")
    print(json.dumps(confirm_portal_output(args.timeout_seconds), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
