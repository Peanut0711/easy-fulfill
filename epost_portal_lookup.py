"""우체국 계약소포 신규출력 목록을 읽기 전용으로 대조한다.

이 도구는 오늘 날짜와 `신규출력` 조건으로 포털 목록을 조회할 수 있지만, 행 체크,
운송장출력 팝업 열기, OZ Viewer 실행, Windows 인쇄 명령은 수행하지 않는다.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import json
from pathlib import Path
import re
import time
from typing import Iterable
from zoneinfo import ZoneInfo

from epost_portal_diagnostic import looks_like_print_page
from epost_portal_session import (
    LOGIN_TIMEOUT_SECONDS,
    PORTAL_URL,
    ensure_portal_login,
    launch_epost_context,
    restore_epost_session,
)
from post_parcel_receipt_store import ParcelReceiptStore, PrintCandidate, ReceiptStoreError


PAGE_READY_TIMEOUT_SECONDS = 600
QUERY_RESULT_TIMEOUT_SECONDS = 15
QUERY_RESULT_MINIMUM_WAIT_SECONDS = 3
START_DATE_SELECTOR = "input[id$='Div00_ipbStartDate_calendaredit_input']"
END_DATE_SELECTOR = "input[id$='Div00_ipbEndDate_calendaredit_input']"
PRINT_TARGET_SELECTOR = "input[id$='Div00_cboPrintYN_comboedit_input']"
REGI_START_SELECTOR = "input[id$='Div00_ipbStartNum_input']"
REGI_END_SELECTOR = "input[id$='Div00_ipbEndNum_input']"
GRID_REGI_NO_COLUMN = 3
GRID_PRINT_STATE_COLUMN = 4
QUERY_CONTROL_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-portal-query-controls.json"
)
GRID_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-portal-grid-diagnostic.json"
)
TEXT_CONTROL_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-portal-text-control-diagnostic.json"
)


def korea_today() -> str:
    """포털 검색 조건에 사용할 한국 날짜를 YYYY-MM-DD로 반환한다."""
    return datetime.now(ZoneInfo("Asia/Seoul")).date().isoformat()


def pending_candidates_for_date(
    candidates: Iterable[PrintCandidate], lookup_date: str,
) -> list[PrintCandidate]:
    """해당 날짜에 프로그램이 실제 접수해 아직 인쇄하지 않은 건만 고른다."""
    return [
        candidate for candidate in candidates
        if candidate.received_at[:10] == lookup_date and candidate.regi_no
    ]


def total_count_from_page_text(body_text: str) -> int | None:
    """포털 화면의 총건수 표기를 숫자로만 추출한다."""
    match = re.search(r"총\s*건수\s*[:：]\s*(\d+)", body_text or "")
    return int(match.group(1)) if match else None


def matched_registration_numbers(body_text: str, candidates: Iterable[PrintCandidate]) -> list[str]:
    """포털 본문에서 프로그램이 이미 알고 있는 등기번호의 노출 여부만 확인한다."""
    return [
        candidate.regi_no for candidate in candidates
        if re.search(rf"(?<!\d){re.escape(candidate.regi_no)}(?!\d)", body_text or "")
    ]


def grid_row_indexes_from_ids(control_ids: Iterable[str]) -> list[int]:
    """Nexacro `grdList`의 실제 본문 행 ID에서 행 번호만 추출한다."""
    indexes = set()
    for control_id in control_ids:
        match = re.search(r"_grdList_body_gridrow_(\d+)$", control_id or "")
        if match:
            indexes.add(int(match.group(1)))
    return sorted(indexes)


def normalize_registration_number(value: str) -> str:
    """화면 표기의 하이픈·공백과 관계없이 등기번호를 비교한다."""
    return "".join(re.findall(r"\d", value or ""))


def verify_target_rows(expected_regi_nos: Iterable[str], rows: Iterable[dict[str, object]]) -> dict[str, object]:
    """이번 접수 건만 정확히 1회씩 미출력 행으로 존재하는지 판정한다."""
    expected = [normalize_registration_number(value) for value in expected_regi_nos]
    expected = [value for value in expected if value]
    expected_set = set(expected)
    normalized_rows = [
        {
            **row,
            "regiNo": normalize_registration_number(str(row.get("regiNo", ""))),
            "printState": str(row.get("printState", "") or "").strip(),
        }
        for row in rows
    ]
    target_rows = [row for row in normalized_rows if row["regiNo"] in expected_set]
    target_counts = {regi_no: 0 for regi_no in expected}
    for row in target_rows:
        target_counts[row["regiNo"]] += 1
    missing = [regi_no for regi_no, count in target_counts.items() if count == 0]
    duplicates = [regi_no for regi_no, count in target_counts.items() if count > 1]
    unexpected_count = sum(
        1 for row in normalized_rows
        if row["regiNo"] and row["regiNo"] not in expected_set
    )
    printed = [
        row["regiNo"] for row in target_rows
        if row["printState"] != "미출력"
    ]
    missing_checkbox = [
        row["regiNo"] for row in target_rows
        if not row.get("hasCheckbox")
    ]
    return {
        "targetRows": target_rows,
        "missingRegiNos": missing,
        "duplicateRegiNos": duplicates,
        "unexpectedRowCount": unexpected_count,
        "printedRegiNos": printed,
        "missingCheckboxRegiNos": missing_checkbox,
        "verified": not (missing or duplicates or unexpected_count or printed or missing_checkbox),
    }


def _single_visible(page, selector: str, label: str):
    locator = page.locator(selector)
    visible = [locator.nth(index) for index in range(locator.count()) if locator.nth(index).is_visible()]
    if len(visible) != 1:
        raise RuntimeError(f"{label} 입력칸을 하나로 식별하지 못했습니다. (발견: {len(visible)}개)")
    return visible[0]


def _single_visible_id(page, control_id: str, label: str):
    """현재 작업영역에서 조합한 정확한 ID의 표시 컨트롤 하나를 반환한다."""
    locator = page.locator(f'[id="{control_id}"]')
    visible = [locator.nth(index) for index in range(locator.count()) if locator.nth(index).is_visible()]
    if len(visible) != 1:
        raise RuntimeError(f"{label} 입력칸을 현재 운송장출력 영역에서 하나로 식별하지 못했습니다. (발견: {len(visible)}개)")
    return visible[0]


def _click_visible_exact_text(page, text: str, label: str) -> None:
    """Nexacro의 div 기반 컨트롤도 포함해 정확한 표시 텍스트의 단일 항목만 누른다."""
    locator = page.get_by_text(text, exact=True)
    visible = [locator.nth(index) for index in range(locator.count()) if locator.nth(index).is_visible()]
    identified = [item for item in visible if item.get_attribute("id")]
    candidates = identified or visible
    if len(candidates) == 1:
        candidates[0].click()
        return
    outer_id = preferred_outer_text_control_id([
        item.get_attribute("id") or "" for item in identified
    ])
    if outer_id:
        page.locator(f'[id="{outer_id}"]').click()
        return
    if len(candidates) != 1:
        save_text_control_diagnostic(candidates, label)
        raise RuntimeError(f"{label} 컨트롤을 하나로 식별하지 못했습니다. (발견: {len(candidates)}개)")


def preferred_outer_text_control_id(control_ids: Iterable[str]) -> str:
    """부모 컨트롤과 `TextBoxElement` 자식이 함께 잡힌 경우 부모만 선택한다."""
    unique_ids = [control_id for control_id in dict.fromkeys(control_ids) if control_id]
    roots = [
        control_id for control_id in unique_ids
        if all(other_id == control_id or other_id.startswith(control_id) for other_id in unique_ids)
    ]
    return roots[0] if len(roots) == 1 else ""


def print_target_combo_keys(print_target_text: str) -> tuple[str, ...]:
    """텍스트가 화면 전체에 반복되는 출력대상은 콤보 키 선택 규칙을 반환한다."""
    return ("Home", "Enter") if print_target_text == "전체" else ()


def selected_print_target_matches(actual: str, expected: str) -> bool:
    """콤보 입력값이 요청한 출력대상과 같은지 공백을 무시하고 확인한다."""
    return str(actual or "").strip() == str(expected or "").strip()


def selected_date_matches(actual: str, expected: str) -> bool:
    """날짜 입력칸이 요청한 날짜로 실제 반영됐는지 구분 문자와 무관하게 확인한다."""
    actual_digits = "".join(re.findall(r"\d", str(actual or "")))
    expected_digits = "".join(re.findall(r"\d", str(expected or "")))
    return bool(actual_digits and expected_digits and actual_digits == expected_digits)


def set_calendar_date(date_input, lookup_date: str) -> None:
    """Nexacro 날짜 입력칸에 키보드 이벤트로 날짜를 넣고 반영을 확인한다.

    `fill()`만 사용하면 화면에 보이는 값은 바뀌어도 Nexacro의 날짜 편집 이벤트가
    발생하지 않는 환경이 있다. 실제 사용자가 입력하는 것과 같이 전체 선택·삭제·입력·
    Enter/Tab 순서로 넣는다.
    """
    compact_date = "".join(re.findall(r"\d", lookup_date))
    for value in (lookup_date, compact_date):
        date_input.click()
        date_input.press("Control+A")
        date_input.press("Backspace")
        date_input.type(value, delay=25)
        date_input.press("Enter")
        date_input.press("Tab")
        deadline = time.monotonic() + 1
        while time.monotonic() < deadline:
            if selected_date_matches(date_input.input_value(), lookup_date):
                return
            time.sleep(0.1)
    raise RuntimeError("검색 날짜 입력칸에 요청한 날짜를 반영하지 못했습니다.")


def registration_filter_values(regi_no: str) -> tuple[str, str]:
    """등기번호 범위 필터에 넣을 숫자만 남긴 시작·끝 값을 만든다."""
    normalized = normalize_registration_number(regi_no)
    if not normalized:
        raise RuntimeError("등기번호 필터에 사용할 숫자가 없습니다.")
    return normalized, normalized


def registration_input_matches(actual: str, expected: str) -> bool:
    """Nexacro 등기번호 입력칸이 숫자 기준으로 요청값과 같은지 확인한다."""
    return normalize_registration_number(actual) == normalize_registration_number(expected)


def set_registration_input(registration_input, regi_no: str, label: str) -> None:
    """Nexacro 등기번호 입력칸에 키보드 이벤트로 값을 넣고 반영을 확인한다."""
    normalized = normalize_registration_number(regi_no)
    registration_input.click()
    registration_input.press("Control+A")
    registration_input.press("Backspace")
    registration_input.type(normalized, delay=10)
    registration_input.press("Enter")
    registration_input.press("Tab")
    deadline = time.monotonic() + 1
    while time.monotonic() < deadline:
        if registration_input_matches(registration_input.input_value(), normalized):
            return
        time.sleep(0.1)
    raise RuntimeError(f"{label} 입력칸에 요청한 등기번호를 반영하지 못했습니다.")


def save_text_control_diagnostic(candidates, label: str) -> None:
    """모호한 표시 텍스트 후보의 식별자·위치만 로컬에 기록한다.

    표시 텍스트·입력값·그리드 행 데이터는 저장하지 않는다.
    """
    controls = []
    for candidate in candidates:
        box = candidate.bounding_box() or {}
        controls.append({
            "id": candidate.get_attribute("id") or "",
            "left": round(box.get("x", 0)),
            "top": round(box.get("y", 0)),
            "width": round(box.get("width", 0)),
            "height": round(box.get("height", 0)),
        })
    TEXT_CONTROL_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    TEXT_CONTROL_DIAGNOSTIC_PATH.write_text(
        json.dumps({"label": label, "controls": controls}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _work_area_prefix(date_input) -> str:
    """현재 운송장출력 탭의 고정 작업 영역 접두어를 입력칸 ID에서 얻는다."""
    control_id = date_input.get_attribute("id") or ""
    prefix, marker, _ = control_id.rpartition("Div00_")
    if not marker:
        raise RuntimeError("운송장출력 작업 영역을 입력칸 ID에서 식별하지 못했습니다.")
    return prefix


def _click_query_in_work_area(page, work_prefix: str) -> None:
    """다른 탭·중첩 요소가 아닌 현재 작업 영역의 조회 버튼만 누른다."""
    buttons = page.locator("[id]").evaluate_all(
        """
        (elements, options) => elements.filter(element => {
            const style = getComputedStyle(element);
            const text = (element.innerText || '').trim();
            return element.id.startsWith(options.prefix)
                && /btn/i.test(element.id)
                && text === options.text
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
        {"prefix": work_prefix, "text": "조회"},
    )
    button_ids = [item["id"] for item in buttons]
    button_id = preferred_query_button_id(button_ids)
    if not button_id:
        save_query_control_diagnostic(work_prefix, buttons)
        raise RuntimeError(f"조회 컨트롤을 현재 운송장출력 영역에서 하나로 식별하지 못했습니다. (발견: {len(button_ids)}개)")
    page.locator(f'[id="{button_id}"]').click()


def preferred_query_button_id(button_ids: Iterable[str]) -> str:
    """Nexacro의 버튼 본체와 내부 요소가 중복될 때 버튼 본체만 고른다."""
    unique_ids = list(dict.fromkeys(button_ids))
    if len(unique_ids) == 1:
        return unique_ids[0]
    roots = [
        control_id for control_id in unique_ids
        if re.search(r"_btn(?:InfoLoad|Search|Srch|Qry|Inq)$", control_id, flags=re.IGNORECASE)
    ]
    return roots[0] if len(roots) == 1 else ""


def save_query_control_diagnostic(work_prefix: str, buttons: list[dict[str, object]]) -> None:
    """모호한 조회 컨트롤의 DOM 식별자·위치만 로컬에 저장한다."""
    payload = {
        "diagnosedAt": datetime.now(ZoneInfo("Asia/Seoul")).isoformat(timespec="seconds"),
        "workPrefix": work_prefix,
        "queryControls": buttons,
    }
    QUERY_CONTROL_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    QUERY_CONTROL_DIAGNOSTIC_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8",
    )


def save_grid_structure_diagnostic(page, work_prefix: str) -> int:
    """조회된 Nexacro 그리드의 조작 구조만 기록하고 셀 값·행 텍스트는 제외한다."""
    controls = page.locator("[id]").evaluate_all(
        """
        (elements, prefix) => elements.filter(element =>
            element.id.startsWith(prefix) && /grd|grid/i.test(element.id)
        ).map(element => ({
            id: element.id,
            tag: element.tagName.toLowerCase(),
            role: element.getAttribute('role') || '',
            type: element.getAttribute('type') || '',
            ariaLabel: element.getAttribute('aria-label') || '',
            title: element.getAttribute('title') || '',
            className: typeof element.className === 'string' ? element.className : ''
        }))
        """,
        work_prefix,
    )
    payload = {
        "diagnosedAt": datetime.now(ZoneInfo("Asia/Seoul")).isoformat(timespec="seconds"),
        "workPrefix": work_prefix,
        "gridControlCount": len(controls),
        "gridControls": controls,
    }
    GRID_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    GRID_DIAGNOSTIC_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8",
    )
    return len(controls)


def _visible_grid_row_indexes(page, work_prefix: str) -> list[int]:
    control_ids = page.locator("[id]").evaluate_all(
        """
        (elements, prefix) => elements.filter(element => {
            const style = getComputedStyle(element);
            return element.id.startsWith(prefix)
                && /_grdList_body_gridrow_\\d+$/.test(element.id)
                && style.display !== 'none'
                && style.visibility !== 'hidden'
                && element.getClientRects().length > 0;
        }).map(element => element.id)
        """,
        work_prefix,
    )
    return grid_row_indexes_from_ids(control_ids)


def _grid_cell_text(page, work_prefix: str, row_index: int, column_index: int) -> str:
    control_id = (
        f"{work_prefix}grdList_body_gridrow_{row_index}"
        f"_cell_{row_index}_{column_index}"
    )
    locator = page.locator(f'[id="{control_id}"]')
    if locator.count() != 1 or not locator.is_visible():
        return ""
    return locator.inner_text(timeout=3_000).strip()


def read_portal_grid_rows(page, work_prefix: str) -> list[dict[str, object]]:
    """등기번호·출력여부·체크박스 유무만 읽고 다른 행 정보는 반환하지 않는다."""
    rows = []
    for row_index in _visible_grid_row_indexes(page, work_prefix):
        checkbox_id = (
            f"{work_prefix}grdList_body_gridrow_{row_index}"
            f"_cell_{row_index}_0_controlcheckbox"
        )
        rows.append({
            "rowIndex": row_index,
            "regiNo": _grid_cell_text(page, work_prefix, row_index, GRID_REGI_NO_COLUMN),
            "printState": _grid_cell_text(page, work_prefix, row_index, GRID_PRINT_STATE_COLUMN),
            "hasCheckbox": page.locator(f'[id="{checkbox_id}"]').count() == 1,
        })
    return rows


def apply_print_target_query(
    page, lookup_date: str, print_target_text: str, regi_no: str = "", regi_no_end: str = "",
) -> tuple[str, int | None]:
    """날짜와 출력대상 조건을 입력한 뒤 조회만 실행한다."""
    start_date = _single_visible(page, START_DATE_SELECTOR, "검색 시작일")
    end_date = _single_visible(page, END_DATE_SELECTOR, "검색 종료일")
    set_calendar_date(start_date, lookup_date)
    set_calendar_date(end_date, lookup_date)
    work_prefix = _work_area_prefix(start_date)

    if regi_no:
        start_regi, default_end_regi = registration_filter_values(regi_no)
        end_regi = normalize_registration_number(regi_no_end) or default_end_regi
        start_input = _single_visible_id(
            page, f"{work_prefix}Div00_ipbStartNum_input", "등기번호 시작",
        )
        end_input = _single_visible_id(
            page, f"{work_prefix}Div00_ipbEndNum_input", "등기번호 끝",
        )
        set_registration_input(start_input, start_regi, "등기번호 시작")
        set_registration_input(end_input, end_regi, "등기번호 끝")

    print_target = _single_visible(page, PRINT_TARGET_SELECTOR, "출력대상")
    print_target.click()
    combo_keys = print_target_combo_keys(print_target_text)
    if combo_keys:
        # `전체`는 화면의 다른 필터 값에도 반복돼 DOM 텍스트만으로는 안전하게
        # 구별할 수 없다. 이미 식별한 출력대상 콤보에서 첫 항목을 선택한다.
        for key in combo_keys:
            print_target.press(key)
    else:
        _click_visible_exact_text(page, print_target_text, "출력대상")
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline:
        if selected_print_target_matches(print_target.input_value(), print_target_text):
            break
        time.sleep(0.1)
    else:
        raise RuntimeError("출력대상 콤보 값이 요청한 조건으로 바뀌었는지 확인하지 못했습니다.")
    total_before_query = total_count_from_page_text(page.locator("body").inner_text(timeout=3_000))
    _click_query_in_work_area(page, work_prefix)
    return work_prefix, total_before_query


def apply_new_print_query(page, lookup_date: str) -> str:
    """오늘 날짜와 신규출력 조건을 입력한 뒤 읽기 전용 조회만 실행한다."""
    work_prefix, _ = apply_print_target_query(page, lookup_date, "신규출력")
    return work_prefix


def wait_for_query_result(
    page, candidates: list[PrintCandidate], total_before_query: int | None = None,
) -> tuple[str, int | None]:
    """조회 결과가 반영될 때까지 기다리되, 개인정보를 파일로 저장하지 않는다."""
    expected = {candidate.regi_no for candidate in candidates}
    deadline = time.monotonic() + QUERY_RESULT_TIMEOUT_SECONDS
    # 포털의 가상 그리드는 조회 직후에도 이전 총건수를 계속 표시할 수 있다.
    # 최소 대기 뒤에는 대상 행을 아직 찾지 못했더라도 다음의 엄격한 행 검증 단계로
    # 넘긴다. 그 단계에서 대상 등기번호가 1건이 아니면 체크나 팝업 열기는 불가능하다.
    result_check_after = time.monotonic() + QUERY_RESULT_MINIMUM_WAIT_SECONDS
    last_text = ""
    while time.monotonic() < deadline:
        last_text = page.locator("body").inner_text(timeout=3_000)
        has_expected = any(
            re.search(rf"(?<!\d){re.escape(regi_no)}(?!\d)", last_text)
            for regi_no in expected
        )
        current_total = total_count_from_page_text(last_text)
        if has_expected or (time.monotonic() >= result_check_after and (
            current_total is None or current_total != total_before_query
        )):
            return last_text, current_total
        # 날짜 변경 과정에서 Nexacro가 이미 총건수를 갱신한 경우에는 기준 건수도
        # 새 값이 될 수 있다. 이때는 가상 그리드의 실제 행 탐색으로 판정한다.
        if time.monotonic() >= result_check_after + 2:
            return last_text, current_total
        time.sleep(0.5)
    raise RuntimeError("포털 조회 결과 반영을 확인하지 못했습니다. 조회 조건과 포털 상태를 확인해 주세요.")


def wait_for_print_page(page, timeout_seconds: int) -> None:
    """사용자가 팝업을 닫고 운송장출력 화면으로 이동할 때까지 기다린다."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if looks_like_print_page(page):
            return
        time.sleep(0.5)
    raise RuntimeError("운송장출력 화면을 확인하지 못했습니다. 계약소포 > 운송장출력으로 이동해 주세요.")


def lookup(timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS) -> dict[str, object]:
    """오늘의 프로그램 접수 대기 건과 포털 신규출력 조회 결과를 대조한다."""
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
            work_prefix = apply_new_print_query(page, lookup_date)
            body_text, portal_row_count = wait_for_query_result(page, candidates)
            matched = matched_registration_numbers(body_text, candidates)
            expected = [candidate.regi_no for candidate in candidates]
            missing = [regi_no for regi_no in expected if regi_no not in matched]
            grid_control_count = save_grid_structure_diagnostic(page, work_prefix)
            grid_verification = verify_target_rows(
                expected,
                read_portal_grid_rows(page, work_prefix),
            )
            return {
                "lookedUp": True,
                "lookupDate": lookup_date,
                "expectedRegiNos": expected,
                "matchedRegiNos": matched,
                "missingRegiNos": missing,
                "portalRowCount": portal_row_count,
                "gridControlCount": grid_control_count,
                "targetRows": grid_verification["targetRows"],
                "duplicateRegiNos": grid_verification["duplicateRegiNos"],
                "unexpectedRowCount": grid_verification["unexpectedRowCount"],
                "printedRegiNos": grid_verification["printedRegiNos"],
                "missingCheckboxRegiNos": grid_verification["missingCheckboxRegiNos"],
                "targetRowsVerified": grid_verification["verified"],
                "selectionOrPrintExecuted": False,
            }
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--lookup", action="store_true", help="오늘 신규출력 목록을 읽기 전용으로 대조합니다.")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=PAGE_READY_TIMEOUT_SECONDS)
    args = parser.parse_args()

    if args.self_test:
        assert total_count_from_page_text("총건수 : 2") == 2
        assert korea_today().count("-") == 2
        print("self-test: ok")
        return
    if not args.lookup:
        parser.error("--lookup을 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")

    print(json.dumps(lookup(args.timeout_seconds), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
