"""우체국 단건 재출력의 OZ Viewer에서 Windows 인쇄 창 열림만 확인한다.

포털의 재출력 인쇄 요청과 준비된 OZ Viewer의 프린터 아이콘 클릭까지 수행하지만 Windows 인쇄 창의
용지 선택·확인·취소 및 실제 프린터 전송은 수행하지 않는다.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import time

from epost_desktop_windows import (
    activate_window,
    click_oz_viewer_print_toolbar_button,
    oz_viewer_windows,
    print_dialog_windows,
    top_toolbar_child_diagnostics,
    oz_toolbar_handle_from_diagnostics,
    toolbar_button_commands,
    wait_for_new_single_oz_viewer,
    wait_for_new_single_print_dialog,
    wait_for_window_close,
)
from epost_portal_lookup import PAGE_READY_TIMEOUT_SECONDS, wait_for_print_page
from epost_portal_oz_viewer import OZ_VIEWER_TIMEOUT_SECONDS, click_popup_print_button
from epost_portal_reprint_popup import prepare_verified_reprint_popup
from epost_portal_session import (
    LOGIN_TIMEOUT_SECONDS,
    PORTAL_URL,
    ensure_portal_login,
    launch_epost_context,
    restore_epost_session,
)
from post_parcel_receipt_store import ParcelReceiptStore, ReceiptStoreError


PRINT_DIALOG_TIMEOUT_SECONDS = 15
PRINT_DIALOG_OPEN_ATTEMPTS = 4
PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS = 3
OZ_TOOLBAR_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-oz-viewer-toolbar-controls.json"
)


def save_oz_toolbar_diagnostic(handle: int) -> None:
    """프린터 아이콘 식별용 상단 툴바 메타데이터만 로컬에 기록한다."""
    OZ_TOOLBAR_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    children = top_toolbar_child_diagnostics(handle)
    toolbar_handle = oz_toolbar_handle_from_diagnostics(children)
    command_read_error = ""
    commands = []
    if toolbar_handle:
        try:
            commands = toolbar_button_commands(toolbar_handle)
        except RuntimeError as error:
            # 진단용 메모리 읽기 실패는 OZ Viewer의 인쇄 명령 실행을 막지 않는다.
            command_read_error = str(error)
    OZ_TOOLBAR_DIAGNOSTIC_PATH.write_text(
        json.dumps(
            {
                "toolbarChildren": children,
                "toolbarButtonCommands": commands,
                "commandReadError": command_read_error,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def open_reprint_print_dialog(timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS) -> dict[str, object]:
    """출력 확인 단건의 재출력 OZ Viewer를 열어 인쇄 창만 확인한다."""
    store = ParcelReceiptStore()
    try:
        candidates = store.list_portal_print_confirmed()
    except ReceiptStoreError as error:
        raise RuntimeError("프로그램의 포털 출력 확인 이력을 읽지 못했습니다.") from error
    if len(candidates) != 1:
        raise RuntimeError("단건 인쇄 창 확인은 포털 출력 확인 건이 정확히 1건일 때만 실행할 수 있습니다.")
    if oz_viewer_windows():
        raise RuntimeError("기존 OZ Report Viewer 창이 열려 있습니다. 먼저 닫아 주세요.")
    existing_dialogs = {window.handle for window in print_dialog_windows()}
    if existing_dialogs:
        raise RuntimeError("기존 Windows 인쇄 창이 열려 있습니다. 확인을 누르지 말고 먼저 닫아 주세요.")

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
            save_oz_toolbar_diagnostic(viewer.handle)
            # 제목 표시만으로는 Viewer 툴바가 준비됐다고 볼 수 없다. 동일한 프린터
            # 아이콘만 짧은 간격으로 재시도하고, 인쇄 창을 감지하는 즉시 멈춘다.
            time.sleep(3)
            dialog = None
            for _ in range(PRINT_DIALOG_OPEN_ATTEMPTS):
                click_oz_viewer_print_toolbar_button(viewer.handle)
                try:
                    dialog = wait_for_new_single_print_dialog(
                        existing_dialogs, PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS,
                    )
                    break
                except RuntimeError:
                    time.sleep(0.5)
            if dialog is None:
                raise RuntimeError("Windows 인쇄 창이 열렸는지 확인하지 못했습니다.")
            # 인쇄 창이 OZ Viewer 뒤에서 열리면 사용자가 확인할 수 없으므로, 확인된
            # Windows 인쇄 창의 제목 표시줄만 활성화해 전면에 보이게 한다.
            activate_window(dialog.handle)
            if not wait_for_window_close(dialog.handle, timeout_seconds):
                raise RuntimeError("Windows 인쇄 창 검토 시간이 만료됐습니다. 확인을 누르지 말고 창을 닫은 뒤 다시 실행해 주세요.")
            if not wait_for_window_close(viewer.handle, timeout_seconds):
                raise RuntimeError("OZ Report Viewer가 아직 열려 있습니다. 프린터 아이콘을 누르지 말고 창을 닫아 주세요.")
            return {
                "printDialogOpened": True,
                **prepared,
                "printDialogReviewed": True,
                "printCommandExecuted": False,
            }
        finally:
            context.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--open-reprint-print-dialog", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=PAGE_READY_TIMEOUT_SECONDS)
    args = parser.parse_args()
    if not args.open_reprint_print_dialog:
        parser.error("--open-reprint-print-dialog를 지정하세요.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")
    print(json.dumps(open_reprint_print_dialog(args.timeout_seconds), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
