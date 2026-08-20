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
    click_oz_viewer_print_toolbar_icon,
    oz_viewer_windows,
    print_dialog_windows,
    send_enter_to_foreground_window,
    oz_toolbar_handle_from_diagnostics,
    toolbar_button_commands,
    wait_for_oz_toolbar_diagnostics,
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
# OZ Viewer는 프린터 아이콘 한 번으로 인쇄 창을 열지만, 창 생성까지 수 초가 걸릴
# 수 있다. 짧은 재클릭은 Viewer 창 핸들이 바뀐 뒤 실패하거나 중복 요청을 만들 수
# 있으므로 한 번만 요청하고 충분히 기다린다.
PRINT_DIALOG_OPEN_ATTEMPTS = 1
PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS = 30
# 일부 드라이버 인쇄 창은 프로세스 밖 창 열거에서 보이지 않는다. 실제 인쇄는
# 열린 모달창이 전면이라는 포털 동작을 이용해 짧게 기다린 뒤 Enter를 보낸다.
PRINT_DIALOG_FOREGROUND_READY_SECONDS = 12
OZ_TOOLBAR_DIAGNOSTIC_PATH = (
    Path(__file__).resolve().parent / "output" / "epost-oz-viewer-toolbar-controls.json"
)


def close_context_after_print(context) -> None:
    """OZ Viewer가 Chromium을 먼저 닫은 경우에는 세션 정리 오류를 무시한다."""
    try:
        context.close()
    except Exception as error:
        if "Target page, context or browser has been closed" not in str(error):
            raise


def save_oz_toolbar_diagnostic(handle: int) -> None:
    """프린터 아이콘 식별용 상단 툴바 메타데이터만 로컬에 기록한다."""
    OZ_TOOLBAR_DIAGNOSTIC_PATH.parent.mkdir(parents=True, exist_ok=True)
    children = wait_for_oz_toolbar_diagnostics(handle)
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


def open_reprint_print_dialog(
    timeout_seconds: int = PAGE_READY_TIMEOUT_SECONDS,
    execute_print: bool = False,
) -> dict[str, object]:
    """출력 확인 단건의 재출력 OZ Viewer를 열어 인쇄 창을 준비한다.

    ``execute_print``는 호출자가 실제 인쇄 전송에 대한 명시적 동의를 받은 경우에만
    사용할 수 있다. 이력은 요청 전송까지만 기록하며 물리 출력 성공은 확인하지 않는다.
    """
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
            # 제목 표시만으로는 Viewer 툴바가 준비됐다고 볼 수 없다. 다만 프린터
            # 아이콘을 재클릭하면 창 핸들이 바뀌거나 중복 인쇄 요청이 될 수 있으므로
            # 한 번만 요청하고 인쇄 창 생성을 충분히 기다린다.
            time.sleep(3)
            # 실행 중인 MFC Viewer에 WM_COMMAND를 보내는 방식은 간헐적으로
            # 무시된다. 사용자와 같은 방식으로 검증된 툴바의 프린터 아이콘을
            # 실제로 한 번 클릭한다. 이 동작은 Windows 인쇄 창을 여는 것뿐이며
            # 프린터 전송은 아래의 명시적 실제 인쇄 단계에서만 수행한다.
            click_oz_viewer_print_toolbar_icon(viewer.handle)
            # 실제 OZ Viewer는 프린터 아이콘 클릭 뒤 인쇄 창 생성까지 10초 이상
            # 걸릴 수 있다. 그 사이 아이콘을 다시 누르면 중복 요청이 될 수 있으므로
            # 한 번만 요청하고 충분히 기다린다.
            if execute_print:
                time.sleep(PRINT_DIALOG_FOREGROUND_READY_SECONDS)
                send_enter_to_foreground_window()
                store.mark_windows_print_requested([candidate.regi_no for candidate in candidates])
                return {
                    "printDialogOpened": None,
                    **prepared,
                    "printDialogReviewed": False,
                    "printCommandExecuted": True,
                }
            dialog = wait_for_new_single_print_dialog(
                existing_dialogs, PRINT_DIALOG_ATTEMPT_TIMEOUT_SECONDS,
            )
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
            close_context_after_print(context)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--open-reprint-print-dialog", action="store_true")
    parser.add_argument("--execute-reprint-print", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=PAGE_READY_TIMEOUT_SECONDS)
    args = parser.parse_args()
    if not (args.open_reprint_print_dialog or args.execute_reprint_print):
        parser.error("--open-reprint-print-dialog 또는 --execute-reprint-print를 지정하세요.")
    if args.open_reprint_print_dialog and args.execute_reprint_print:
        parser.error("인쇄 창 확인과 실제 인쇄는 함께 지정할 수 없습니다.")
    if args.timeout_seconds <= 0:
        parser.error("--timeout-seconds는 1 이상이어야 합니다.")
    print(json.dumps(open_reprint_print_dialog(
        args.timeout_seconds, execute_print=args.execute_reprint_print,
    ), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
