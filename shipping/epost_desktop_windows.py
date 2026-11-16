"""우체국 운송장 출력 단계에서 필요한 Windows 창 제어 도구.

OZ Viewer와 인쇄 대화상자는 웹 브라우저 밖의 별도 프로그램이다. 실제 인쇄를
실행하지 않으며, 현재 구현 범위에서는 OZ Viewer의 인쇄 창 열기와 용지 선택만
수행한다.
"""

from __future__ import annotations

import ctypes
import os
import time
from ctypes import wintypes
from dataclasses import dataclass


OZ_VIEWER_TITLE_FRAGMENT = "오즈 리포트 뷰어"
PRINT_DIALOG_TITLE = "인쇄"
SW_RESTORE = 9
VK_CONTROL = 0x11
VK_P = 0x50
VK_RETURN = 0x0D
KEYEVENTF_KEYUP = 0x0002
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
WINDOW_ACTIVATE_TIMEOUT_SECONDS = 8
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_COMMAND = 0x0111
IDOK = 1
MK_LBUTTON = 0x0001
WM_USER = 0x0400
TB_GETBUTTON = WM_USER + 23
TB_BUTTONCOUNT = WM_USER + 24
PROCESS_VM_OPERATION = 0x0008
PROCESS_VM_READ = 0x0010
PROCESS_VM_WRITE = 0x0020
PROCESS_QUERY_INFORMATION = 0x0400
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04
MFC_ID_FILE_PRINT = 0xE107
# 현재 계약소포가 실행하는 OZ Report Viewer의 상단 MFC 툴바는 첫 번째가 저장,
# 두 번째가 인쇄 버튼이다. OZ는 일반 MFC의 ``ID_FILE_PRINT`` 대신 실행마다
# 할당되는 자체 명령 ID(확인 사례: 32832)를 사용한다. 따라서 숫자를 고정하지 않고
# 실제 실행 중인 툴바의 두 번째 버튼 ID만 읽어 사용한다.
OZ_PRINT_TOOLBAR_BUTTON_INDEX = 1
# OZ Viewer MFC 툴바 내부 기준 프린터 아이콘 중심점. 첫 번째 저장 아이콘 다음의
# 두 번째 아이콘이며, 뷰어 창의 화면 좌표가 아닌 툴바 클라이언트 좌표를 사용한다.
OZ_PRINT_TOOLBAR_CLIENT_X = 49
OZ_PRINT_TOOLBAR_CLIENT_Y = 20


def enable_per_monitor_dpi_awareness() -> None:
    """창 사각형과 실제 마우스 좌표를 같은 DPI 기준으로 읽는다.

    Windows는 DPI 비인식 프로세스의 ``GetWindowRect`` 결과를 논리 좌표로
    가상화할 수 있지만 ``SetCursorPos``는 실제 화면 좌표를 사용한다. 이 모듈은
    별도 단명 프로세스에서 실행되므로 스레드를 모니터별 DPI 인식으로 전환한다.
    """
    if os.name != "nt":
        return
    setter = getattr(ctypes.windll.user32, "SetThreadDpiAwarenessContext", None)
    if setter is not None:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        setter(ctypes.c_void_p(-4))


@dataclass(frozen=True)
class DesktopWindow:
    """개인정보를 포함하지 않는 최상위 창 식별 정보."""

    handle: int
    title: str


def matching_visible_windows(
    windows: list[DesktopWindow], title_fragment: str,
) -> list[DesktopWindow]:
    """표시 중인 창 제목에 지정 문구가 든 항목만 남긴다."""
    return [window for window in windows if title_fragment in window.title]


def visible_top_level_windows() -> list[DesktopWindow]:
    """현재 Windows의 표시 중인 최상위 창 제목만 열거한다."""
    if os.name != "nt":
        raise RuntimeError("OZ Viewer 창 확인은 Windows에서만 지원합니다.")

    windows: list[DesktopWindow] = []
    user32 = ctypes.windll.user32
    enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    @enum_proc_type
    def collect(handle, _lparam):
        if not user32.IsWindowVisible(handle):
            return True
        length = user32.GetWindowTextLengthW(handle)
        if length <= 0:
            return True
        title_buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(handle, title_buffer, len(title_buffer))
        title = title_buffer.value.strip()
        if title:
            windows.append(DesktopWindow(handle=int(handle), title=title))
        return True

    if not user32.EnumWindows(collect, 0):
        raise ctypes.WinError(ctypes.get_last_error())
    return windows


def oz_viewer_windows() -> list[DesktopWindow]:
    """현재 표시 중인 OZ Report Viewer 창만 반환한다."""
    return matching_visible_windows(visible_top_level_windows(), OZ_VIEWER_TITLE_FRAGMENT)


def top_toolbar_child_diagnostics(handle: int) -> list[dict[str, object]]:
    """OZ Viewer 상단 툴바의 Win32 자식 창 메타데이터만 읽는다.

    문서 영역·송장 내용은 제외하고, 창의 클래스·텍스트·상대 위치만 반환한다.
    """
    if os.name != "nt":
        raise RuntimeError("Windows 창 진단은 Windows에서만 지원합니다.")
    user32 = ctypes.windll.user32
    parent_rect = wintypes.RECT()
    if not user32.GetWindowRect(wintypes.HWND(handle), ctypes.byref(parent_rect)):
        raise RuntimeError("OZ Report Viewer 창 위치를 읽지 못했습니다.")
    children: list[dict[str, object]] = []
    enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    @enum_proc_type
    def collect(child, _lparam):
        rect = wintypes.RECT()
        if not user32.IsWindowVisible(child) or not user32.GetWindowRect(child, ctypes.byref(rect)):
            return True
        relative_top = rect.top - parent_rect.top
        if relative_top < 0 or relative_top >= 120 or rect.right <= rect.left or rect.bottom <= rect.top:
            return True
        class_buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(child, class_buffer, len(class_buffer))
        text_length = user32.GetWindowTextLengthW(child)
        text_buffer = ctypes.create_unicode_buffer(min(text_length, 120) + 1)
        user32.GetWindowTextW(child, text_buffer, len(text_buffer))
        children.append({
            "handle": int(child),
            "className": class_buffer.value,
            "text": text_buffer.value.strip(),
            "left": rect.left - parent_rect.left,
            "top": relative_top,
            "width": rect.right - rect.left,
            "height": rect.bottom - rect.top,
        })
        return True

    if not user32.EnumChildWindows(wintypes.HWND(handle), collect, 0):
        raise ctypes.WinError(ctypes.get_last_error())
    return children


def oz_toolbar_handle_from_diagnostics(children: list[dict[str, object]]) -> int:
    """상단에 있는 유일한 MFC 툴바 핸들만 선택한다."""
    matches = [
        int(child["handle"])
        for child in children
        if str(child.get("className", "")).startswith("Afx:ToolBar")
        and 40 <= int(child.get("top", -1)) <= 80
    ]
    return matches[0] if len(matches) == 1 else 0


def wait_for_oz_toolbar_diagnostics(
    handle: int, timeout_seconds: float = 8,
) -> list[dict[str, object]]:
    """새 OZ Viewer의 상단 툴바가 실제로 준비될 때까지 짧게 기다린다."""
    deadline = time.monotonic() + timeout_seconds
    last_error: RuntimeError | None = None
    while time.monotonic() < deadline:
        try:
            children = top_toolbar_child_diagnostics(handle)
            if oz_toolbar_handle_from_diagnostics(children):
                return children
        except RuntimeError as error:
            last_error = error
        time.sleep(0.25)
    if last_error is not None:
        raise last_error
    raise RuntimeError("OZ Report Viewer의 상단 MFC 툴바가 준비됐는지 확인하지 못했습니다.")


class _ToolbarButton(ctypes.Structure):
    """Win32 `TBBUTTON`의 현재 프로세스 포인터 크기 대응 구조다."""

    _fields_ = [
        ("iBitmap", ctypes.c_int),
        ("idCommand", ctypes.c_int),
        ("fsState", ctypes.c_ubyte),
        ("fsStyle", ctypes.c_ubyte),
        ("bReserved", ctypes.c_ubyte * (6 if ctypes.sizeof(ctypes.c_void_p) == 8 else 2)),
        ("dwData", ctypes.c_void_p),
        ("iString", ctypes.c_void_p),
    ]


def toolbar_button_commands(toolbar_handle: int) -> list[dict[str, int]]:
    """MFC 툴바의 버튼 순서·명령 ID만 읽는다.

    `TB_GETBUTTON`은 다른 프로세스의 버퍼 주소를 요구하므로, 해당 프로세스에
    임시 버퍼를 할당해 결과만 읽고 즉시 해제한다. 버튼을 누르거나 상태를 변경하지는 않는다.
    """
    if os.name != "nt":
        raise RuntimeError("MFC 툴바 진단은 Windows에서만 지원합니다.")
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    pid = wintypes.DWORD()
    if not user32.GetWindowThreadProcessId(wintypes.HWND(toolbar_handle), ctypes.byref(pid)):
        raise RuntimeError("MFC 툴바 프로세스를 식별하지 못했습니다.")
    process = kernel32.OpenProcess(
        PROCESS_QUERY_INFORMATION | PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE,
        False,
        pid.value,
    )
    if not process:
        raise RuntimeError("MFC 툴바 명령을 읽을 권한이 없습니다.")
    remote_buffer = None
    try:
        remote_buffer = kernel32.VirtualAllocEx(
            process, None, ctypes.sizeof(_ToolbarButton), MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE,
        )
        if not remote_buffer:
            raise RuntimeError("MFC 툴바 진단용 임시 버퍼를 만들지 못했습니다.")
        count = int(user32.SendMessageW(wintypes.HWND(toolbar_handle), TB_BUTTONCOUNT, 0, 0))
        if count <= 0:
            raise RuntimeError("MFC 툴바 버튼 수를 읽지 못했습니다.")
        commands = []
        for index in range(count):
            if not user32.SendMessageW(wintypes.HWND(toolbar_handle), TB_GETBUTTON, index, remote_buffer):
                raise RuntimeError("MFC 툴바 버튼 정보를 읽지 못했습니다.")
            button = _ToolbarButton()
            bytes_read = ctypes.c_size_t()
            if not kernel32.ReadProcessMemory(
                process, remote_buffer, ctypes.byref(button), ctypes.sizeof(button), ctypes.byref(bytes_read),
            ) or bytes_read.value != ctypes.sizeof(button):
                raise RuntimeError("MFC 툴바 버튼 명령을 읽지 못했습니다.")
            commands.append({"index": index, "commandId": int(button.idCommand)})
        return commands
    finally:
        if remote_buffer:
            kernel32.VirtualFreeEx(process, remote_buffer, 0, MEM_RELEASE)
        kernel32.CloseHandle(process)


def oz_print_command_id(commands: list[dict[str, int]]) -> int:
    """현재 OZ 툴바에서 검증된 인쇄 버튼의 명령 ID만 반환한다.

    표준 MFC 인쇄 명령이 있으면 그것을 우선 사용한다. 현재 OZ Viewer는 표준
    ``ID_FILE_PRINT`` 대신 두 번째 툴바 버튼에 자체 명령 ID를 사용하므로, 표준
    명령이 없을 때에만 확인된 두 번째 버튼을 사용한다.
    """
    matches = [
        int(command["commandId"])
        for command in commands
        if int(command.get("commandId", 0)) == MFC_ID_FILE_PRINT
    ]
    if len(matches) == 1:
        return matches[0]

    oz_print_buttons = [
        int(command["commandId"])
        for command in commands
        if int(command.get("index", -1)) == OZ_PRINT_TOOLBAR_BUTTON_INDEX
        and int(command.get("commandId", 0)) > 0
    ]
    return oz_print_buttons[0] if len(oz_print_buttons) == 1 else 0


def print_dialog_windows() -> list[DesktopWindow]:
    """Windows 공통 인쇄 대화상자만 제목이 정확히 일치할 때 반환한다."""
    return [
        window for window in visible_top_level_windows()
        if window.title == PRINT_DIALOG_TITLE
    ]


def click_print_dialog_confirm(dialog_handle: int) -> None:
    """Windows 인쇄 창의 기본 확인 동작을 Enter로 실행한다.

    호출자는 이 동작 전에 사용자에게 실제 프린터 전송 사실을 명시적으로 확인받아야 한다.
    """
    if os.name != "nt":
        raise RuntimeError("Windows 인쇄 창 제어는 Windows에서만 지원합니다.")
    user32 = ctypes.windll.user32
    activate_window(dialog_handle)
    # 이 표준 인쇄 대화상자의 기본 버튼은 확인이다. 키 입력은 창 위치·해상도·DPI
    # 배율 및 다중 모니터 좌표에 의존하지 않는다.
    user32.keybd_event(VK_RETURN, 0, 0, 0)
    user32.keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)


def send_enter_to_foreground_window() -> None:
    """현재 전면의 모달 창에 Enter를 한 번 전송한다.

    일부 프린터 드라이버 대화상자는 다른 프로세스의 최상위 창 열거에 나타나지
    않는다. 이 함수는 그 창의 존재를 추정하지 않고, 포털이 연 모달 인쇄창이
    전면인 명시적 실제 인쇄 단계에서만 사용한다.
    """
    if os.name != "nt":
        raise RuntimeError("Windows 인쇄 창 제어는 Windows에서만 지원합니다.")
    user32 = ctypes.windll.user32
    user32.keybd_event(VK_RETURN, 0, 0, 0)
    user32.keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)


def activate_window(handle: int) -> None:
    """확인된 창을 전면으로 가져와 사용자 입력 대상이 되도록 한다.

    제목 표시줄 이외의 화면 요소는 누르지 않는다.
    """
    if os.name != "nt":
        raise RuntimeError("OZ Viewer 단축키 제어는 Windows에서만 지원합니다.")
    user32 = ctypes.windll.user32
    target = wintypes.HWND(handle)
    deadline = time.monotonic() + WINDOW_ACTIVATE_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        user32.ShowWindow(target, SW_RESTORE)
        user32.BringWindowToTop(target)
        user32.SetForegroundWindow(target)
        if int(user32.GetForegroundWindow()) == handle:
            break
        # 새 OZ Viewer는 열림 직후 Windows의 포커스 정책 때문에 단순 전면 전환이
        # 거부될 수 있다. 확인된 창의 제목 표시줄만 한 번 클릭해 사용자와 같은
        # 방식으로 활성화한다. 인쇄 버튼이나 문서 영역은 클릭하지 않는다.
        rect = wintypes.RECT()
        if user32.GetWindowRect(target, ctypes.byref(rect)):
            user32.SetCursorPos(rect.left + 180, rect.top + 18)
            user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.25)
    else:
        raise RuntimeError("대상 창이 활성화되지 않아 다음 동작을 실행하지 않았습니다.")


def send_ctrl_p_to_window(handle: int) -> None:
    """확인된 창을 전면으로 가져온 뒤 Ctrl+P만 전송한다.

    이 함수는 Windows 인쇄 대화상자의 확인·취소·프린터 선택을 누르지 않는다.
    """
    if os.name != "nt":
        raise RuntimeError("OZ Viewer 단축키 제어는 Windows에서만 지원합니다.")
    user32 = ctypes.windll.user32
    activate_window(handle)
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(VK_P, 0, 0, 0)
    user32.keybd_event(VK_P, 0, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


def click_oz_viewer_print_toolbar_button(handle: int) -> None:
    """확인된 OZ Viewer의 실제 MFC 인쇄 명령만 한 번 호출한다.

    화면 마우스 입력이 아닌, 실행 중인 Viewer 툴바에서 읽어 검증한 인쇄 명령을
    부모 창으로 보낸다. 멀티 모니터, 창 위치, 포커스 상태의 영향을 받지 않는다.
    """
    if os.name != "nt":
        raise RuntimeError("OZ Viewer 툴바 제어는 Windows에서만 지원합니다.")
    user32 = ctypes.windll.user32
    toolbar_handle = oz_toolbar_handle_from_diagnostics(wait_for_oz_toolbar_diagnostics(handle))
    if not toolbar_handle:
        raise RuntimeError("OZ Report Viewer의 상단 MFC 툴바를 하나로 식별하지 못했습니다.")
    command_id = oz_print_command_id(toolbar_button_commands(toolbar_handle))
    if not command_id:
        raise RuntimeError("OZ Viewer 툴바에서 확인된 인쇄 명령을 하나로 찾지 못했습니다.")
    user32.SendMessageW(
        wintypes.HWND(handle), WM_COMMAND, command_id, wintypes.HWND(toolbar_handle),
    )


def click_oz_viewer_print_toolbar_icon(handle: int) -> None:
    """준비된 OZ Viewer의 프린터 아이콘을 화면 좌표로 한 번 클릭한다.

    MFC 명령 메시지가 무시되는 환경에서만 사용하는 보조 경로다. 창이나 모니터의
    절대 좌표를 고정하지 않고 실제 툴바의 클라이언트 좌표를 화면 좌표로 변환한다.
    """
    if os.name != "nt":
        raise RuntimeError("OZ Viewer 툴바 제어는 Windows에서만 지원합니다.")
    enable_per_monitor_dpi_awareness()
    user32 = ctypes.windll.user32
    toolbar_handle = oz_toolbar_handle_from_diagnostics(wait_for_oz_toolbar_diagnostics(handle))
    if not toolbar_handle:
        raise RuntimeError("OZ Report Viewer의 상단 MFC 툴바를 하나로 식별하지 못했습니다.")
    activate_window(handle)
    point = wintypes.POINT(OZ_PRINT_TOOLBAR_CLIENT_X, OZ_PRINT_TOOLBAR_CLIENT_Y)
    if not user32.ClientToScreen(wintypes.HWND(toolbar_handle), ctypes.byref(point)):
        raise RuntimeError("OZ Viewer 프린터 아이콘의 화면 위치를 읽지 못했습니다.")
    user32.SetCursorPos(point.x, point.y)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def wait_for_new_single_print_dialog(
    existing_handles: set[int], timeout_seconds: int,
) -> DesktopWindow:
    """Ctrl+P 뒤 새 Windows 인쇄 대화상자 한 개가 열릴 때만 반환한다."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        new_windows = [
            window for window in print_dialog_windows()
            if window.handle not in existing_handles
        ]
        if len(new_windows) == 1:
            return new_windows[0]
        if len(new_windows) > 1:
            raise RuntimeError("새 Windows 인쇄 창이 여러 개 열렸습니다. 확인을 누르지 않고 중단합니다.")
        time.sleep(0.5)
    raise RuntimeError("Windows 인쇄 창이 열렸는지 확인하지 못했습니다.")


def wait_for_new_single_oz_viewer(
    existing_handles: set[int], timeout_seconds: int,
) -> DesktopWindow:
    """인쇄 요청 뒤 새로 열린 OZ Viewer 한 개를 기다린다."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        new_windows = [
            window for window in oz_viewer_windows()
            if window.handle not in existing_handles
        ]
        if len(new_windows) == 1:
            return new_windows[0]
        if len(new_windows) > 1:
            raise RuntimeError("새 OZ Report Viewer 창이 여러 개 열렸습니다. 인쇄하지 않고 중단합니다.")
        time.sleep(0.5)
    raise RuntimeError("OZ Report Viewer가 새로 열렸는지 확인하지 못했습니다.")


def wait_for_window_close(handle: int, timeout_seconds: int) -> bool:
    """사용자가 검토 후 창을 닫을 때까지 버튼 조작 없이 대기한다."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if all(window.handle != handle for window in visible_top_level_windows()):
            return True
        time.sleep(0.5)
    return False
