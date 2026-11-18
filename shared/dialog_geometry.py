"""팝업 창(상세 HTML 편집기·이미지 보기) 위치·크기를 화면에 맞추고 이 PC에 기억한다(알림·상태바 계획 E-2).

메인 창과 같은 규칙(shared/window_geometry.py)을 쓴다. 설정은 창을 띄운 메인 창의 앱 설정(이 PC)에만 저장한다.
"""

from PySide6.QtCore import QRect, Qt
from PySide6.QtGui import QGuiApplication

from shared.window_geometry import (
    centered_position, dialog_default_size, geometry_to_setting, is_restorable, parse_saved_geometry,
)

DETAIL_EDITOR_GEOMETRY_KEY = "detail_html_editor_geometry"
IMAGE_DIALOG_GEOMETRY_KEY = "image_dialog_geometry"
DIALOG_RESTORE_MINIMUM = (320, 240)


def _settings_owner(dialog):
    owner = dialog.parent()
    if owner is not None and hasattr(owner, "get_app_setting") and hasattr(owner, "set_app_setting"):
        return owner
    return None


def apply_dialog_geometry(dialog, setting_key, base_size):
    """저장한 위치·크기가 지금 연결된 화면에 맞으면 되살리고, 아니면 원래 크기(``base_size``)를 넘지 않고
    작업 영역 90% 안에 드는 크기로 화면 가운데에 둔다."""
    owner = _settings_owner(dialog)
    saved = parse_saved_geometry(owner.get_app_setting(setting_key)) if owner is not None else None
    if saved is not None:
        rect, maximized = saved
        areas = [(a.x(), a.y(), a.width(), a.height())
                 for a in (item.availableGeometry() for item in QGuiApplication.screens())]
        if is_restorable(rect, areas, DIALOG_RESTORE_MINIMUM):
            dialog.setGeometry(QRect(*rect))
            if maximized:
                dialog.setWindowState(dialog.windowState() | Qt.WindowState.WindowMaximized)
            return
    screen = (owner.screen() if owner is not None else None) or QGuiApplication.primaryScreen()
    if screen is None:
        dialog.resize(*base_size)
        return
    area = screen.availableGeometry()
    width, height = dialog_default_size((area.width(), area.height()), base_size)
    x, y = centered_position((area.x(), area.y(), area.width(), area.height()), (width, height))
    dialog.setGeometry(QRect(x, y, width, height))


def save_dialog_geometry(dialog, setting_key):
    """닫을 때 위치·크기·최대화를 이 PC 앱 설정에 저장한다(공유 시트에는 쓰지 않음)."""
    owner = _settings_owner(dialog)
    if owner is None:
        return
    maximized = bool(dialog.windowState() & Qt.WindowState.WindowMaximized)
    rect = dialog.normalGeometry() if maximized else dialog.geometry()
    if rect.width() <= 0 or rect.height() <= 0:
        return
    owner.set_app_setting(setting_key, geometry_to_setting((rect.x(), rect.y(), rect.width(), rect.height()), maximized))
