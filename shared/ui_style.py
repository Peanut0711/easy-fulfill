"""앱 공통 버튼 역할과 상태 표시 도우미.

색과 스타일시트는 ``shared/theme.py``(디자인 기준값) 한 곳에서 정한다. 이 모듈은 위젯에 동적 속성
(``role``·``status``·``emphasis``·``band``)을 붙이고 다시 그리게 하는 도우미만 둔다.
- 역할·상태를 바꿀 때는 위젯 자체의 ``setStyleSheet``로 색을 넣지 않는다(위젯 스타일시트가 앱
  스타일시트보다 우선해 색이 덮인다).
"""

from shared import theme

COLOR_PRIMARY = theme.PRIMARY
COLOR_PRIMARY_HOVER = theme.PRIMARY_HOVER
COLOR_PRIMARY_PRESSED = theme.PRIMARY_PRESSED
COLOR_PRIMARY_DISABLED = theme.PRIMARY_DISABLED
COLOR_DANGER = theme.DANGER
COLOR_OK = theme.OK
COLOR_WARN = theme.WARN
COLOR_IDLE = theme.IDLE

BUTTON_ROLES = ("primary", "danger")
STATUS_KINDS = ("ok", "warn", "error", "idle")
STATUS_MARKS = {"ok": "✓", "warn": "!", "error": "✗", "idle": ""}

APP_STYLESHEET = theme.APP_STYLESHEET


def apply_app_style(app) -> None:
    """QApplication 생성 직후 한 번 부른다(테마 전체 적용: ``shared.theme.apply_theme``)."""
    theme.apply_theme(app)


def _repolish(widget) -> None:
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.updateGeometry()  # 역할(굵은 글씨·여백)이 바뀌면 크기도 바뀌므로 레이아웃이 다시 계산하게 한다
    widget.update()


def set_button_role(button, role) -> None:
    """``role``: ``"primary"``(그 화면의 다음 동작), ``"danger"``(되돌리기 어려운 동작), ``None``(기본)."""
    if role is not None and role not in BUTTON_ROLES:
        raise ValueError(f"알 수 없는 버튼 역할: {role}")
    if button.property("role") == role or (role is None and not button.property("role")):
        return
    button.setProperty("role", role or "")
    _repolish(button)


def first_line(text) -> str:
    return str(text or "").strip().splitlines()[0] if str(text or "").strip() else ""


def set_status(label, kind, text, detail="", *, emphasis=False, mark=True, single_line=True) -> None:
    """상태 문구를 색과 기호(✓ ! ✗)로 보인다. 여러 줄이면 첫 줄만 보이고 원문은 툴팁에 둔다.

    ``detail``이 있으면 툴팁은 ``detail``이다. ``mark=False``면 기호를 붙이지 않는다.
    ``single_line=False``면 여러 줄을 그대로 보인다.
    """
    if kind not in STATUS_KINDS:
        raise ValueError(f"알 수 없는 상태: {kind}")
    shown = first_line(text) if single_line else str(text or "")
    symbol = STATUS_MARKS[kind] if mark else ""
    if symbol and not shown.startswith(symbol):
        shown = f"{symbol} {shown}" if shown else symbol
    multiline = single_line and "\n" in str(text or "").strip()
    tooltip = detail or (str(text).strip() if multiline else "")
    changed = label.property("status") != kind or bool(label.property("emphasis")) != emphasis
    label.setProperty("status", kind)
    label.setProperty("emphasis", "true" if emphasis else "")
    label.setText(shown)
    label.setToolTip(tooltip)
    if changed:
        _repolish(label)


BAND_KINDS = ("warn",)
BAND_MARGINS = (8, 4, 8, 4)


def set_band(label, kind) -> None:
    """``kind``가 있으면 라벨을 색 띠(배경·테두리)로 보이고, ``None``이면 띠를 숨긴다."""
    if kind is not None and kind not in BAND_KINDS:
        raise ValueError(f"알 수 없는 띠: {kind}")
    if (label.property("band") or None) != kind:
        label.setProperty("band", kind or "")
        _repolish(label)
    # 스타일시트 padding은 처음 polish 때만 반영되어 속성을 나중에 바꾸면 적용되지 않는다. 여백은 직접 준다.
    if kind is not None:
        label.setContentsMargins(*BAND_MARGINS)
    else:
        label.setContentsMargins(0, 0, 0, 0)
    label.setVisible(kind is not None)


def clear_status(label) -> None:
    """상태 색을 지우고 기본 글자색으로 돌린다(글자는 그대로)."""
    if label.property("status") or label.property("emphasis"):
        label.setProperty("status", "")
        label.setProperty("emphasis", "")
        _repolish(label)
