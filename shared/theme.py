"""앱 전체 테마: 디자인 기준값과 그 값으로 만든 Fusion 팔레트·스타일시트(docs/ui-design-guide.md).

규칙
- 색·간격·모서리·글꼴 크기 값은 이 파일에서만 정한다. 화면 코드는 이 이름이나 동적 속성
  (``role``·``status``·``notice``·``band``·``chip``·``emphasis`` 등)을 쓴다.
- 위젯 자체의 ``setStyleSheet``로 색을 넣지 않는다(앱 스타일시트보다 우선해 테마가 덮인다).
- ``apply_theme(app)``은 ``QApplication``을 만든 직후 한 번 부른다. ``main()``과 측정·캡처 도구가 같은
  함수를 불러 같은 모양으로 잰다.
"""

# ── 색 ─────────────────────────────────────────────────────────────
BG_APP = "#f4f6f8"
BG_SURFACE = "#ffffff"
BG_SUBTLE = "#f8f9fb"
BG_PRESSED = "#eceff3"
BORDER = "#dde2e8"
BORDER_STRONG = "#c3cad3"
TEXT = "#1f2933"
TEXT_SECONDARY = "#52606d"
TEXT_MUTED = "#616e7c"
TEXT_PLACEHOLDER = "#8a94a3"
PRIMARY = "#1a73e8"
PRIMARY_HOVER = "#1765cc"
PRIMARY_PRESSED = "#155ab6"
PRIMARY_DISABLED = "#9fbff0"
PRIMARY_SUBTLE = "#e8f0fe"
DANGER = "#c0392b"
DANGER_SUBTLE = "#fdecea"
DANGER_BORDER = "#e6b0aa"
ON_DANGER = "#ffffff"  # danger 바탕 위 글자(측면 툴바 숫자 배지)
OK = "#2e7d32"
OK_SUBTLE = "#edf7ee"
OK_BORDER = "#a5d6a7"
WARN = "#a15f00"  # 흰 바탕 대비 5.1(이전 #b26a00은 4.2로 기준 4.5 미달)
WARN_SUBTLE = "#fff4e5"
WARN_BORDER = "#f0c27a"
IDLE = "#888888"
DISABLED_TEXT = "#a0a8b3"
DISABLED_BG = "#eef0f3"
TOOLTIP_BG = "#1f2933"
TOOLTIP_TEXT = "#ffffff"
INFO = "#075985"
INFO_SUBTLE = "#e0f2fe"
INFO_BORDER = "#7dd3fc"
# 화면을 덮는 오버레이 바탕(반투명)
SCRIM = "rgba(31, 41, 51, 0.48)"
DROP_OVERLAY_BG = "rgba(232, 240, 254, 0.88)"
LOADING_VEIL = "rgba(255, 255, 255, 0.69)"
# 데이터 색: 표 행 바탕(주문 미리보기·배송추적 위험), 스토어 글자 배지
ROW_DANGER_BG = "#fde2e1"
ROW_WARN_BG = "#fff3c4"
TRACK_RISK_HIGH_BG = "#ffb3b3"
TRACK_RISK_WARN_BG = "#ffe0b3"
STORE_11ST = "#f43142"
SEARCH_HIGHLIGHT_BG = "#fff59d"  # 편집기 찾기 결과 강조

# ── 글꼴(포인트) ───────────────────────────────────────────────────
FONT_FAMILIES = ("Malgun Gothic", "맑은 고딕")
FONT_SIZE = 10.0  # 본문(계획 2-3에서 9pt→10pt, 크기 예산 재측정)
FONT_SIZE_SMALL = 9.0

# ── 간격·크기(논리 픽셀) ───────────────────────────────────────────
SPACE_XS = 4
SPACE_SM = 8
SPACE_MD = 12
SPACE_LG = 16
SPACE_XL = 24
CONTROL_HEIGHT = 30
CONTROL_HEIGHT_COMPACT = 26
CHIP_HEIGHT = 22
BUTTON_MIN_WIDTH = 64
RADIUS_CONTROL = 6
RADIUS_CARD = 8
RADIUS_CHIP = 11
FOCUS_RING = 2

# 아이콘(스타일시트 url은 절대 경로를 슬래시로 적는다)
from shared.paths import APP_ROOT as _APP_ROOT

ICON_DIR = (_APP_ROOT / "image" / "ui").as_posix()
ICON_CHEVRON_DOWN = f"{ICON_DIR}/chevron-down.svg"
ICON_CHEVRON_UP = f"{ICON_DIR}/chevron-up.svg"

# 컨트롤 안쪽 여백: 위아래는 높이에서 글자 높이(약 16)와 테두리를 뺀 값의 절반
_PAD_V = max(2, (CONTROL_HEIGHT - 16 - 2) // 2)
_PAD_H = 12


def build_stylesheet() -> str:
    """기준값으로 앱 전체 스타일시트를 만든다."""
    focus_pad_v = _PAD_V - (FOCUS_RING - 1)
    focus_pad_h = _PAD_H - (FOCUS_RING - 1)
    return f"""
/* ── 바탕 ───────────────────────────────────────────── */
QMainWindow, QDialog {{ background: {BG_APP}; }}
QToolTip {{
    background: {TOOLTIP_BG}; color: {TOOLTIP_TEXT}; border: none;
    padding: {SPACE_XS}px {SPACE_SM}px; border-radius: 4px;
}}

/* ── 메뉴 ───────────────────────────────────────────── */
QMenuBar {{ background: {BG_APP}; border-bottom: 1px solid {BORDER}; padding: 2px 4px; }}
QMenuBar::item {{ padding: 4px 10px; border-radius: 4px; background: transparent; }}
QMenuBar::item:selected {{ background: {BG_PRESSED}; }}
QMenu {{ background: {BG_SURFACE}; border: 1px solid {BORDER}; padding: 4px; }}
QMenu::item {{ padding: 6px 24px 6px 12px; border-radius: 4px; }}
QMenu::item:selected {{ background: {PRIMARY_SUBTLE}; color: {TEXT}; }}
QMenu::item:disabled {{ color: {DISABLED_TEXT}; }}
QMenu::separator {{ height: 1px; background: {BORDER}; margin: 4px 8px; }}

/* ── 탭 ─────────────────────────────────────────────── */
QTabWidget::pane {{ border: none; border-top: 1px solid {BORDER}; background: {BG_APP}; top: -1px; }}
QTabBar::tab {{
    background: transparent; color: {TEXT_SECONDARY};
    padding: 8px 14px; margin-right: 2px; border: none; border-bottom: 2px solid transparent;
}}
QTabBar::tab:hover {{ color: {TEXT}; background: {BG_PRESSED}; }}
QTabBar::tab:selected {{ color: {TEXT}; border-bottom: 2px solid {PRIMARY}; }}

/* ── 측면 툴바(메인 탭 이동) ────────────────────────────── */
/* 메인 탭은 이름 줄을 숨기고 왼쪽 툴바로 고른다. 탭 이름 줄이 없으니 페이지 위 경계선도 뺀다. */
QTabWidget#tabWidget::pane {{ border: none; top: 0px; }}
QToolBar#navigationRail {{
    background: {BG_PRESSED}; border: none; border-right: 1px solid {BORDER};
    padding: {SPACE_SM}px {SPACE_XS}px; spacing: {SPACE_XS}px;
}}
QToolButton#navigationButton {{
    min-width: 56px; max-width: 56px; padding: 6px 2px; border: none; border-radius: {RADIUS_CARD}px;
    background: transparent; color: {TEXT_SECONDARY};
}}
QToolButton#navigationButton:hover {{ background: {BORDER}; color: {TEXT}; }}
QToolButton#navigationButton:checked {{ background: {PRIMARY_SUBTLE}; color: {PRIMARY_PRESSED}; font-weight: bold; }}

/* ── 묶음(카드) ─────────────────────────────────────── */
/* 제목을 카드 안 왼쪽 위에 굵게 둔다(공통 부품 SectionCard와 같은 모양, 계획 4-3). */
QGroupBox {{
    background: {BG_SURFACE}; border: 1px solid {BORDER}; border-radius: {RADIUS_CARD}px;
    margin-top: 0px; padding: {SPACE_MD + 22}px {SPACE_SM}px {SPACE_SM}px {SPACE_SM}px;
    font-weight: bold; font-size: {FONT_SIZE + 0.5:g}pt;
}}
QGroupBox::title {{
    subcontrol-origin: padding; subcontrol-position: top left;
    left: {SPACE_MD}px; top: {SPACE_MD}px; padding: 0; color: {TEXT};
}}

/* ── 버튼(보조 기본값) ──────────────────────────────── */
QPushButton {{
    background: {BG_SURFACE}; color: {TEXT};
    border: 1px solid {BORDER_STRONG}; border-radius: {RADIUS_CONTROL}px;
    padding: {_PAD_V}px {_PAD_H}px; min-width: {BUTTON_MIN_WIDTH - 2 * _PAD_H}px;
}}
QPushButton:hover {{ background: {BG_SUBTLE}; border-color: {TEXT_PLACEHOLDER}; }}
QPushButton:pressed {{ background: {BG_PRESSED}; }}
QPushButton:checked {{ background: {PRIMARY_SUBTLE}; border-color: {PRIMARY}; }}
QPushButton:disabled {{ color: {DISABLED_TEXT}; background: {DISABLED_BG}; border-color: {BORDER}; }}
QPushButton:focus {{ border: {FOCUS_RING}px solid {PRIMARY}; padding: {focus_pad_v}px {focus_pad_h}px; }}
QToolButton {{ border: 1px solid transparent; border-radius: 4px; padding: 2px 6px; background: transparent; }}
QToolButton:hover {{ background: {BG_PRESSED}; }}

QPushButton[role="primary"] {{
    background: {PRIMARY}; color: white; font-weight: bold; border: 1px solid {PRIMARY_HOVER};
}}
QPushButton[role="primary"]:hover {{ background: {PRIMARY_HOVER}; }}
QPushButton[role="primary"]:pressed {{ background: {PRIMARY_PRESSED}; }}
QPushButton[role="primary"]:focus {{ border: {FOCUS_RING}px solid {PRIMARY_PRESSED}; }}
QPushButton[role="primary"]:disabled {{
    background: {PRIMARY_DISABLED}; border-color: {PRIMARY_DISABLED}; color: white;
}}
QPushButton[role="danger"] {{ color: {DANGER}; background: {BG_SURFACE}; border: 1px solid {DANGER_BORDER}; }}
QPushButton[role="danger"]:hover {{ background: {DANGER_SUBTLE}; }}
QPushButton[role="danger"]:pressed {{ background: #f9d6d2; }}
QPushButton[role="danger"]:checked {{ background: #f9d6d2; border-color: {DANGER}; font-weight: bold; }}
QPushButton[role="danger"]:focus {{ border: {FOCUS_RING}px solid {DANGER}; }}
QPushButton[role="danger"]:disabled {{ color: #d9a29b; border-color: #f0d4d0; background: {BG_SURFACE}; }}
QPushButton[chip="true"] {{ padding: 1px 8px; min-width: 0px; }}

/* ── 입력 ───────────────────────────────────────────── */
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit, QTimeEdit, QDateTimeEdit {{
    background: {BG_SURFACE}; color: {TEXT};
    border: 1px solid {BORDER_STRONG}; border-radius: {RADIUS_CONTROL}px;
    padding: {_PAD_V - 1}px {SPACE_SM}px; min-height: 18px;
    selection-background-color: {PRIMARY}; selection-color: white;
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus,
QTimeEdit:focus, QDateTimeEdit:focus {{ border: 1px solid {PRIMARY}; }}
QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QDateEdit:disabled {{
    background: {DISABLED_BG}; color: {DISABLED_TEXT}; border-color: {BORDER};
}}
QLineEdit:read-only {{ background: {BG_SUBTLE}; }}
QComboBox::drop-down, QDateEdit::drop-down {{
    subcontrol-origin: padding; subcontrol-position: top right; border: none; width: 22px;
}}
QComboBox::down-arrow, QDateEdit::down-arrow {{ image: url({ICON_CHEVRON_DOWN}); width: 12px; height: 12px; }}
QComboBox::down-arrow:disabled, QDateEdit::down-arrow:disabled {{ image: none; }}
QSpinBox, QDoubleSpinBox {{ padding-right: 20px; }}
QSpinBox::up-button, QDoubleSpinBox::up-button {{
    subcontrol-origin: border; subcontrol-position: top right; width: 18px;
    border: none; border-left: 1px solid {BORDER}; border-top-right-radius: {RADIUS_CONTROL}px;
}}
QSpinBox::down-button, QDoubleSpinBox::down-button {{
    subcontrol-origin: border; subcontrol-position: bottom right; width: 18px;
    border: none; border-left: 1px solid {BORDER}; border-bottom-right-radius: {RADIUS_CONTROL}px;
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover,
QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover {{ background: {BG_PRESSED}; }}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{ image: url({ICON_CHEVRON_UP}); width: 10px; height: 10px; }}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{ image: url({ICON_CHEVRON_DOWN}); width: 10px; height: 10px; }}
QComboBox QAbstractItemView {{
    background: {BG_SURFACE}; border: 1px solid {BORDER}; selection-background-color: {PRIMARY_SUBTLE};
    selection-color: {TEXT}; outline: none;
}}
QPlainTextEdit, QTextEdit, QTextBrowser {{
    background: {BG_SURFACE}; color: {TEXT}; border: 1px solid {BORDER}; border-radius: {RADIUS_CONTROL}px;
    selection-background-color: {PRIMARY}; selection-color: white;
}}
QPlainTextEdit:focus, QTextEdit:focus {{ border: 1px solid {PRIMARY}; }}

/* ── 표·목록 ────────────────────────────────────────── */
QTableView, QTableWidget, QTreeView, QListView, QListWidget {{
    background: {BG_SURFACE}; color: {TEXT}; border: 1px solid {BORDER}; border-radius: {RADIUS_CONTROL}px;
    gridline-color: {BORDER}; selection-background-color: {PRIMARY_SUBTLE}; selection-color: {TEXT};
    outline: none;
}}
QHeaderView::section {{
    background: {BG_SUBTLE}; color: {TEXT_SECONDARY}; font-weight: bold;
    border: none; border-bottom: 1px solid {BORDER}; border-right: 1px solid {BORDER};
    padding: 4px {SPACE_SM}px;
}}
QHeaderView::section:last {{ border-right: none; }}
QTableCornerButton::section {{ background: {BG_SUBTLE}; border: none; border-bottom: 1px solid {BORDER}; }}

/* ── 기타 ───────────────────────────────────────────── */
QStatusBar {{ background: {BG_APP}; border-top: 1px solid {BORDER}; color: {TEXT_SECONDARY}; }}
QStatusBar::item {{ border: none; }}
QProgressBar {{
    background: {BG_PRESSED}; border: none; border-radius: 4px; text-align: center; color: {TEXT};
}}
QProgressBar::chunk {{ background: {PRIMARY}; border-radius: 4px; }}
QSplitter::handle {{ background: {BG_APP}; }}
QScrollArea {{ border: none; }}

/* ── 공통 부품(shared/ui_components.py) ───────────────── */
QFrame[card="true"], QWidget[card="true"] {{
    background: {BG_SURFACE}; border: 1px solid {BORDER}; border-radius: {RADIUS_CARD}px;
}}
QLabel[sectionTitle="true"], QToolButton[sectionTitle="true"] {{
    color: {TEXT}; font-weight: bold; font-size: {FONT_SIZE + 0.5:g}pt;
}}
QLabel[stepBadge="true"] {{
    background: {PRIMARY_SUBTLE}; color: {PRIMARY}; border-radius: 11px; font-weight: bold;
}}
QLabel[stepBadge="true"][stepState="done"] {{ background: {OK_SUBTLE}; color: {OK}; }}
QLabel[formLabel="true"] {{ color: {TEXT_SECONDARY}; }}
QLabel[formLabel="true"][status="warn"] {{ color: {WARN}; }}
QLabel[fieldError="true"] {{ color: {DANGER}; font-size: {FONT_SIZE_SMALL:g}pt; }}
QLabel[emptyText="true"] {{ color: {TEXT_MUTED}; }}
QWidget[emptyPanel="true"] {{
    background: {BG_SUBTLE}; border: 1px dashed {BORDER_STRONG}; border-radius: {RADIUS_CONTROL}px;
}}
QTableView[dataTable="true"]::item {{ border-bottom: 1px solid {BORDER}; padding: 0 {SPACE_SM}px; }}
QTableView[dataTable="true"]::item:selected {{ background: {PRIMARY_SUBTLE}; color: {TEXT}; }}

/* ── 설정 탭 분류 목록(사이드바) ───────────────────────── */
QListWidget#listWidget_settings_nav {{
    background: transparent; border: none; border-right: 1px solid {BORDER};
    padding: 0 {SPACE_SM}px 0 0; outline: none;
}}
QListWidget#listWidget_settings_nav::item {{
    color: {TEXT_SECONDARY}; padding: {SPACE_SM}px {SPACE_MD}px; margin: 1px 0; border-radius: {RADIUS_CONTROL}px;
}}
QListWidget#listWidget_settings_nav::item:hover {{ background: {BG_PRESSED}; }}
QListWidget#listWidget_settings_nav::item:selected {{ background: {PRIMARY_SUBTLE}; color: {PRIMARY}; }}

/* ── 상세 HTML 편집기 ───────────────────────────────── */
QFrame#detailPreviewCanvas {{ background: {BG_PRESSED}; }}
QPushButton[swatch="true"] {{
    border: 2px solid transparent; border-radius: 3px; padding: 0; min-width: 0px;
}}
QPushButton[swatch="true"]:hover, QPushButton[swatch="true"]:focus {{ border-color: {TEXT_PLACEHOLDER}; }}
QPushButton[swatch="true"]:checked {{ border-color: {TEXT}; }}
QPushButton[swatch="true"]:disabled {{ border-color: {BORDER}; }}

/* ── 오버레이(시작·작업 중, 배송추적 목록 전환, 파일 끌어다 놓기) ── */
QFrame#startupLoadingOverlay, QFrame#busyProcessingOverlay {{ background: {SCRIM}; border: none; }}
QFrame#startupLoadingCard, QFrame#busyProcessingCard {{
    background: {BG_SURFACE}; border: 1px solid {BORDER}; border-radius: 12px;
}}
QLabel[overlayText="badge"] {{ color: {PRIMARY}; font-weight: bold; font-size: 11pt; }}
QLabel[overlayText="title"] {{ color: {TEXT}; font-weight: bold; font-size: 14pt; }}
QLabel[overlayText="sub"] {{ color: {TEXT_SECONDARY}; font-size: 10pt; }}
QLabel[overlayText="hint"] {{ color: {TEXT_MUTED}; }}
QProgressBar#startupProgressBar {{
    background: {PRIMARY_SUBTLE}; border: none; border-radius: 6px; color: {TEXT}; font-weight: bold;
}}
QProgressBar#startupProgressBar::chunk {{ background: {PRIMARY}; border-radius: 6px; }}
QWidget#trackingLoadingOverlay {{ background: {LOADING_VEIL}; }}
QLabel#trackingLoadingMessage {{
    background: {BG_SURFACE}; border: 1px solid {BORDER}; border-radius: {RADIUS_CARD}px;
    color: {TEXT}; font-weight: bold;
}}
QLabel[dropOverlay="true"] {{
    background: {DROP_OVERLAY_BG}; color: {PRIMARY}; font-weight: bold;
    border: 2px dashed {PRIMARY}; border-radius: {RADIUS_CONTROL}px;
}}

/* ── 주문·발송 파일 칸(끌어다 놓는 곳) ───────────────── */
QLabel[dropZone="true"] {{
    color: {IDLE}; border: 1px dashed {BORDER_STRONG}; border-radius: {RADIUS_CONTROL}px;
}}
QLabel[dropZone="true"][fileState="ok"] {{ color: {OK}; font-weight: bold; }}

/* ── 상태바 ─────────────────────────────────────────── */
QStatusBar {{ min-height: 30px; }}
QLabel#statusMessageLabel {{ color: {TEXT_SECONDARY}; padding: 0 {SPACE_SM}px; }}

/* ── 배지 ───────────────────────────────────────────── */
QLabel[badge="info"] {{
    color: {INFO}; background: {INFO_SUBTLE}; border: 1px solid {INFO_BORDER};
    border-radius: {RADIUS_CARD}px; padding: 2px {SPACE_SM}px; font-weight: bold;
}}

/* ── 상태 표시(shared/ui_style.set_status·set_band, shared/ui_notice) ── */
QLabel[status="ok"] {{ color: {OK}; }}
QLabel[status="warn"] {{ color: {WARN}; }}
QLabel[status="error"] {{ color: {DANGER}; }}
QLabel[status="idle"] {{ color: {IDLE}; }}
QLabel[emphasis="true"] {{ font-weight: bold; }}
QLabel[hint="true"] {{ color: {TEXT_MUTED}; font-size: {FONT_SIZE_SMALL:g}pt; }}
QFrame[notice="ok"] {{ background: {OK_SUBTLE}; border-top: 1px solid {OK_BORDER}; }}
QFrame[notice="warn"] {{ background: {WARN_SUBTLE}; border-top: 1px solid {WARN_BORDER}; }}
QFrame[notice="error"] {{ background: {DANGER_SUBTLE}; border-top: 1px solid {DANGER_BORDER}; }}
QLabel[band="warn"] {{
    color: {WARN}; background: {WARN_SUBTLE}; font-weight: bold;
    border: 1px solid {WARN_BORDER}; border-radius: 4px;
}}
"""


APP_STYLESHEET = build_stylesheet()


def build_palette():
    """Fusion 스타일이 스타일시트 밖에서 그리는 부분(체크박스, 스크롤바, 선택 등)에 쓰는 팔레트."""
    from PySide6.QtGui import QColor, QPalette

    palette = QPalette()
    roles = {
        QPalette.ColorRole.Window: BG_APP,
        QPalette.ColorRole.WindowText: TEXT,
        QPalette.ColorRole.Base: BG_SURFACE,
        QPalette.ColorRole.AlternateBase: BG_SUBTLE,
        QPalette.ColorRole.Text: TEXT,
        QPalette.ColorRole.Button: BG_SURFACE,
        QPalette.ColorRole.ButtonText: TEXT,
        QPalette.ColorRole.Highlight: PRIMARY,
        QPalette.ColorRole.HighlightedText: "#ffffff",
        QPalette.ColorRole.ToolTipBase: TOOLTIP_BG,
        QPalette.ColorRole.ToolTipText: TOOLTIP_TEXT,
        QPalette.ColorRole.PlaceholderText: TEXT_PLACEHOLDER,
        QPalette.ColorRole.Link: PRIMARY,
        QPalette.ColorRole.Light: BG_SURFACE,
        QPalette.ColorRole.Midlight: BG_SUBTLE,
        QPalette.ColorRole.Mid: BORDER_STRONG,
        QPalette.ColorRole.Dark: BORDER_STRONG,
        QPalette.ColorRole.Shadow: TEXT_MUTED,
    }
    for role, color in roles.items():
        palette.setColor(role, QColor(color))
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(DISABLED_TEXT))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Base, QColor(DISABLED_BG))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Button, QColor(DISABLED_BG))
    return palette


def app_font(base_font):
    """플랫폼 기본 글꼴에서 글꼴 이름과 크기만 테마 값으로 바꾼다(없는 글꼴이면 Qt가 대체한다)."""
    font = type(base_font)(base_font)
    font.setFamilies(list(FONT_FAMILIES))
    font.setPointSizeF(FONT_SIZE)
    return font


def apply_theme(app) -> None:
    """QApplication 생성 직후 한 번 부른다: Fusion 스타일, 팔레트, 글꼴, 앱 전체 스타일시트."""
    app.setStyle("Fusion")
    app.setPalette(build_palette())
    app.setFont(app_font(app.font()))
    app.setStyleSheet(APP_STYLESHEET)


def contrast_ratio(foreground: str, background: str) -> float:
    """WCAG 대비율(1~21)."""

    def luminance(hex_color):
        hex_color = hex_color.lstrip("#")
        channels = [int(hex_color[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        linear = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in channels]
        return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]

    lighter, darker = sorted((luminance(foreground), luminance(background)), reverse=True)
    return (lighter + 0.05) / (darker + 0.05)
