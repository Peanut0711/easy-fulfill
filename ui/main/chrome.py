"""메인 창 공통 틀: 메뉴바, 상태바·상태 칩, 알림 띠, 시작·작업 중 오버레이, 탭 순서, 단축키·도움말."""

import time
from pathlib import Path

from PySide6.QtCore import QEasingCurve, QPropertyAnimation, QRectF, QSize, Qt, QTimer, QUrl
from PySide6.QtGui import (
    QAction,
    QActionGroup,
    QColor,
    QDesktopServices,
    QFont,
    QIcon,
    QKeySequence,
    QPainter,
    QPen,
    QPixmap,
    QShortcut,
)
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFrame,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QLabel,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from shared import theme
from shared.ui_notice import StatusChip
from shared.ui_style import set_status, STATUS_MARKS
from sheets.version_gate import APP_VERSION

# 시작 시 DB동기화 로딩 바: 약 2초+여유 안에 99%까지 선형 증가, 완료 시 즉시 100%
STARTUP_SYNC_PROGRESS_CAP_MS = 2200
STARTUP_SYNC_PROGRESS_TICK_MS = 40
# 100% 도달 후 오버레이를 닫기까지 대기(너무 짧으면 급하게 사라진 느낌, 너무 길면 지루)
STARTUP_SYNC_PROGRESS_HIDE_DELAY_MS = 500


# 왼쪽 측면 툴바(탭 이름 줄 대신): (탭 objectName, 짧은 이름, 아이콘 모양). 설정은 맨 아래에 따로 둔다.
NAV_RAIL_ITEMS = (
    ("tab", "주문", '<path d="M3 7.5 12 3l9 4.5v9L12 21l-9-4.5z"/><path d="M3 7.5 12 12l9-4.5M12 12v9"/>'),
    ("tab_tracking", "추적", '<path d="M2 6h11v10H2zM13 9h4l4 4v3h-8z"/><circle cx="6" cy="17.5" r="2"/>'
                             '<circle cx="17" cy="17.5" r="2"/>'),
    ("tab_quick_invoice", "송장", '<path d="M5 3h14v18l-2.5-1.5L14 21l-2-1.5L10 21l-2.5-1.5L5 21z"/>'
                                  '<path d="M9 8h6M9 12h6"/>'),
    ("tab_documents", "문서", '<path d="M6 3h8l5 5v13H6z"/><path d="M14 3v5h5M9 13h7M9 17h7"/>'),
    ("tab_detail_page", "상세", '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/>'
                                '<path d="m21 16-5-5-9 9"/>'),
    ("tab_settings", "설정", '<path d="M4 6h9M17 6h3M4 12h3M11 12h9M4 18h11M19 18h1"/><circle cx="15" cy="6" r="2"/>'
                             '<circle cx="9" cy="12" r="2"/><circle cx="17" cy="18" r="2"/>'),
)
NAV_RAIL_ICON_SIZE = 24


def nav_rail_icon(shape: str) -> QIcon:
    """선 아이콘(24 격자 SVG). 평소는 보조 글자색, 선택되면 강조색으로 그린다(2배 해상도로 그려 배율에서도 선명)."""
    icon = QIcon()
    for color, state in ((theme.TEXT_SECONDARY, QIcon.State.Off), (theme.PRIMARY_PRESSED, QIcon.State.On)):
        svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><g fill="none" stroke="'
               + color + '" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' + shape + '</g></svg>')
        pixmap = QPixmap(NAV_RAIL_ICON_SIZE * 2, NAV_RAIL_ICON_SIZE * 2)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        QSvgRenderer(svg.encode()).render(painter)
        painter.end()
        pixmap.setDevicePixelRatio(2)
        icon.addPixmap(pixmap, QIcon.Mode.Normal, state)
    return icon


class CircularBusySpinner(QWidget):
    """Google 시트·처리 대기용 무한 회전 링 스피너."""

    def __init__(self, parent=None, size: int = 58, line_width: int = 4, color: str = theme.PRIMARY):
        super().__init__(parent)
        self._line_width = line_width
        self._color = QColor(color)
        self._rotation = 0
        self.setFixedSize(size, size)
        self._timer = QTimer(self)
        self._timer.setInterval(40)
        self._timer.timeout.connect(self._on_tick)

    def _on_tick(self):
        self._rotation = (self._rotation + 24) % 360
        self.update()

    def showEvent(self, event):
        super().showEvent(event)
        self._timer.start()

    def hideEvent(self, event):
        self._timer.stop()
        super().hideEvent(event)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        s = min(self.width(), self.height())
        m = self._line_width + 1
        rf = (s - 2 * m) / 2.0
        p.translate(self.width() / 2.0, self.height() / 2.0)
        p.rotate(self._rotation)
        pen = QPen(self._color)
        pen.setWidth(self._line_width)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(pen)
        rect = QRectF(-rf, -rf, 2 * rf, 2 * rf)
        p.drawArc(rect, 40 * 16, 290 * 16)


class ChromeMixin:
    """MainWindow 믹스인. 메인 창 공통 틀: 메뉴바, 상태바·상태 칩, 알림 띠, 시작·작업 중 오버레이, 탭 순서, 단축키·도움말."""

    def _apply_tab_order(self):
        """합의된 6개 상위 탭 순서를 런타임에서 보장합니다.
        (UI 파일의 정의 순서와 무관하게 런타임에서 보장)"""
        tw = getattr(self.ui, "tabWidget", None)
        if tw is None:
            return
        desired = [
            "tab", "tab_tracking", "tab_quick_invoice", "tab_documents",
            "tab_detail_page", "tab_settings",
        ]
        bar = tw.tabBar()
        for target in range(len(desired)):
            name = desired[target]
            cur = -1
            for i in range(tw.count()):
                if tw.widget(i).objectName() == name:
                    cur = i
                    break
            if cur != -1 and cur != target:
                bar.moveTab(cur, target)
        tw.setCurrentIndex(0)

    def _setup_navigation_rail(self, tabs):
        """탭 이름 줄을 숨기고 왼쪽 측면 툴바로 탭을 고른다.

        탭 화면·전환은 그대로 QTabWidget이 맡는다(단축키·currentChanged·setCurrentWidget 코드가 그대로 동작).
        툴바 버튼은 objectName으로 탭을 찾아 이동하고, 탭이 어떤 경로로 바뀌든 currentChanged로 선택 표시를 맞춘다.
        """
        tabs.tabBar().hide()
        rail = QToolBar("탭", self)
        rail.setObjectName("navigationRail")
        rail.setMovable(False)
        rail.setFloatable(False)
        rail.setAllowedAreas(Qt.ToolBarArea.LeftToolBarArea)
        rail.setOrientation(Qt.Orientation.Vertical)
        rail.toggleViewAction().setVisible(False)  # 오른쪽 클릭 메뉴로 툴바를 숨기지 못하게
        rail.setContextMenuPolicy(Qt.ContextMenuPolicy.PreventContextMenu)
        self.addToolBar(Qt.ToolBarArea.LeftToolBarArea, rail)
        group = QActionGroup(rail)
        group.setExclusive(True)
        self.nav_rail = rail
        self.nav_rail_actions = {}
        pages = {tabs.widget(i).objectName(): tabs.widget(i) for i in range(tabs.count())}
        for name, text, shape in NAV_RAIL_ITEMS:
            page = pages.get(name)
            if page is None:
                continue
            if name == "tab_settings":  # 업무 탭과 떼어 맨 아래에 둔다
                spacer = QWidget(rail)
                spacer.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
                rail.addWidget(spacer)
            action = QAction(nav_rail_icon(shape), text, rail)
            action.setCheckable(True)
            action.setObjectName(f"navAction_{name}")
            action.triggered.connect(lambda _checked=False, p=page: tabs.setCurrentWidget(p))
            group.addAction(action)
            button = QToolButton(rail)
            button.setObjectName("navigationButton")
            button.setDefaultAction(action)
            button.setIconSize(QSize(NAV_RAIL_ICON_SIZE, NAV_RAIL_ICON_SIZE))
            button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            rail.addWidget(button)
            self.nav_rail_actions[name] = action
        tabs.currentChanged.connect(self._sync_navigation_rail)
        self._sync_navigation_rail()

    def _sync_navigation_rail(self, _index=None):
        """현재 탭에 맞춰 측면 툴바 선택 표시와 툴팁(전체 탭 이름·단축키)을 맞춘다."""
        tabs = getattr(self.ui, "tabWidget", None)
        actions = getattr(self, "nav_rail_actions", None)
        if tabs is None or not actions:
            return
        for index in range(tabs.count()):
            page = tabs.widget(index)
            action = actions.get(page.objectName())
            if action is None:
                continue
            title = tabs.tabText(index).replace("⚙", "").strip()
            tip = f"{title} (Ctrl+{index + 1})" if index < 6 else title
            if action.toolTip() != tip:
                action.setToolTip(tip)
            action.setChecked(page is tabs.currentWidget())

    def _setup_status_chips(self):
        """상태바 오른쪽 상시 표시(계획 C-2): Google 연결 · 인덱스 동기화 · 앱 버전. 누르면 해당 설정 페이지.

        값은 이미 있는 상태(인증 확인, 인덱스 동기화, 버전 확인)를 갱신할 때 함께 바꾼다. 새 읽기·호출은 없다.
        """
        bar = self.statusBar()
        chips = (
            ("statusChip_google", "Google", "integration"),
            ("statusChip_index", "인덱스 —", "orders"),
            ("statusChip_version", f"v{APP_VERSION}", "program"),
        )
        for name, text, page in chips:
            chip = StatusChip(name, text, bar)
            set_status(chip, "idle", text)
            chip.clicked.connect(lambda key=page: self.show_settings_page(key))
            bar.addPermanentWidget(chip)
            setattr(self, name, chip)
        self.statusChip_google.setToolTip("Google Sheets 연결 확인 중 · 누르면 「설정 › 연동」")
        self.statusChip_index.setToolTip("주문번호 인덱스 동기화 · 누르면 「설정 › 주문·사은품」")
        self.statusChip_version.setToolTip("앱 버전 · 누르면 「설정 › 프로그램」")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        ov = getattr(self, "_startup_overlay", None)
        if ov is not None and ov.isVisible():
            self._layout_startup_overlay_full()
        bov = getattr(self, "_busy_overlay", None)
        if bov is not None and bov.isVisible():
            self._layout_busy_overlay_full()

    def _layout_startup_overlay_full(self):
        if getattr(self, "_startup_overlay", None) is None:
            return
        self._startup_overlay.setGeometry(
            0, 0, max(1, self.width()), max(1, self.height())
        )

    def _setup_startup_loading_overlay(self):
        # 색·글꼴 크기는 테마(shared/theme.py)의 objectName·overlayText 규칙에서 정한다.
        self._startup_overlay = QFrame(self)
        self._startup_overlay.setObjectName("startupLoadingOverlay")
        self._startup_overlay.hide()
        outer = QVBoxLayout(self._startup_overlay)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addStretch(1)

        row = QHBoxLayout()
        row.addStretch(1)

        card = QFrame()
        card.setObjectName("startupLoadingCard")
        card.setMaximumWidth(480)
        card.setMinimumWidth(320)

        self._startup_card_opacity_effect = QGraphicsOpacityEffect(card)
        self._startup_card_opacity_effect.setOpacity(1.0)
        card.setGraphicsEffect(self._startup_card_opacity_effect)

        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(32, 28, 32, 28)
        card_lay.setSpacing(12)

        badge = QLabel("Easy Fulfill")
        badge.setAlignment(Qt.AlignCenter)
        badge.setProperty("overlayText", "badge")

        title = QLabel("초기 데이터를 불러오는 중…")
        title.setAlignment(Qt.AlignCenter)
        title.setProperty("overlayText", "title")

        sub = QLabel("Google 시트와 주문 인덱스를 동기화합니다.")
        sub.setAlignment(Qt.AlignCenter)
        sub.setProperty("overlayText", "sub")
        sub.setWordWrap(True)

        self._startup_progress_bar = QProgressBar()
        self._startup_progress_bar.setObjectName("startupProgressBar")
        self._startup_progress_bar.setRange(0, 100)
        self._startup_progress_bar.setValue(0)
        self._startup_progress_bar.setTextVisible(True)
        self._startup_progress_bar.setFormat("%p%")
        self._startup_progress_bar.setFixedSize(300, 20)
        self._startup_progress_timer = QTimer(self)
        self._startup_progress_timer.setInterval(STARTUP_SYNC_PROGRESS_TICK_MS)
        self._startup_progress_timer.timeout.connect(self._on_startup_progress_tick)
        self._startup_sync_progress_done = False

        bar_row = QHBoxLayout()
        bar_row.addStretch(1)
        bar_row.addWidget(self._startup_progress_bar)
        bar_row.addStretch(1)

        hint = QLabel("환경에 따라 수 초 걸릴 수 있습니다.")
        hint.setAlignment(Qt.AlignCenter)
        hint.setProperty("overlayText", "hint")
        hint.setWordWrap(True)

        card_lay.addWidget(badge)
        card_lay.addWidget(title)
        card_lay.addWidget(sub)
        card_lay.addSpacing(8)
        card_lay.addLayout(bar_row)
        card_lay.addWidget(hint)

        row.addWidget(card, alignment=Qt.AlignmentFlag.AlignHCenter)
        row.addStretch(1)
        outer.addLayout(row)
        outer.addStretch(1)

        self._startup_card_entrance_anim = None

    def _setup_busy_processing_overlay(self):
        self._busy_overlay = QFrame(self)
        self._busy_overlay.setObjectName("busyProcessingOverlay")
        self._busy_overlay.hide()
        outer = QVBoxLayout(self._busy_overlay)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addStretch(1)
        row = QHBoxLayout()
        row.addStretch(1)

        card = QFrame()
        card.setObjectName("busyProcessingCard")
        card.setMaximumWidth(440)
        card.setMinimumWidth(280)

        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(28, 24, 28, 24)
        card_lay.setSpacing(14)

        spin_row = QHBoxLayout()
        spin_row.addStretch(1)
        self._busy_spinner = CircularBusySpinner(card, size=56, line_width=4, color=theme.PRIMARY)
        spin_row.addWidget(self._busy_spinner)
        spin_row.addStretch(1)

        self._busy_overlay_title = QLabel("처리 중…")
        self._busy_overlay_title.setAlignment(Qt.AlignCenter)
        self._busy_overlay_title.setProperty("overlayText", "title")

        self._busy_overlay_sub = QLabel("")
        self._busy_overlay_sub.setAlignment(Qt.AlignCenter)
        self._busy_overlay_sub.setProperty("overlayText", "sub")
        self._busy_overlay_sub.setWordWrap(True)

        card_lay.addLayout(spin_row)
        card_lay.addWidget(self._busy_overlay_title)
        card_lay.addWidget(self._busy_overlay_sub)

        row.addWidget(card, alignment=Qt.AlignmentFlag.AlignHCenter)
        row.addStretch(1)
        outer.addLayout(row)
        outer.addStretch(1)

    def _layout_busy_overlay_full(self):
        if getattr(self, "_busy_overlay", None) is None:
            return
        self._busy_overlay.setGeometry(
            0, 0, max(1, self.width()), max(1, self.height())
        )

    def _show_busy_processing_overlay(self, title: str, subtitle: str = ""):
        if getattr(self, "_busy_overlay", None) is None:
            return
        self._busy_overlay_title.setText(title)
        self._busy_overlay_sub.setText(subtitle)
        self._busy_overlay_sub.setVisible(bool(subtitle and subtitle.strip()))
        self._layout_busy_overlay_full()
        self._busy_overlay.raise_()
        if (
            getattr(self, "_startup_overlay", None) is not None
            and self._startup_overlay.isVisible()
        ):
            self._startup_overlay.raise_()
        self._busy_overlay.show()

    def _hide_busy_processing_overlay(self):
        if getattr(self, "_busy_overlay", None) is not None:
            self._busy_overlay.hide()

    def _on_startup_progress_tick(self):
        if getattr(self, "_startup_sync_progress_done", False):
            return
        t0 = getattr(self, "_startup_sync_debug_t0", None)
        if t0 is None:
            return
        elapsed_ms = (time.perf_counter() - t0) * 1000
        if elapsed_ms <= 0:
            return
        pct = min(99, int(elapsed_ms * 100 / STARTUP_SYNC_PROGRESS_CAP_MS))
        pb = getattr(self, "_startup_progress_bar", None)
        if pb is not None and pb.value() < pct:
            pb.setValue(pct)

    def _hide_startup_loading_overlay_deferred(self):
        if getattr(self, "_startup_overlay", None) is not None:
            self._startup_overlay.hide()

    def _start_startup_overlay_entrance_anim(self):
        self._startup_card_opacity_effect.setOpacity(0.0)
        anim = QPropertyAnimation(self._startup_card_opacity_effect, b"opacity", self)
        anim.setDuration(280)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._startup_card_entrance_anim = anim
        anim.start()

    def _set_startup_syncing_indicator(self):
        """시작 동기화 진행 중임을 막지 않는 가벼운 라벨 표시로만 알립니다."""
        if hasattr(self.ui, "label_index_sheet_sync"):
            set_status(self.ui.label_index_sheet_sync, "idle", "Google 시트와 인덱스 동기화 중…")
        self._update_index_status_chip(syncing=True)

    def setup_menubar(self):
        """메뉴/단축키용 액션을 정의하고 상단 메뉴바를 구성합니다(툴바 없음)."""

        # ── 액션 정의(메뉴·단축키용) ─────────────────────────────
        self.act_load = QAction(QIcon('image/open-file-icon.png'), '주문 불러오기', self)
        self.act_load.setShortcut('Ctrl+O')
        self.act_load.setStatusTip('주문 엑셀 파일 불러오기 (Ctrl+O)')
        self.act_load.triggered.connect(self.select_excel_file)

        self.act_load_invoice = QAction(QIcon('image/open-file-icon.png'), '송장 불러오기', self)
        self.act_load_invoice.setStatusTip('우체국 송장 엑셀 파일 불러오기 (Ctrl+Shift+O)')
        self.act_load_invoice.setShortcut('Ctrl+Shift+O')
        self.act_load_invoice.triggered.connect(self.load_invoice_file)

        self.act_copy = QAction(QIcon('image/copy-icon.png'), '복사', self)
        self.act_copy.setShortcut('Ctrl+C')
        self.act_copy.setStatusTip('주문 정보를 클립보드에 복사 (Ctrl+C)')
        self.act_copy.triggered.connect(self.copy_to_clipboard)

        self.act_export = QAction(QIcon('image/microsoft-excel-icon.png'), '송장 출력', self)
        self.act_export.setShortcut('Ctrl+E')
        self.act_export.setStatusTip('네이버 업로드용 엑셀 송장 출력 (Ctrl+E)')
        self.act_export.triggered.connect(self.export_invoice_excel)

        self.act_openfolder = QAction(QIcon('image/folder-icon.png'), '출력 폴더', self)
        self.act_openfolder.setShortcut('Ctrl+F')
        self.act_openfolder.setStatusTip('출력 폴더 열기 (Ctrl+F)')
        self.act_openfolder.triggered.connect(self.open_output_folder)

        self.act_courier_export = QAction(
            QIcon('image/microsoft-excel-icon.png'), '택배사 접수목록 내보내기(당일)…', self)
        self.act_courier_export.setStatusTip(
            '오늘 「송장추적」에 등록된 송장을 순번/송장번호/이름 엑셀로 내보내기(택배사 제출용)')
        self.act_courier_export.triggered.connect(self.export_courier_receipt)

        self.act_clear = QAction(QIcon('image/reset-icon.png'), '초기화', self)
        self.act_clear.setShortcut('Ctrl+R')
        self.act_clear.setStatusTip('주문 정보 초기화 (Ctrl+R)')
        self.act_clear.triggered.connect(self.clear_list)

        # API·우체국 동작 — 툴바엔 두지 않는다. 진입점·실제 동작은 메뉴, 확인·진단은 「설정 › 프로그램」 버튼
        self.act_api_load = QAction(QIcon('image/open-file-icon.png'), 'API로 주문 불러오기', self)
        self.act_api_load.setStatusTip('네이버 주문을 API로 직접 불러오기(발송 전 결제완료·발주확인)')
        self.act_api_load.triggered.connect(self.load_naver_orders_via_api)

        self.act_dispatch = QAction(QIcon('image/microsoft-excel-icon.png'), 'API 발송처리', self)
        self.act_dispatch.setStatusTip('매칭된 송장번호를 네이버에 API로 발송처리(배송중 전환)')
        self.act_dispatch.triggered.connect(self.dispatch_naver_via_api)

        self.act_epost_test_receipt = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 테스트 접수(송장 엑셀)', self)
        self.act_epost_test_receipt.setStatusTip(
            '송장 엑셀 1건을 testYn=Y로 우체국 계약소포 API에 테스트 접수')
        self.act_epost_test_receipt.triggered.connect(self.run_epost_test_receipt)

        self.act_epost_real_receipt = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 실제 접수 확인(송장 엑셀)', self)
        self.act_epost_real_receipt.setStatusTip(
            '송장 엑셀 1건을 실제 접수하되 운송장 출력과 네이버 발송은 하지 않음')
        self.act_epost_real_receipt.triggered.connect(self.run_epost_real_receipt)

        self.act_epost_print_targets = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 출력 대상 확인', self)
        self.act_epost_print_targets.setStatusTip(
            '프로그램이 실제 접수한 미출력 등기번호 목록을 확인')
        self.act_epost_print_targets.triggered.connect(self.show_epost_print_targets)

        self.act_epost_portal_login = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 포털 로그인 연결', self)
        self.act_epost_portal_login.setStatusTip(
            '전용 Chromium 창에서 직접 로그인해 우체국 포털 세션을 이 PC에 연결')
        self.act_epost_portal_login.triggered.connect(self.connect_epost_portal_login)

        self.act_epost_portal_diagnostic = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 출력 화면 진단', self)
        self.act_epost_portal_diagnostic.setStatusTip(
            '전용 Chromium에서 운송장출력 화면의 컨트롤만 읽기 전용으로 점검')
        self.act_epost_portal_diagnostic.triggered.connect(self.run_epost_portal_diagnostic)

        self.act_epost_portal_lookup = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 신규출력 대조', self)
        self.act_epost_portal_lookup.setStatusTip(
            '오늘 신규출력 목록을 조회해 프로그램의 실제 접수 등기번호와 읽기 전용으로 대조')
        self.act_epost_portal_lookup.triggered.connect(self.run_epost_portal_lookup)

        self.act_epost_portal_print_popup = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 운송장출력 팝업 열기', self)
        self.act_epost_portal_print_popup.setStatusTip(
            '검증된 오늘 미출력 행만 선택해 포털 운송장출력 팝업을 열고 인쇄 전에서 멈춤')
        self.act_epost_portal_print_popup.triggered.connect(self.run_epost_portal_print_popup)

        self.act_epost_portal_oz_viewer = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 OZ 뷰어 열기 확인', self)
        self.act_epost_portal_oz_viewer.setStatusTip(
            '검증된 행의 포털 인쇄 요청 후 OZ Viewer 열림만 확인하고 실제 인쇄 전에서 멈춤')
        self.act_epost_portal_oz_viewer.triggered.connect(self.run_epost_portal_oz_viewer)

        self.act_epost_portal_output_confirm = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 포털 출력여부 확인', self)
        self.act_epost_portal_output_confirm.setStatusTip(
            '포털 인쇄 요청 건의 출력여부를 전체 조회로 읽기 전용 확인')
        self.act_epost_portal_output_confirm.triggered.connect(self.run_epost_portal_output_confirm)

        self.act_epost_portal_reprint_popup = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 단건 재출력 팝업 열기', self)
        self.act_epost_portal_reprint_popup.setStatusTip(
            '포털 출력 확인 이력의 정확한 단건만 선택해 재출력 팝업을 열고 인쇄 전에서 멈춤')
        self.act_epost_portal_reprint_popup.triggered.connect(self.run_epost_portal_reprint_popup)

        self.act_epost_portal_reprint_oz_viewer = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 단건 재출력 OZ 뷰어 확인', self)
        self.act_epost_portal_reprint_oz_viewer.setStatusTip(
            '출력 확인 단건의 포털 인쇄 요청 후 OZ Viewer 열림만 확인하고 실제 인쇄 전에서 멈춤')
        self.act_epost_portal_reprint_oz_viewer.triggered.connect(
            lambda: self.run_epost_portal_oz_viewer(reprint=True))

        self.act_epost_oz_print_dialog = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 단건 재출력 인쇄 창 확인', self)
        self.act_epost_oz_print_dialog.setStatusTip(
            '출력 확인 단건의 OZ Viewer에서 Windows 인쇄 창 열림만 확인하고 확인 버튼 전에서 멈춤')
        self.act_epost_oz_print_dialog.triggered.connect(self.run_epost_oz_print_dialog)

        self.act_epost_actual_print = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 단건 실제 인쇄', self)
        self.act_epost_actual_print.setStatusTip(
            '확정된 단건의 Windows 인쇄 창 확인을 눌러 실제 프린터 요청을 전송')
        self.act_epost_actual_print.triggered.connect(
            lambda: self.run_epost_oz_print_dialog(execute_print=True))

        self.act_epost_batch_actual_print = QAction(
            QIcon('image/korea-post-icon.png'), '우체국 다건 실제 인쇄', self)
        self.act_epost_batch_actual_print.setStatusTip(
            '같은 접수일의 포털 출력 확인 건만 모두 선택해 실제 프린터 요청을 전송')
        self.act_epost_batch_actual_print.triggered.connect(
            lambda: self.run_epost_oz_print_dialog(execute_print=True, allow_batch=True))

        # 외부 바로가기
        self.act_db = QAction(QIcon('image/database-icon.png'), '데이터베이스 시트', self)
        self.act_db.setShortcut('Ctrl+D')
        self.act_db.setStatusTip('공유 데이터베이스 구글시트 열기 (Ctrl+D)')
        self.act_db.triggered.connect(lambda: QDesktopServices.openUrl(QUrl(
            "https://docs.google.com/spreadsheets/d/1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw/edit?gid=195401368#gid=195401368")))

        self.act_epost = QAction(QIcon('image/korea-post-icon.png'), '우체국', self)
        self.act_epost.setStatusTip('우체국 홈페이지로 이동')
        self.act_epost.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://biz.epost.go.kr/ui/index.jsp")))

        self.act_naver_store = QAction(QIcon('image/smart-store-icon.png'), '네이버 스마트스토어', self)
        self.act_naver_store.setStatusTip('네이버 스마트스토어 페이지로 이동')
        self.act_naver_store.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://sell.smartstore.naver.com/#/home/about")))

        self.act_coupang_store = QAction(QIcon('image/coupang-wing-icon.png'), '쿠팡 윙', self)
        self.act_coupang_store.setStatusTip('쿠팡 스토어 페이지로 이동')
        self.act_coupang_store.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://wing.coupang.com/")))

        self.act_exit = QAction(QIcon('image/exit-icon.png'), '종료', self)
        self.act_exit.setShortcut('Ctrl+Q')
        self.act_exit.setStatusTip('프로그램 종료 (Ctrl+Q)')
        self.act_exit.triggered.connect(QApplication.quit)

        # 상단 메뉴바 구성
        self._setup_menus()

    def _real_api_actions(self):
        """외부에 실제로 반영되는 동작(네이버 발송처리, 우체국 실제 접수·인쇄). 모두 실행 전 확인 창을 띄운다."""
        return [self.act_dispatch, self.act_epost_real_receipt,
                self.act_epost_actual_print, self.act_epost_batch_actual_print]

    def _setup_menus(self):
        """상단 메뉴바를 구성합니다(생성 UI의 기본 메뉴 대신 여기서 일괄 정의).

        툴바엔 매일 쓰는 파일 기반 동작만 두고, 나머지(API·외부링크·종료)는
        여기로 모은다. API 항목은 아직 개발 단계라 별도 그룹으로 분리.
        """
        mb = self.menuBar()
        mb.clear()
        # 메뉴바와 탭 사이 경계선·여백은 테마(QMenuBar 규칙)에서 정한다.
        # 구분선과 탭 헤더 사이 여백(레이아웃 안 위젯은 stylesheet margin이 안 먹어
        # 레이아웃 상단 마진으로 준다)
        lay = getattr(self.ui, "verticalLayout", None)
        if lay is not None:
            m = lay.contentsMargins()
            lay.setContentsMargins(m.left(), 14, m.right(), m.bottom())

        m_file = mb.addMenu('파일(&F)')
        m_file.addAction(self.act_load)
        m_file.addAction(self.act_load_invoice)
        m_file.addAction(self.act_openfolder)
        m_file.addSeparator()
        m_file.addAction(self.act_courier_export)
        m_file.addSeparator()
        m_file.addAction(self.act_exit)

        m_order = mb.addMenu('주문·발송(&O)')
        m_order.addAction(self.act_copy)
        m_order.addAction(self.act_export)
        m_order.addSeparator()
        m_order.addAction(self.act_clear)

        # 검토 리스트 R7(2026-11-19 사용자 결정): API 기능은 개발 중이라 실무자는 쓰지 않는다. 메뉴바에 두지 않고
        # 「⚙ 설정 › 프로그램」의 「개발 중 기능」 묶음에 모은다. 기본은 비활성, 「켜기(이 PC)」를 체크한 PC만 켜진다.
        warning_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxWarning)
        for action in self._real_api_actions():
            action.setIcon(warning_icon)

        m_link = mb.addMenu('바로가기(&L)')
        m_link.addAction(self.act_db)
        m_link.addAction(self.act_epost)
        m_link.addAction(self.act_naver_store)
        m_link.addAction(self.act_coupang_store)

        m_help = mb.addMenu('도움말(&H)')
        self.act_shortcut_help = QAction('단축키(&K)', self)
        self.act_shortcut_help.setStatusTip('이 프로그램의 단축키 목록')
        self.act_shortcut_help.triggered.connect(self.show_shortcut_help)
        m_help.addAction(self.act_shortcut_help)
        self.act_program_log = QAction('프로그램 로그(&L)', self)
        self.act_program_log.setStatusTip('프로그램이 남긴 기록(시작·동기화·오류 메시지)을 시각과 함께 보기')
        self.act_program_log.triggered.connect(self.show_program_log)
        m_help.addAction(self.act_program_log)
        self.act_about = QAction('프로그램 정보(&A)', self)
        self.act_about.triggered.connect(self.show_about)
        m_help.addAction(self.act_about)

    def setup_status_bar(self):
        """상태바를 초기화하고 기본 메시지를 설정합니다."""
        statusbar = self.statusBar()
        
        # 기본 레이블 생성 및 설정(상태바답게 좌측·작게·일반 굵기 — 제목/버튼처럼 보이지 않도록)
        self.status_label = QLabel()
        self.status_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.status_label.setMinimumHeight(24)
        self.status_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.status_label.setObjectName("statusMessageLabel")  # 색·여백은 테마
        
        # 상태바에 레이블을 추가하고 너비를 최대로 설정
        statusbar.addWidget(self.status_label, 1)  # stretch factor 1로 설정하여 최대 너비 사용
        
        # 상태바 바탕·경계선·높이는 테마(QStatusBar 규칙)에서 정한다.
        
        # 메뉴 안내·진행 메시지가 끝나면 마지막 알림을 다시 보여 준다.
        self._last_notice_text = ""
        self._status_clear_timer = QTimer(self)
        self._status_clear_timer.setSingleShot(True)
        self._status_clear_timer.timeout.connect(lambda: self.status_label.setText(self._last_notice_text))
        STATUS_AUTO_CLEAR_MS = 3000

        def update_status_label(message):
            """상태바 메시지 업데이트 시 레이블도 함께 업데이트"""
            if message:  # 메시지가 있을 때만 업데이트
                self.status_label.setText(message)
                statusbar.clearMessage()  # 기본 메시지 클리어
                # 일정 시간 후 자동으로 메시지 제거
                self._status_clear_timer.start(STATUS_AUTO_CLEAR_MS)

        # 상태바 메시지 변경 시그널 연결
        statusbar.messageChanged.connect(update_status_label)
        statusbar.showMessage("준비")  # 초기 메시지 설정

    def notify(self, kind, text, actions=(), detail=""):
        """확인 창 대신 알림 띠로 알린다. kind: ok·warn·error. actions: [(버튼 글자, 함수), …](최대 2개).

        같은 문구를 상태바에도 남겨, 띠가 사라진 뒤에도 마지막 알림을 볼 수 있다.
        """
        self._last_notice_text = f"{STATUS_MARKS.get(kind, '')} {text}".strip()
        self._status_clear_timer.stop()
        self.status_label.setText(self._last_notice_text)
        bar = getattr(self, "notice_bar", None)
        if bar is not None:
            bar.show_notice(kind, text, actions, detail)

    def _open_created_file(self, output_file):
        if not self.open_file_with_default_app(output_file):
            QMessageBox.warning(self, "오류", "파일을 열 수 없습니다.\n연결된 기본 앱을 확인해주세요.")

    def _open_created_file_location(self, output_file):
        if not self.open_file_location(output_file):
            QMessageBox.warning(self, "오류", "파일 위치를 열 수 없습니다.")

    def show_excel_created_message(self, output_file, message_text, open_file_label="엑셀 열기"):
        """파일을 만든 뒤 알림 띠에 「열기」·「폴더 열기」를 보인다(확인 창 없음, 계획 C-1)."""
        name = Path(str(output_file)).name
        self.notify(
            "ok",
            f"{message_text.rstrip('.')}: {name}",
            [(open_file_label, lambda: self._open_created_file(output_file)),
             ("폴더 열기", lambda: self._open_created_file_location(output_file))],
            detail=str(output_file),
        )

    # 탭 안에서만 동작하는 단축키(QAction이 아닌 것). 단축키 목록 창에 함께 보인다.
    EXTRA_SHORTCUT_HELP = (
        ("단건 송장 탭", "엑셀 생성", "Ctrl+Enter"),
    )

    def _setup_tab_shortcuts(self, tabs):
        """Ctrl+1~6: 상위 탭으로 바로 이동(왼쪽부터 순서대로). 탭 순서가 바뀌어도 위치 기준으로 동작한다."""
        self._tab_shortcuts = []
        for index in range(6):
            shortcut = QShortcut(QKeySequence(f"Ctrl+{index + 1}"), self)
            shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
            shortcut.activated.connect(lambda i=index: tabs.setCurrentIndex(i) if i < tabs.count() else None)
            self._tab_shortcuts.append(shortcut)

    def shortcut_help_rows(self):
        """단축키 목록 창의 행: (구분, 동작, 단축키). 메뉴 동작의 단축키는 실제 QAction에서 모은다."""
        rows = []
        tabs = self.ui.tabWidget
        for index, shortcut in enumerate(getattr(self, "_tab_shortcuts", [])):
            if index < tabs.count():
                title = tabs.tabText(index).replace("⚙", "").strip()
                rows.append(("탭 이동", f"「{title}」 탭", shortcut.key().toString()))
        seen = set()
        for menu in self.menuBar().findChildren(QMenu):
            for action in menu.actions():
                key = action.shortcut().toString()
                if not key or id(action) in seen:
                    continue
                seen.add(id(action))
                rows.append((menu.title().split("(")[0], action.text().replace("&", ""), key))
        rows.extend(self.EXTRA_SHORTCUT_HELP)
        return rows

    def show_shortcut_help(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("단축키")
        layout = QVBoxLayout(dialog)
        rows = self.shortcut_help_rows()
        table = QTableWidget(len(rows), 3, dialog)
        table.setObjectName("tableWidget_shortcuts")
        table.setHorizontalHeaderLabels(["구분", "동작", "단축키"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                table.setItem(r, c, QTableWidgetItem(text))
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)
        close = QPushButton("닫기")
        close.clicked.connect(dialog.accept)
        layout.addWidget(close, 0, Qt.AlignmentFlag.AlignRight)
        dialog.resize(460, 420)
        dialog.exec()

    def show_program_log(self):
        """「도움말 › 프로그램 로그」: 모달이 아닌 창 하나를 띄워 두고 다시 누르면 앞으로 가져온다."""
        window = getattr(self, "_program_log_window", None)
        if window is None:
            from ui.log_window import LogWindow
            window = self._program_log_window = LogWindow(self)
        window.show()
        window.raise_()
        window.activateWindow()

    def show_about(self):
        """프로그램 정보를 보여주는 대화상자를 표시합니다."""
        QMessageBox.about(
            self,
            "프로그램 정보",
            "Easy Fulfill - 주문 처리 시스템\n\n"
            f"버전 {APP_VERSION}\n"
            "Copyright © 2024\n\n"
            "오픈마켓 주문 처리를 위한 자동 생성 프로그램입니다."
        )
        self.statusBar().showMessage("프로그램 정보 표시됨")
