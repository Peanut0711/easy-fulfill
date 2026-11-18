"""배송추적 탭의 UI 제어와 화면 상태를 담당한다.

이 모듈은 Qt widget 신호를 연결하고 worker 결과를 화면에 반영한다.
Google Sheets/KPOST I/O는 tracking.workers 및 tracking.repository에 위임한다.
"""

from __future__ import annotations

import os
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

from shared import theme
from shared.ui_components import fixed_width_button, setup_table
from shared.lazy_import import optional_deferred_module
from shared.ui_style import clear_status, set_button_role, set_status

gspread = optional_deferred_module("gspread")  # 설치 여부만 확인하고 처음 쓸 때 불러온다.

from PySide6.QtCore import QEvent, QObject, Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLineEdit,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .repository import normalize_tracking_no as _normalize_tracking_no
from .service import (
    CONFIG_KEY_STALE_HUB,
    CONFIG_KEY_STALE_PICKUP,
    CONFIG_KEY_STALE_REMOTE_BONUS,
    CONFIG_KEY_STALE_TRANSIT,
    STALE_DEFAULTS,
    STALE_CONFIG_KEYS,
    STALE_REMOTE_BONUS_DEFAULT,
    TRACKING_COMPLETED_LOOKBACK_DAYS,
    TRACKING_MANAGEMENT_ACTIVE,
    TRACKING_MANAGEMENT_CANDIDATE,
    TRACKING_MANAGEMENT_COL,
    TRACKING_MANAGEMENT_DISCARDED,
    TRACKING_MANAGEMENT_EXCLUDED,
    TRACKING_MANAGEMENT_MANUAL_STOP,
    build_risk_digest_text,
    evaluate_risk,
    is_weekday as _is_weekday,
    risk_signature,
)
from .workers import (
    CONFIG_KEY_KPOST_REGKEY,
    CONFIG_KEY_SLACK_WEBHOOK,
    DigestSendThread,
    SlackSendThread,
    TrackingConfigReadThread,
    TrackingConfigWriteThread,
    CourierReceiptExportThread,
    TrackingKeyValidateThread,
    TrackingListThread,
    TrackingManagementUpdateThread,
    TrackingNotesUpdateThread,
    TrackingRefreshOneThread,
    TrackingRefreshThread,
    TrackingUpsertThread,
)


TRACKING_SHEET_PUSH_DEBOUNCE_MS = 1200
TRACKING_TABLE_REFRESH_DEBOUNCE_MS = 250  # 검색 입력·기준 변경 후 표를 다시 그리기까지 대기
CONFIG_KEY_AUTO_REFRESH_MIN = "tracking_auto_refresh_min"
TRACKING_AUTO_REFRESH_DEFAULT_MIN = 60
# 배송추적 표 보기(이 PC 앱 설정). 숨긴 열 이름 목록. 설정이 없으면 기본 숨김 열을 쓴다.
SETTING_TRACKING_HIDDEN_COLUMNS = "tracking_hidden_columns"
TRACKING_DEFAULT_HIDDEN_COLUMNS = ("최근조회", "스토어")
TRACKING_ALWAYS_VISIBLE_COLUMNS = ("수취인명",)
TRACKING_COLUMN_MAX_WIDTH = 220   # 내용이 길어도 한 열이 이보다 넓어지지 않는다(원문은 툴팁)
TRACKING_LAST_COLUMN_MIN_WIDTH = 80  # 마지막 열(메모)에 남겨 둘 폭
TRACKING_TOOLTIP_COLUMNS = (2, 6, 8, 10, 11)  # 스토어·배송상태·마지막위치·메모·최근이벤트
# 위험 칩: key → 글자. "all"은 위험 전체.
TRACKING_RISK_CHIPS = (("all", "⚠ 위험"), ("허브정체", "허브"), ("수거누락", "수거누락"), ("이동정체", "이동"))
COURIER_RECEIPT_PREFIX = "하이제니스"


class TrackingController(QObject):
    """MainWindow에 보관된 상태와 배송추적 UI의 연결을 관리한다."""

    def __init__(self, host):
        super().__init__(host)
        self.host = host
        self.ui = getattr(host, "ui", None)

    def initialize_runtime(self):
        """배송추적 전용 비동기 상태와 타이머를 한곳에서 초기화한다."""
        self._tracking_op_thread = None
        self._tracking_pending_records = []
        self._tracking_inflight_records = []
        self._tracking_push_timer = self._new_timer(single_shot=True)
        self._tracking_push_timer.timeout.connect(self._flush_tracking_records)
        self._tracking_refresh_thread = None
        self._tracking_config_read_thread = None
        self._tracking_config_write_thread = None
        # 쓰기 진행 중에 들어온 공유 설정 변경(키 → 최신 값). 끝나면 이어서 쓴다.
        self._tracking_config_pending_writes = {}
        self._tracking_key_validate_thread = None
        self._key_validate_mode = "status"
        self._key_validate_pending_key = ""
        self._tracking_list_thread = None
        self._tracking_list_values = []
        self._tracking_management_update_thread = None
        self._tracking_notes_update_thread = None
        self._tracking_notes_pending = {}
        self._tracking_notes_inflight = {}
        self._populating_tracking_table = False
        self._courier_export_thread = None
        self._tracking_list_timer = self._new_timer(interval=120000)
        self._tracking_list_timer.timeout.connect(self._on_tracking_list_poll)
        self._tracking_list_timer.start()
        self._tracking_auto_refresh_timer = self._new_timer()
        self._tracking_auto_refresh_timer.timeout.connect(self._on_tracking_auto_refresh_tick)
        self._reconfigure_auto_refresh_timer()
        self._tracking_refresh_one_thread = None
        self._slack_send_thread = None
        self._digest_thread = None
        self._slack_notify_after_reload = False

    def _new_timer(self, interval=None, single_shot=False):
        from PySide6.QtCore import QTimer

        timer = QTimer(self.host)
        timer.setSingleShot(single_shot)
        if interval is not None:
            timer.setInterval(interval)
        return timer

    def __getattr__(self, name):
        return getattr(self.host, name)

    def __setattr__(self, name, value):
        if name in {"host", "ui"}:
            super().__setattr__(name, value)
        elif (name.startswith("_tracking_") or name.startswith("_slack_")
              or name.startswith("_dlg_") or name == "_admin_tracking_built"
              or name in {
            "_populating_tracking_table",
            "_slack_notify_after_reload",
            "_digest_thread",
        }):
            setattr(self.host, name, value)
        else:
            super().__setattr__(name, value)

    def bind_ui(self):
        """배송추적 탭 object name과 기존 signal 연결 계약을 유지한다."""
        self.ui = self.host.ui
        if hasattr(self.ui, "pushButton_refresh_tracking"):
            self.ui.pushButton_refresh_tracking.clicked.connect(
                self.on_refresh_tracking_clicked
            )
        if hasattr(self.ui, "pushButton_refresh_tracking_exceptions"):
            self.ui.pushButton_refresh_tracking_exceptions.clicked.connect(
                self.on_refresh_tracking_exceptions_clicked
            )
        if hasattr(self.ui, "pushButton_tracking_stop_selected"):
            self.ui.pushButton_tracking_stop_selected.clicked.connect(
                self._on_tracking_stop_selected_clicked
            )
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is not None:
            table.cellDoubleClicked.connect(self._on_tracking_cell_double_clicked)
            table.itemChanged.connect(self._on_tracking_item_changed)
            table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            table.customContextMenuRequested.connect(self._on_tracking_context_menu)
            header = table.horizontalHeader()
            header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            header.customContextMenuRequested.connect(self._on_tracking_header_menu)
            header.setToolTip("열 머리를 오른쪽 클릭하면 보일 열을 고를 수 있습니다.")
            # 창 폭이 바뀌면 열 폭을 다시 맞추고 빈 목록 안내 위치를 옮긴다.
            table.viewport().installEventFilter(self)
            self._build_tracking_risk_chips()
        self._apply_tracking_design()
        if hasattr(self.ui, "comboBox_tracking_filter"):
            self.ui.comboBox_tracking_filter.currentIndexChanged.connect(
                self._on_tracking_filter_changed
            )
        if hasattr(self.ui, "lineEdit_tracking_recipient_search"):
            # 글자마다 표 전체를 다시 그리지 않고 입력이 멈춘 뒤 한 번 그린다.
            self.ui.lineEdit_tracking_recipient_search.textChanged.connect(
                self._schedule_tracking_table_refresh
            )

    def eventFilter(self, obj, event):
        if event.type() != QEvent.Type.Resize:
            return False
        try:
            table = getattr(getattr(self, "ui", None), "tableWidget_tracking", None)
            if table is not None and obj is table.viewport():
                if event.size().width() != event.oldSize().width():
                    self._fit_tracking_columns()
                self._place_tracking_empty_label()
        except RuntimeError:  # 창을 닫는 중 표가 먼저 지워진 경우
            pass
        return False

    def _apply_tracking_design(self):
        """배송추적 탭 배치(디자인 개편 4-2).

        - 위 줄: 건수·위험 칩·진행 요약(왼쪽, 좁으면 말줄임) · 동작 버튼 3개(오른쪽).
        - 아래 줄: 분류 · 수취인명 검색(남는 폭).
        버튼은 글자 폭 그대로 두고(.ui의 최대 폭 제한 해제), 열이 다 들어가지 않으면 표 머리글은 끝을 「…」로 줄인다.
        """
        ui = self.ui
        page_layout = getattr(ui, "verticalLayout_tracking", None)
        if page_layout is not None:
            page_layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
            page_layout.setSpacing(theme.SPACE_SM)
        status_row = getattr(ui, "horizontalLayout_tracking_filter", None)
        filter_row = getattr(ui, "horizontalLayout_tracking_actions", None)
        buttons = [getattr(ui, name, None) for name in (
            "pushButton_tracking_stop_selected", "pushButton_refresh_tracking",
            "pushButton_refresh_tracking_exceptions")]
        if status_row is not None and filter_row is not None:
            status_row.setSpacing(theme.SPACE_SM)
            for button in buttons:
                if button is None:
                    continue
                filter_row.removeWidget(button)
                button.setMinimumSize(0, 0)  # .ui의 최소 높이 34·최대 폭 130/150 고정을 풀어 테마 크기를 쓴다
                button.setMaximumWidth(16777215)
                fixed_width_button(button)
                status_row.addWidget(button)
            # 아래 줄: 버튼이 빠진 자리의 빈 칸을 지우고 검색칸이 남는 폭을 쓴다.
            for i in reversed(range(filter_row.count())):
                if filter_row.itemAt(i).spacerItem() is not None:
                    filter_row.takeAt(i)
            filter_row.setSpacing(theme.SPACE_SM)
            search = getattr(ui, "lineEdit_tracking_recipient_search", None)
            if search is not None:
                filter_row.setStretchFactor(search, 1)
        table = getattr(ui, "tableWidget_tracking", None)
        if table is not None:
            setup_table(table, single=False, stretch_last=table.horizontalHeader().stretchLastSection())
            table.horizontalHeader().setTextElideMode(Qt.TextElideMode.ElideRight)

    def _build_tracking_risk_chips(self):
        """건수 줄에 위험 칩(누르면 그 위험만 거름)을 넣는다. .ui에 건수 줄이 없으면 만들지 않는다."""
        layout = getattr(self.ui, "horizontalLayout_tracking_filter", None)
        count_label = getattr(self.ui, "label_tracking_count", None)
        if layout is None or count_label is None or getattr(self, "_tracking_risk_chips", None):
            return
        chips = {}
        index = layout.indexOf(count_label) + 1
        for offset, (key, _text) in enumerate(TRACKING_RISK_CHIPS):
            chip = QPushButton()
            chip.setObjectName(f"pushButton_tracking_risk_{'all' if key == 'all' else offset}")
            chip.setCheckable(True)
            chip.setToolTip("누르면 이 위험 송장만 보입니다. 다시 누르면 모두 보입니다.")
            chip.setProperty("chip", "true")  # 작은 여백(공통 스타일)
            set_button_role(chip, "danger")
            chip.clicked.connect(lambda _=False, k=key: self._on_tracking_risk_chip_clicked(k))
            chip.hide()
            layout.insertWidget(index + offset, chip)
            chips[key] = chip
        self._tracking_risk_chips = chips
        # 칩이 자리를 차지해도 건수 글자(마지막 조회·자동 간격)가 잘리지 않게 진행 문구보다 폭을 두 배로 받는다.
        layout.setStretchFactor(count_label, 2)

    def _on_tracking_risk_chip_clicked(self, key):
        current = getattr(self, "_tracking_risk_filter", None)
        self._tracking_risk_filter = None if current == key else key
        self._populate_tracking_table()

    def _tracking_view_settings(self):
        """표 한 번 그릴 때 쓰는 설정(정체 기준·도서산간 가산, 숨긴 열, 자동 새로고침 간격)을 한 번에 읽는다."""
        reader = getattr(self.host, "read_app_settings", None)
        if callable(reader):
            settings = reader()
            get = lambda key: settings.get(key, "")  # noqa: E731
            hidden = settings.get(SETTING_TRACKING_HIDDEN_COLUMNS)
            auto = settings.get(CONFIG_KEY_AUTO_REFRESH_MIN, "")
        else:
            get = None
            hidden = self.get_app_setting(SETTING_TRACKING_HIDDEN_COLUMNS, None)
            auto = self.get_app_setting(CONFIG_KEY_AUTO_REFRESH_MIN, "")
        thresholds = {category: self._stale_threshold(category, get) for category in STALE_DEFAULTS}
        if not isinstance(hidden, (list, tuple)):
            hidden = TRACKING_DEFAULT_HIDDEN_COLUMNS
        hidden = {name for name in hidden if name not in TRACKING_ALWAYS_VISIBLE_COLUMNS}
        return (thresholds, self._remote_bonus_hours(get)), hidden, self._parse_auto_refresh_min(auto)

    def _on_tracking_header_menu(self, pos):
        """열 머리 오른쪽 클릭: 보일 열을 고른다(이 PC에 저장)."""
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return
        menu = QMenu(table)
        menu.addSection("보일 열")
        for index, (_src, label) in enumerate(self._TRACKING_TABLE_COLUMNS):
            action = menu.addAction(label)
            action.setCheckable(True)
            action.setChecked(not table.isColumnHidden(index))
            if label in TRACKING_ALWAYS_VISIBLE_COLUMNS:
                action.setEnabled(False)
            action.toggled.connect(lambda visible, name=label: self._set_tracking_column_visible(name, visible))
        menu.addSeparator()
        menu.addAction("기본 열로 되돌리기", self._reset_tracking_columns)
        menu.exec(table.horizontalHeader().mapToGlobal(pos))

    def _hidden_tracking_columns(self):
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return set()
        return {label for index, (_src, label) in enumerate(self._TRACKING_TABLE_COLUMNS)
                if index < table.columnCount() and table.isColumnHidden(index)}

    def _set_tracking_column_visible(self, name, visible):
        hidden = self._hidden_tracking_columns()
        (hidden.discard if visible else hidden.add)(name)
        self.set_app_setting(SETTING_TRACKING_HIDDEN_COLUMNS, sorted(hidden))
        self._apply_tracking_hidden_columns(hidden)
        self._fit_tracking_columns()

    def _reset_tracking_columns(self):
        self.set_app_setting(SETTING_TRACKING_HIDDEN_COLUMNS, list(TRACKING_DEFAULT_HIDDEN_COLUMNS))
        self._apply_tracking_hidden_columns(set(TRACKING_DEFAULT_HIDDEN_COLUMNS))
        self._fit_tracking_columns()

    def _apply_tracking_hidden_columns(self, hidden):
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return
        for index, (_src, label) in enumerate(self._TRACKING_TABLE_COLUMNS):
            if index < table.columnCount():
                table.setColumnHidden(index, label in hidden)

    def _fit_tracking_columns(self):
        """보이는 열을 표 폭에 맞추고 마지막 열에 남은 폭을 할당한다."""
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None or table.columnCount() == 0:
            return
        header = table.horizontalHeader()
        visible = [c for c in range(table.columnCount()) if not table.isColumnHidden(c)]
        if not visible:
            return
        *fixed, last = visible
        wanted = {c: min(max(table.sizeHintForColumn(c), header.sectionSizeHint(c)), TRACKING_COLUMN_MAX_WIDTH)
                  for c in fixed}
        floor = {c: min(header.sectionSizeHint(c), wanted[c]) for c in fixed}
        viewport_width = table.viewport().width()
        last_min = max(TRACKING_LAST_COLUMN_MIN_WIDTH, header.sectionSizeHint(last))
        budget = viewport_width - last_min
        widths = dict(wanted)
        if sum(widths.values()) > budget > 0:
            if sum(floor.values()) > budget:
                # 열을 더 켜거나 창을 좁히면 머리글 폭보다 줄여 표 폭을 맞춘다.
                floor = dict.fromkeys(fixed, header.minimumSectionSize())
            low, high = 0, max(wanted.values())
            while low < high:  # 합이 예산 안에 드는 가장 큰 상한을 찾는다
                mid = (low + high + 1) // 2
                if sum(max(floor[c], min(wanted[c], mid)) for c in fixed) <= budget:
                    low = mid
                else:
                    high = mid - 1
            widths = {c: max(floor[c], min(wanted[c], low)) for c in fixed}
        for c, width in widths.items():
            header.resizeSection(c, width)
        # Qt의 마지막 열 늘리기는 이전 섹션 폭(기본 100px)을 최소 폭으로 유지할 수 있다.
        header.resizeSection(last, max(last_min, viewport_width - sum(widths.values())))

    def _tracking_empty_text(self, mode, search):
        if getattr(self, "_tracking_risk_filter", None):
            return "조건에 맞는 위험 송장이 없습니다. 위험 칩을 다시 누르면 모두 보입니다."
        if search:
            return f"「{search}」 수취인을 찾지 못했습니다."
        return {
            "배송중": "배송중인 송장이 없습니다.",
            "오늘": "오늘 등록된 송장이 없습니다.",
            "완료": f"최근 {TRACKING_COMPLETED_LOOKBACK_DAYS}일 안에 완료된 송장이 없습니다.",
        }.get(mode, f"「{mode}」 송장이 없습니다.")

    def _show_tracking_empty_label(self, text):
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return
        label = getattr(self, "_tracking_empty_label", None)
        if label is None:
            label = QLabel(table.viewport())
            label.setObjectName("label_tracking_empty")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            label.setWordWrap(True)
            set_status(label, "idle", "")
            self._tracking_empty_label = label
        if text:
            set_status(label, "idle", text)
            label.show()
            self._place_tracking_empty_label()
        else:
            label.hide()

    def _place_tracking_empty_label(self):
        label = getattr(self, "_tracking_empty_label", None)
        table = getattr(getattr(self, "ui", None), "tableWidget_tracking", None)
        if label is None or table is None:
            return
        rect = table.viewport().rect()
        label.setGeometry(rect.adjusted(20, 0, -20, 0))

    @staticmethod
    def _last_checked_text(rows):
        """행의 최근조회 시각(J열) 중 가장 늦은 것. 다른 PC의 조회도 포함된다."""
        latest = None
        for row in rows:
            value = row[9] if len(row) > 9 else ""
            try:
                parsed = datetime.strptime(str(value).strip(), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
            latest = parsed if latest is None or parsed > latest else latest
        if latest is None:
            return ""
        return latest.strftime("%H:%M") if latest.date() == date.today() else latest.strftime("%m-%d %H:%M")

    def _update_tracking_risk_chips(self, counts):
        chips = getattr(self, "_tracking_risk_chips", None)
        if not chips:
            return
        total = sum(counts.values())
        current = getattr(self, "_tracking_risk_filter", None)
        for key, text in TRACKING_RISK_CHIPS:
            n = total if key == "all" else counts.get(key, 0)
            chip = chips[key]
            chip.setText(f"{text} {n}")
            chip.setVisible(n > 0)
            chip.setChecked(current == key)

    def show_help(self):
        """배송추적 사용 안내를 표시한다."""
        QMessageBox.information(
            self.host, "배송추적 사용 안내",
            "우체국 OpenAPI(종추적조회)로 공유 시트(「송장추적」 탭)의 배송 상태를 갱신합니다.\n"
            "송장 불러오기·일괄발송 시 등기번호가 자동으로 누적됩니다.\n"
            "종추적 인증키는 「⚙ 설정 › 연동」에서 등록합니다.\n\n"
            "※ 표에서 행을 더블클릭하면 해당 등기번호의 우체국 배송조회(상세) 웹페이지가 열립니다.",
        )

    def _enter_tracking_tab(self):
        """배송추적 탭 진입: 공유 키 상태 갱신 + 목록(표) 로드."""
        if gspread is not None:
            self._begin_tracking_config_read()  # 완료 후 _refresh_key_status 호출
        else:
            self._refresh_key_status()
        self._load_tracking_list()

    # ── 배송추적 목록(표) ────────────────────────────────────────────────
    # 표시 컬럼: (시트 컬럼 인덱스, 헤더 라벨)
    _TRACKING_TABLE_COLUMNS = [
        (4, "수취인명"), (6, "배송상태"), (8, "마지막위치"),
        (11, "최근이벤트"), (12, "관리상태"), (7, "완료"),
        (3, "주문번호"), (0, "등기번호"), (2, "스토어"),
        (9, "최근조회"), (10, "메모"),
    ]

    @classmethod
    def _tracking_table_col(cls, source_col):
        return next(i for i, (idx, _label) in enumerate(cls._TRACKING_TABLE_COLUMNS)
                    if idx == source_col)
    def _risk_settings(self):
        """정체 기준과 도서산간 가산시간을 설정 파일 한 번 읽기로 계산한다.

        표 갱신·다이제스트처럼 여러 행을 판정할 때 행마다 설정 JSON을 다시 열지 않도록
        호출부에서 한 번 만들어 모든 행에 재사용한다.
        """
        reader = getattr(self.host, "read_app_settings", None)
        if callable(reader):
            settings = reader()
            get = lambda key: settings.get(key, "")  # noqa: E731
        else:
            get = None
        thresholds = {
            category: self._stale_threshold(category, get)
            for category in STALE_DEFAULTS
        }
        return thresholds, self._remote_bonus_hours(get)

    def _evaluate_risk(self, status, where, done, ref, now_dt, management="",
                       risk_settings=None):
        """분류별 영업시간 기준으로 위험 여부와 경과 시간을 계산한다."""
        thresholds, remote_bonus = risk_settings or self._risk_settings()
        return evaluate_risk(
            status, where, done, ref, now_dt, thresholds,
            remote_bonus, management)

    def _on_tracking_list_poll(self):
        """주기 타이머: 배송추적 탭을 보고 있을 때만 목록을 자동 재로딩."""
        if gspread is None or self._tracking_list_thread is not None:
            return
        tw = getattr(self.ui, "tabWidget", None)
        if tw is None:
            return
        cur = tw.currentWidget()
        if cur is None or cur.objectName() != "tab_tracking":
            return
        self._load_tracking_list()

    def _on_tracking_context_menu(self, pos):
        """행 우클릭 메뉴: 단건 조회·웹조회·수동 추적 중지/재개."""
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return
        item = table.itemAt(pos)
        if item is not None:
            table.setCurrentCell(item.row(), self._tracking_table_col(0))
        regino = self._selected_tracking_regino()
        if not regino:
            return
        management = self._tracking_management_for_regino(regino)
        menu = QMenu(table)
        act_query = None
        if management != TRACKING_MANAGEMENT_MANUAL_STOP:
            act_query = menu.addAction("이 송장만 우체국 조회")
        act_web = menu.addAction("우체국 웹조회 열기")
        menu.addSeparator()
        manual_stop = management != TRACKING_MANAGEMENT_MANUAL_STOP
        act_management = menu.addAction("추적 중지" if manual_stop else "추적 재개")
        chosen = menu.exec(table.viewport().mapToGlobal(pos))
        if chosen == act_query:
            self._on_tracking_refresh_one_clicked()
        elif chosen == act_web:
            QDesktopServices.openUrl(QUrl(self._kpost_trace_web_url.format(regino=regino)))
        elif chosen == act_management:
            self._change_tracking_management([regino], manual_stop)

    def _set_tracking_count_text(self, text, detail="", error=False):
        """표 위 건수 라벨은 한 줄만 쓴다. 여러 줄 오류 원문은 툴팁에 둔다
        (그대로 넣으면 라벨 높이를 넘어 위아래가 잘려 보인다)."""
        label = getattr(self.ui, "label_tracking_count", None)
        if label is None:
            return
        if error:
            set_status(label, "error", text, detail, mark=False)
        else:
            clear_status(label)
            label.setText(text)
            label.setToolTip(detail)

    def _load_tracking_list(self):
        """현재 목록 모드에 필요한 송장추적 행만 백그라운드로 읽어온다."""
        if gspread is None or self._tracking_list_thread is not None:
            return
        self._set_tracking_count_text("목록 불러오는 중…")
        self._set_tracking_table_loading(True)
        mode = "배송중"
        if hasattr(self.ui, "comboBox_tracking_filter"):
            mode = self.ui.comboBox_tracking_filter.currentText()
        thread = TrackingListThread(mode, self)
        self._tracking_list_thread = thread
        thread.result_ready.connect(self._on_tracking_list_finished)
        thread.finished.connect(self._cleanup_tracking_list_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_list_finished(self, payload: dict):
        if not payload.get("ok"):
            error = str(payload.get("error", "") or "unknown error")
            print(f"! Tracking list load failed: {error}")
            first = next((line.strip() for line in error.splitlines() if line.strip()), error)
            more = " …" if len(error.strip().splitlines()) > 1 else ""
            self._set_tracking_count_text(
                f"목록 불러오기 실패: {first}{more}", detail=error, error=True
            )
            self._set_tracking_table_loading(False)
            return
        self._tracking_list_values = payload.get("values", []) or []
        self._tracking_list_total_rows = int(payload.get("total_rows", 0) or 0)
        self._tracking_list_loaded = True
        self._populate_tracking_table()
        self._set_tracking_table_loading(False)
        if self._slack_notify_after_reload:
            self._slack_notify_after_reload = False
            self._maybe_send_daily_digest()

    def _cleanup_tracking_list_thread(self):
        self._tracking_list_thread = None

    def _on_tracking_filter_changed(self, _index):
        """필터 변경 시 해당 모드의 상세 행만 다시 읽는다."""
        self._load_tracking_list()

    def _set_tracking_table_loading(self, loading: bool):
        """기존 표를 남긴 채 목록 전환 중임을 반투명 오버레이로 표시한다."""
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return
        viewport = table.viewport()
        overlay = getattr(self, "_tracking_loading_overlay", None)
        if overlay is None:
            overlay = QWidget(viewport)
            overlay.setObjectName("trackingLoadingOverlay")  # 모양은 테마
            overlay.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            message = QLabel("목록 불러오는 중…", overlay)
            message.setObjectName("trackingLoadingMessage")
            message.setAlignment(Qt.AlignmentFlag.AlignCenter)
            message.setFixedSize(220, 56)
            self._tracking_loading_overlay = overlay
            self._tracking_loading_message = message
        overlay.setGeometry(viewport.rect())
        message = self._tracking_loading_message
        message.move(
            (overlay.width() - message.width()) // 2,
            (overlay.height() - message.height()) // 2,
        )
        if loading:
            overlay.raise_()
            overlay.show()
        else:
            overlay.hide()

    def _schedule_tracking_table_refresh(self, *_args):
        """검색어·정체 기준이 빠르게 바뀌는 동안은 기다렸다가 마지막 변경 뒤 한 번만 표를 그린다."""
        timer = getattr(self.host, "_tracking_table_refresh_timer", None)
        if timer is None:
            timer = self._new_timer(single_shot=True)
            timer.setInterval(TRACKING_TABLE_REFRESH_DEBOUNCE_MS)
            timer.timeout.connect(self._populate_tracking_table)
            self._tracking_table_refresh_timer = timer
        timer.start()

    def _populate_tracking_table(self, *_args):
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return
        values = self._tracking_list_values
        data_rows = values[1:] if len(values) > 1 else []
        total = len(data_rows)

        mode = "배송중"
        if hasattr(self.ui, "comboBox_tracking_filter"):
            mode = self.ui.comboBox_tracking_filter.currentText()
        today = date.today().strftime("%Y-%m-%d")

        def _cell(row, idx):
            return row[idx] if idx < len(row) else ""

        if mode == "오늘":
            rows = [r for r in data_rows if _cell(r, 1).startswith(today)]
        elif mode == "전체":
            rows = list(data_rows)
        elif mode == "완료":
            rows = [r for r in data_rows if _cell(r, 7).strip().upper() == "Y"]
        elif mode == "추적 중지":
            rows = [r for r in data_rows
                    if _cell(r, TRACKING_MANAGEMENT_COL) == TRACKING_MANAGEMENT_MANUAL_STOP]
        elif mode == "폐기 후보":
            rows = [r for r in data_rows
                    if _cell(r, TRACKING_MANAGEMENT_COL) == TRACKING_MANAGEMENT_CANDIDATE]
        elif mode == "폐기":
            rows = [r for r in data_rows
                    if _cell(r, TRACKING_MANAGEMENT_COL) == TRACKING_MANAGEMENT_DISCARDED]
        else:  # 배송중(=미완료): 완료여부 != Y
            rows = [r for r in data_rows
                    if _cell(r, 7).strip().upper() != "Y"
                    and _cell(r, TRACKING_MANAGEMENT_COL) not in TRACKING_MANAGEMENT_EXCLUDED]
        search = ""
        if hasattr(self.ui, "lineEdit_tracking_recipient_search"):
            search = self.ui.lineEdit_tracking_recipient_search.text().strip().casefold()
        if search:
            rows = [r for r in rows if search in _cell(r, 4).casefold()]

        # 위험 판정: 미완료 + 마지막 이벤트(없으면 등록) 후 분류별 기준(영업시간) 무이동.
        now_dt = datetime.now()
        bg_hub = QColor(theme.TRACK_RISK_HIGH_BG)   # 허브 정체 = 빨강(위험)
        bg_warn = QColor(theme.TRACK_RISK_WARN_BG)  # 수거누락·이동정체 = 주황(주의)
        counts = {"허브정체": 0, "수거누락": 0, "이동정체": 0}
        risk_settings, hidden_columns, auto_refresh_min = self._tracking_view_settings()

        evaluated = []
        for row in rows:
            done = _cell(row, 7).strip().upper() == "Y"
            ref = self._parse_event_dt(_cell(row, 11)) or self._parse_event_dt(_cell(row, 1))
            risk, _elapsed = self._evaluate_risk(
                _cell(row, 6), _cell(row, 8), done, ref, now_dt,
                _cell(row, TRACKING_MANAGEMENT_COL), risk_settings)
            if risk:
                counts[risk] += 1
            evaluated.append((row, risk))
        # 위험 칩 거르기. 거르던 위험이 없어지면 거르기를 푼다.
        risk_filter = getattr(self, "_tracking_risk_filter", None)
        if risk_filter and not (sum(counts.values()) if risk_filter == "all" else counts.get(risk_filter, 0)):
            risk_filter = self._tracking_risk_filter = None
        if risk_filter == "all":
            shown = [(row, risk) for row, risk in evaluated if risk]
        elif risk_filter:
            shown = [(row, risk) for row, risk in evaluated if risk == risk_filter]
        else:
            shown = evaluated

        cols = self._TRACKING_TABLE_COLUMNS
        prev_regino = self._selected_tracking_regino()  # 자동 갱신 시 선택 유지용
        table.setSortingEnabled(False)
        self._populating_tracking_table = True
        try:
            table.clearContents()
            table.setColumnCount(len(cols))
            table.setHorizontalHeaderLabels([label for _, label in cols])
            table.setRowCount(len(shown))
            for r, (row, risk) in enumerate(shown):
                bg = bg_hub if risk == "허브정체" else (bg_warn if risk else None)
                for c, (src_idx, _label) in enumerate(cols):
                    text = _cell(row, src_idx)
                    item = QTableWidgetItem(text)
                    if src_idx != 10:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    if src_idx in TRACKING_TOOLTIP_COLUMNS and text:
                        item.setToolTip(text)  # 좁은 열에서 잘린 글자를 볼 수 있게
                    if bg is not None:
                        item.setBackground(bg)
                    table.setItem(r, c, item)
        finally:
            self._populating_tracking_table = False
        table.setSortingEnabled(True)
        self._apply_tracking_hidden_columns(hidden_columns)
        self._fit_tracking_columns()
        if prev_regino:  # 자동 갱신 후 이전 선택 행 복원
            for r in range(table.rowCount()):
                it = table.item(r, self._tracking_table_col(0))
                if it is not None and it.text().strip() == prev_regino:
                    table.setCurrentCell(r, self._tracking_table_col(0))
                    break
        total_risk = sum(counts.values())
        if hasattr(self.ui, "label_tracking_count"):
            if mode == "전체":
                txt = f"표시 {len(shown)}건 / 전체 {total}건"
            elif mode == "완료":
                txt = f"표시 {len(shown)}건 / 최근 {TRACKING_COMPLETED_LOOKBACK_DAYS}일 완료"
            else:
                txt = f"표시 {len(shown)}건"
            if risk_filter:
                txt += " (위험만)"
            last_checked = self._last_checked_text(data_rows)
            if last_checked:
                txt += f"  ·  마지막 조회 {last_checked}"
            txt += f"  ·  자동 {auto_refresh_min}분" if auto_refresh_min else "  ·  자동 꺼짐"
            detail = txt
            if total_risk:
                detail += (f"\n위험 {total_risk}건: 허브 {counts['허브정체']} · 수거누락 {counts['수거누락']}"
                           f" · 이동 {counts['이동정체']}")
            if total_risk and not getattr(self, "_tracking_risk_chips", None):
                # 위험 칩이 없는 화면(예: 테스트용 간단한 화면)에서는 예전처럼 건수 줄에 위험을 보인다.
                set_status(self.ui.label_tracking_count, "error", detail.replace("\n", "  ·  "), detail,
                           emphasis=True, mark=False)
            else:
                clear_status(self.ui.label_tracking_count)
                self.ui.label_tracking_count.setText(txt)
                self.ui.label_tracking_count.setToolTip(detail)
        self._update_tracking_risk_chips(counts)
        loaded = getattr(self, "_tracking_list_loaded", False)
        self._show_tracking_empty_label(self._tracking_empty_text(mode, search) if loaded and not shown else "")

    @staticmethod
    def _parse_event_dt(s):
        """'2026.06.26 02:45'(이벤트) / '2026-06-26 02:45:00'(등록·조회) 등을 파싱."""
        s = (s or "").strip()
        if not s:
            return None
        for fmt in ("%Y.%m.%d %H:%M", "%Y.%m.%d %H:%M:%S",
                    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return None

    def _selected_tracking_regino(self):
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return ""
        r = table.currentRow()
        if r < 0:
            return ""
        item = table.item(r, self._tracking_table_col(0))
        return item.text().strip() if item is not None else ""

    def _selected_tracking_reginos(self):
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None or table.selectionModel() is None:
            return []
        regino_col = self._tracking_table_col(0)
        return [table.item(index.row(), regino_col).text().strip()
                for index in table.selectionModel().selectedRows()
                if table.item(index.row(), regino_col) is not None]

    def _tracking_management_for_regino(self, regino):
        for row in self._tracking_list_values[1:]:
            if _normalize_tracking_no(row[0] if row else "") == regino:
                return row[TRACKING_MANAGEMENT_COL].strip() if len(row) > TRACKING_MANAGEMENT_COL else ""
        return ""

    def _on_tracking_stop_selected_clicked(self):
        self._change_tracking_management(self._selected_tracking_reginos(), True)

    def _change_tracking_management(self, reginos, stop):
        if gspread is None:
            QMessageBox.warning(
                self.host, "배송추적", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        if self._tracking_management_update_thread is not None:
            return
        reginos = list(dict.fromkeys(reginos))
        if not reginos:
            QMessageBox.information(
                self.host, "추적 중지", "먼저 표에서 행을 선택해 주세요.")
            return
        action = "중지" if stop else "재개"
        if QMessageBox.question(
                self.host, f"추적 {action}",
                f"선택한 {len(reginos)}건의 배송 추적을 {action}할까요?\n\n"
                "중지한 건은 자동 새로고침과 위험 알림에서 제외됩니다.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        management = TRACKING_MANAGEMENT_MANUAL_STOP if stop else TRACKING_MANAGEMENT_ACTIVE
        thread = TrackingManagementUpdateThread(reginos, management, self)
        self._tracking_management_update_thread = thread
        self._set_tracking_summary(f"선택 {len(reginos)}건 추적 {action} 중…")
        thread.result_ready.connect(self._on_tracking_management_update_finished)
        thread.finished.connect(self._cleanup_tracking_management_update_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_management_update_finished(self, payload):
        if payload.get("ok"):
            self._set_tracking_summary(f"추적 상태 변경 {payload.get('updated', 0)}건 완료")
            self._load_tracking_list()
        else:
            QMessageBox.warning(
                self.host, "배송추적", f"추적 상태를 변경하지 못했습니다.\n\n{payload.get('error', '')}")

    def _cleanup_tracking_management_update_thread(self):
        self._tracking_management_update_thread = None

    def _on_tracking_item_changed(self, item):
        table = getattr(self.ui, "tableWidget_tracking", None)
        if self._populating_tracking_table or table is None:
            return
        if table.column(item) != self._tracking_table_col(10):
            return
        regino_item = table.item(table.row(item), self._tracking_table_col(0))
        if regino_item is None:
            return
        regino = regino_item.text().strip()
        if not regino:
            return
        self._tracking_notes_pending[regino] = item.text()
        self._flush_tracking_notes()

    def _flush_tracking_notes(self):
        if self._tracking_notes_update_thread is not None or not self._tracking_notes_pending:
            return
        self._tracking_notes_inflight = self._tracking_notes_pending
        self._tracking_notes_pending = {}
        thread = TrackingNotesUpdateThread(self._tracking_notes_inflight, self)
        self._tracking_notes_update_thread = thread
        thread.result_ready.connect(self._on_tracking_notes_update_finished)
        thread.finished.connect(self._cleanup_tracking_notes_update_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_notes_update_finished(self, payload):
        if not payload.get("ok"):
            QMessageBox.warning(
                self.host, "배송추적", f"비고를 저장하지 못했습니다.\n\n{payload.get('error', '')}")

    def _cleanup_tracking_notes_update_thread(self):
        self._tracking_notes_update_thread = None
        self._tracking_notes_inflight = {}
        self._flush_tracking_notes()

    def _on_tracking_cell_double_clicked(self, row, col):
        """행 더블클릭 → 그 등기번호의 우체국 배송조회 웹페이지를 엽니다."""
        if col == self._tracking_table_col(10):
            return  # 비고 열은 인라인 편집용
        table = getattr(self.ui, "tableWidget_tracking", None)
        if table is None:
            return
        item = table.item(row, self._tracking_table_col(0))
        regino = item.text().strip() if item is not None else ""
        if regino:
            QDesktopServices.openUrl(QUrl(self._kpost_trace_web_url.format(regino=regino)))

    def _on_tracking_refresh_one_clicked(self):
        """선택한 행의 등기번호 1건만 우체국으로 재조회해 갱신합니다."""
        if gspread is None:
            QMessageBox.warning(
                self.host, "배송추적", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        if self._tracking_refresh_one_thread is not None:
            return
        regino = self._selected_tracking_regino()
        if not regino:
            QMessageBox.information(
                self.host, "선택 새로고침", "먼저 표에서 행(등기번호)을 선택해 주세요.")
            return
        regkey = self._get_kpost_regkey()
        btn = getattr(self.ui, "pushButton_tracking_refresh_one", None)
        if btn is not None:
            btn.setEnabled(False)
        self._set_tracking_summary(f"{regino} 조회 중…")
        thread = TrackingRefreshOneThread(regkey, regino, self)
        self._tracking_refresh_one_thread = thread
        thread.result_ready.connect(self._on_tracking_refresh_one_finished)
        thread.finished.connect(self._cleanup_tracking_refresh_one_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_refresh_one_finished(self, payload: dict):
        if payload.get("ok"):
            done = " (배달완료)" if payload.get("complete") else ""
            self._set_tracking_summary(f"{payload.get('regino', '')}: {payload.get('status', '')}{done}")
            self._load_tracking_list()
        else:
            QMessageBox.warning(
                self.host, "선택 새로고침",
                f"조회하지 못했습니다.\n\n{payload.get('error', '')}",
            )


    def _enqueue_tracking_registration(self, records):
        """송장 등기번호 등록 요청을 모아 디바운스 후 백그라운드 upsert.
        gspread 미설치·미인증이어도 앱은 비차단(조용히 보류/스킵)."""
        if not records:
            return
        self._tracking_pending_records.extend(records)
        self._tracking_push_timer.start(TRACKING_SHEET_PUSH_DEBOUNCE_MS)

    def _flush_tracking_records(self):
        if gspread is None:
            return  # 패키지 없으면 보류(레코드 유지). 다음 등록 시 함께 시도.
        if not self._tracking_pending_records:
            return
        if self._tracking_op_thread is not None:
            # 진행 중인 upsert가 있으면 잠시 후 재시도(유실 방지).
            self._tracking_push_timer.start(TRACKING_SHEET_PUSH_DEBOUNCE_MS)
            return
        records = self._tracking_pending_records
        self._tracking_pending_records = []
        self._tracking_inflight_records = records
        thread = TrackingUpsertThread(records, self)
        self._tracking_op_thread = thread
        thread.result_ready.connect(self._on_tracking_upsert_finished)
        thread.finished.connect(self._cleanup_tracking_op_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_upsert_finished(self, payload: dict):
        inflight = self._tracking_inflight_records
        self._tracking_inflight_records = []
        if payload.get("ok"):
            reg = payload.get("registered", 0)
            upd = payload.get("updated", 0)
            if reg or upd:
                print(f"✓ 송장추적 시트 반영: 신규 {reg} / 보강 {upd}")
        else:
            # 실패 → 다음 등록 기회에 함께 재시도하도록 되돌림(자동 재폴링은 하지 않음).
            self._tracking_pending_records = inflight + self._tracking_pending_records
            print(f"! 송장추적 시트 반영 실패: {payload.get('error', '')}")

    def _cleanup_tracking_op_thread(self):
        self._tracking_op_thread = None

    def export_courier_receipt(self):
        """오늘 등록한 송장을 택배사 접수목록 Excel로 내보낸다."""
        if gspread is None:
            QMessageBox.warning(
                self.host, "택배사 접수목록",
                "gspread 패키지가 필요합니다. (pip install gspread)",
            )
            return
        if self._courier_export_thread is not None:
            self.host.statusBar().showMessage("택배사 접수목록: 이미 내보내는 중입니다. 잠시만 기다려 주세요.", 5000)
            return
        today = date.today()
        default_name = f"{COURIER_RECEIPT_PREFIX} {today.strftime('%y.%m.%d')}.xlsx"
        start_path = str(Path.home() / "Documents" / default_name)
        path, _ = QFileDialog.getSaveFileName(
            self.host, "택배사 접수목록 저장", start_path, "Excel 파일 (*.xlsx)")
        if not path:
            return
        self.host.statusBar().showMessage("택배사 접수목록 생성 중…")
        thread = CourierReceiptExportThread(path, today.strftime("%Y-%m-%d"), self.host)
        self._courier_export_thread = thread
        thread.result_ready.connect(self._on_courier_export_finished)
        thread.finished.connect(self._cleanup_courier_export_thread)
        thread.start()

    def _on_courier_export_finished(self, payload):
        """택배사 접수목록 Excel 생성 결과를 안내한다."""
        if not payload.get("ok"):
            self.host.statusBar().showMessage("택배사 접수목록 생성 실패", 4000)
            QMessageBox.warning(
                self.host, "택배사 접수목록",
                payload.get("error", "알 수 없는 오류가 발생했습니다."),
            )
            return
        count = payload.get("count", 0)
        if count == 0:
            self.host.notify("warn", "택배사 접수목록: 오늘 송장추적에 등록된 송장이 없습니다.")
            return
        path = payload.get("path")
        # 「폴더를 열까요?」 확인 창 대신 알림 띠의 「엑셀 열기」·「폴더 열기」(계획 C-3).
        self.host.show_excel_created_message(path, f"택배사 접수목록 {count}건을 저장했습니다.")

    def _cleanup_courier_export_thread(self):
        thread = self._courier_export_thread
        self._courier_export_thread = None
        if thread is not None:
            thread.deleteLater()

    # ── 송장 배송추적: 새로고침(스마트택배 조회) ──────────────────────────
    def _set_tracking_summary(self, text):
        if hasattr(self.ui, "label_tracking_summary"):
            # 폭이 좁으면 글자가 잘리므로 원문은 툴팁으로도 보인다(라벨이 창 폭을 늘리지 않음).
            self.ui.label_tracking_summary.setText(text)
            self.ui.label_tracking_summary.setToolTip(text)

    def on_refresh_tracking_clicked(self):
        """화면의 추적중 송장만 우체국 종추적조회로 갱신합니다."""
        self._begin_tracking_refresh("active")

    def on_refresh_tracking_exceptions_clicked(self):
        """폐기후보·폐기(미발송) 송장만 재확인해 늦은 접수를 복구합니다."""
        self._begin_tracking_refresh("exceptions")

    def _begin_tracking_refresh(self, scope):
        if gspread is None:
            QMessageBox.warning(
                self.host, "배송추적", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        if self._tracking_refresh_thread is not None:
            return  # 이미 조회 중
        # 로컬 키가 없어도 진행: 워커가 공유 「설정」 탭에서 회사 공통 regkey 를 자동 조회한다.
        regkey = self._get_kpost_regkey()
        for name in ("pushButton_refresh_tracking", "pushButton_refresh_tracking_exceptions"):
            if hasattr(self.ui, name):
                getattr(self.ui, name).setEnabled(False)
        label = "배송중" if scope == "active" else "폐기/예외 재확인"
        self._tracking_refresh_label = label
        self._tracking_refresh_scope = scope
        self._set_tracking_summary(f"{label} 조회 중…")
        thread = TrackingRefreshThread(regkey, self, scope=scope)
        self._tracking_refresh_thread = thread
        thread.progress.connect(self._on_tracking_refresh_progress)
        thread.result_ready.connect(self._on_tracking_refresh_finished)
        thread.finished.connect(self._cleanup_tracking_refresh_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_refresh_progress(self, done: int, total: int):
        self._set_tracking_summary(
            f"{getattr(self, '_tracking_refresh_label', '배송추적')} 조회 중… ({done}/{total})")

    def _on_tracking_refresh_finished(self, payload: dict):
        if payload.get("ok"):
            total = payload.get("total", 0)
            skipped = payload.get("skipped", 0)
            skip_note = f" · 출고대기 {skipped}건 제외" if skipped else ""
            if total == 0:
                self._set_tracking_summary("조회할 송장이 없습니다." + skip_note)
            else:
                t = datetime.now().strftime("%H:%M:%S")
                msg = (
                    f"총 {total} / 완료 {payload.get('complete', 0)} / "
                    f"진행 {payload.get('progress', 0)} / 실패 {payload.get('failed', 0)} "
                    f"(갱신 {t})"
                )
                if payload.get("aborted"):
                    msg += " · 우체국 부하로 일부 중단(ERR-131)"
                msg += skip_note
                self._set_tracking_summary(msg)
            # 갱신된 상태를 표에 반영하고, 반영 후 정체 건 슬랙 알림 검토
            self._slack_notify_after_reload = self._tracking_refresh_scope == "active"
            self._load_tracking_list()
        else:
            err = payload.get("error", "")
            self._set_tracking_summary("배송추적 조회 실패")
            QMessageBox.warning(
                self.host, "배송추적",
                f"배송 상태를 조회하지 못했습니다.\n\n{err}\n\n{self._oauth_error_dialog_hint()}",
            )

    def _cleanup_tracking_refresh_thread(self):
        self._tracking_refresh_thread = None
        for name in ("pushButton_refresh_tracking", "pushButton_refresh_tracking_exceptions"):
            if hasattr(self.ui, name):
                getattr(self.ui, name).setEnabled(True)

    # ── 송장 배송추적: 자동 새로고침(백그라운드 타이머) ──────────────────────
    def _auto_refresh_interval_min(self):
        """자동 새로고침 간격(분). 미설정이면 기본 60, '0'이면 끔(0 반환)."""
        return self._parse_auto_refresh_min(self.get_app_setting(CONFIG_KEY_AUTO_REFRESH_MIN, ""))

    @staticmethod
    def _parse_auto_refresh_min(value):
        s = str(value if value is not None else "").strip()
        if not s:
            return TRACKING_AUTO_REFRESH_DEFAULT_MIN
        try:
            n = int(s)
        except (TypeError, ValueError):
            return TRACKING_AUTO_REFRESH_DEFAULT_MIN
        return n if n >= 0 else TRACKING_AUTO_REFRESH_DEFAULT_MIN

    def _reconfigure_auto_refresh_timer(self):
        """간격 설정에 맞춰 자동 새로고침 타이머를 켜고/끄고 재시작합니다."""
        timer = getattr(self, "_tracking_auto_refresh_timer", None)
        if timer is None:
            return
        minutes = self._auto_refresh_interval_min()
        if minutes <= 0:
            timer.stop()
            return
        timer.setInterval(minutes * 60000)
        timer.start()  # 새 간격으로 (재)시작

    def _on_tracking_auto_refresh_tick(self):
        """자동 새로고침 타이머: 미완료 송장을 우체국으로 조회(백그라운드, 팝업 없음).
        수동/자동이 이미 조회 중이거나 간격이 안 지났으면(워커가 판정) 조용히 건너뜁니다."""
        if gspread is None or self._tracking_refresh_thread is not None:
            return
        minutes = self._auto_refresh_interval_min()
        if minutes <= 0:
            return
        regkey = self._get_kpost_regkey()
        thread = TrackingRefreshThread(regkey, self, auto=True, interval_min=minutes)
        self._tracking_refresh_thread = thread
        thread.result_ready.connect(self._on_tracking_auto_refresh_finished)
        thread.finished.connect(self._cleanup_tracking_refresh_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_auto_refresh_finished(self, payload: dict):
        """자동 새로고침 결과(비대화형): 팝업 없이 요약만 갱신하고 표를 다시 읽습니다."""
        if not payload.get("ok"):
            # 자동 조회 실패는 조용히 로그만(다음 주기에 재시도). 사용자 흐름을 끊지 않는다.
            print(f"! 자동 배송추적 조회 실패: {payload.get('error', '')}")
            return
        if payload.get("skipped_recent"):
            return  # 다른 PC가 최근에 조회했거나 설정상 꺼짐 → 아무 것도 안 함
        total = payload.get("total", 0)
        if total:
            t = datetime.now().strftime("%H:%M:%S")
            self._set_tracking_summary(
                f"자동 갱신 {t} · 총 {total} / 완료 {payload.get('complete', 0)} / "
                f"진행 {payload.get('progress', 0)} / 실패 {payload.get('failed', 0)}"
            )
        # 갱신 결과를 표에 반영하고, 반영 후 정체 건 슬랙 다이제스트 검토(하루 1통, 시트로 중복 방지)
        self._slack_notify_after_reload = True
        self._load_tracking_list()

    def _sync_auto_refresh_spinbox(self):
        """설정 팝업의 자동 새로고침 간격 스핀박스를 현재 값으로 갱신(시그널 차단)."""
        box = getattr(self, "_dlg_auto_refresh", None)
        if box is not None:
            box.blockSignals(True)
            box.setValue(self._auto_refresh_interval_min())
            box.blockSignals(False)

    def _on_auto_refresh_interval_changed(self, value):
        """설정 팝업의 자동 새로고침 간격 변경 → 로컬 저장 + 공유 「설정」 반영 + 타이머 재설정."""
        self._push_tracking_config(CONFIG_KEY_AUTO_REFRESH_MIN, int(value))
        self._reconfigure_auto_refresh_timer()

    def _cleanup_tracking_refresh_one_thread(self):
        self._tracking_refresh_one_thread = None
        btn = getattr(self.ui, "pushButton_tracking_refresh_one", None)
        if btn is not None:
            btn.setEnabled(True)

    def _refresh_slack_status(self):
        url = str(self.get_app_setting("slack_webhook_url", "") or "").strip()
        self._slack_status_text = "설정됨 ✓" if url else "미설정"
        self._update_status_displays()

    def _on_slack_config_clicked(self):
        cur = str(self.get_app_setting("slack_webhook_url", "") or "")
        url, ok = QInputDialog.getText(
            self.host, "슬랙 웹훅 설정",
            "슬랙 Incoming Webhook URL을 붙여넣으세요.\n"
            "(https://hooks.slack.com/services/...  · 공유 시트에 저장되어 전원 적용)",
            QLineEdit.EchoMode.Normal, cur,
        )
        if not ok:
            return
        url = url.strip()
        self.set_app_setting("slack_webhook_url", url)
        # 공유 「설정」 탭에도 반영(전원 자동 적용)
        if url:
            self.queue_tracking_config_write({CONFIG_KEY_SLACK_WEBHOOK: url})
        self._refresh_slack_status()

    def _on_slack_auto_toggled(self, state):
        self.set_app_setting("slack_auto_notify", bool(state))

    def _on_slack_test_clicked(self):
        url = str(self.get_app_setting("slack_webhook_url", "") or "").strip()
        if not url:
            QMessageBox.information(
                self.host, "슬랙 테스트", "먼저 「슬랙 설정」에서 웹훅 URL을 등록해 주세요.")
            return
        t = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._send_slack_async(f"✅ Easy Fulfill 배송모니터링 테스트 메시지입니다. ({t})", is_test=True)

    def _send_slack_async(self, text, is_test=False):
        # 자동 알림은 토·일 미발송(평일 월~금만). 수동 테스트 버튼은 주말에도 허용.
        if not is_test and not _is_weekday():
            return
        url = str(self.get_app_setting("slack_webhook_url", "") or "").strip()
        if not url or self._slack_send_thread is not None:
            return
        thread = SlackSendThread(url, text, self)
        self._slack_send_thread = thread
        thread.result_ready.connect(lambda p: self._on_slack_sent(p, is_test))
        thread.finished.connect(self._cleanup_slack_send_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_slack_sent(self, payload: dict, is_test: bool):
        if payload.get("ok"):
            if is_test:
                self.host.notify("ok", "슬랙 테스트 메시지를 보냈습니다. 슬랙 채널을 확인해 주세요.")
            else:
                print("✓ 슬랙 알림 전송")
        else:
            err = payload.get("error", "")
            if is_test:
                QMessageBox.warning(
                    self.host, "슬랙 테스트", f"전송 실패:\n\n{err}")
            else:
                print(f"! 슬랙 알림 전송 실패: {err}")

    def _cleanup_slack_send_thread(self):
        self._slack_send_thread = None

    def _collect_risk_rows(self):
        """현재 목록에서 위험 건(상태별)을 모읍니다.
        반환: list[dict]. dict: regino,name,status,where,event_time,elapsed_h,category."""
        values = self._tracking_list_values
        data = values[1:] if len(values) > 1 else []
        now_dt = datetime.now()
        risk_settings = self._risk_settings()

        def _cell(row, idx):
            return row[idx] if idx < len(row) else ""

        out = []
        for row in data:
            done = _cell(row, 7).strip().upper() == "Y"
            ev = self._parse_event_dt(_cell(row, 11))
            ref = ev or self._parse_event_dt(_cell(row, 1))
            risk, elapsed_h = self._evaluate_risk(
                _cell(row, 6), _cell(row, 8), done, ref, now_dt,
                _cell(row, TRACKING_MANAGEMENT_COL), risk_settings)
            if not risk:
                continue
            out.append({
                "regino": _cell(row, 0),
                "name": _cell(row, 4),
                "status": _cell(row, 6),
                "where": _cell(row, 8),
                "event_time": _cell(row, 11) if ev else "",
                "elapsed_h": elapsed_h,
                "category": risk,
            })
        # 위험도 순서: 허브정체 → 수거누락 → 이동정체
        order = {"허브정체": 0, "수거누락": 1, "이동정체": 2}
        out.sort(key=lambda it: (order.get(it["category"], 9), -it["elapsed_h"]))
        return out

    @staticmethod
    def _risk_signature(risks):
        """위험 목록의 내용 서명을 계산해 같은 목록의 재발송을 막는다."""
        return risk_signature(risks)

    def _build_risk_digest_text(self, risks):
        """위험 목록을 Slack 일일 다이제스트 본문으로 변환한다."""
        thresholds, _remote_bonus = self._risk_settings()
        return build_risk_digest_text(risks, thresholds)

    def _maybe_send_daily_digest(self):
        """전체 새로고침 직후, 일일 알림이 켜져 있고 위험 건이 있으면 하루 1통 다이제스트.
        중복은 공유 「설정」 탭의 발송 날짜로 방지(오늘 이미 보냈으면 전송 안 함).
        토·일은 발송하지 않는다(평일 월~금만)."""
        if not bool(self.get_app_setting("slack_auto_notify", False)):
            return
        if not _is_weekday():  # 토·일은 자동 알림 미발송(평일 월~금만)
            return
        webhook = str(self.get_app_setting("slack_webhook_url", "") or "").strip()
        if not webhook or self._digest_thread is not None:
            return
        risks = self._collect_risk_rows()
        if not risks:
            return
        text = self._build_risk_digest_text(risks)
        sig = self._risk_signature(risks)
        today = datetime.now().strftime("%Y-%m-%d")
        thread = DigestSendThread(webhook, text, today, sig, self)
        self._digest_thread = thread
        thread.result_ready.connect(self._on_digest_sent)
        thread.finished.connect(self._cleanup_digest_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_digest_sent(self, payload: dict):
        if payload.get("sent"):
            print("✓ 일일 위험 다이제스트 슬랙 발송")
        elif payload.get("ok"):
            reason = payload.get("reason", "")
            if reason == "unchanged":
                print("· 다이제스트 생략: 직전 발송과 내용 동일")
            else:
                print("· 다이제스트 생략: 오늘 이미 발송됨")
        else:
            print(f"! 다이제스트 발송 실패: {payload.get('error', '')}")

    def _cleanup_digest_thread(self):
        self._digest_thread = None

    def _stale_threshold(self, category, get=None):
        """분류별 정체 판정 기준 시간(영업시간, 공유 설정 동기화된 로컬 값).
        값이 없으면 STALE_DEFAULTS(허브 12 / 수거누락 24 / 이동 48).
        get: 이미 읽어 둔 설정에서 키를 조회하는 함수(없으면 설정 파일을 읽는다)."""
        default = STALE_DEFAULTS.get(category, 12)
        try:
            key = STALE_CONFIG_KEYS[category]
            v = get(key) if get is not None else self.get_app_setting(key, "")
            if str(v).strip():
                return int(v)
        except (KeyError, TypeError, ValueError):
            pass
        return default

    def _remote_bonus_hours(self, get=None):
        """도서산간 가산시간(영업시간, 공유 설정 동기화된 로컬 값). 기본 24."""
        try:
            key = CONFIG_KEY_STALE_REMOTE_BONUS
            v = get(key) if get is not None else self.get_app_setting(key, "")
            if str(v).strip():
                return int(v)
        except (TypeError, ValueError):
            pass
        return STALE_REMOTE_BONUS_DEFAULT

    def _push_tracking_config(self, key, value):
        """설정 1건을 로컬 저장 + 표 갱신 + 공유 「설정」 탭에 반영."""
        self.set_app_setting(key, value)
        self._schedule_tracking_table_refresh()
        self.queue_tracking_config_write({key: value})

    def _on_stale_threshold_changed(self, category, value):
        """설정 팝업의 분류별 정체기준 변경 → 로컬 저장 + 공유 「설정」 탭 반영 + 표 갱신."""
        self._push_tracking_config(STALE_CONFIG_KEYS[category], int(value))

    def _on_remote_bonus_changed(self, value):
        """설정 팝업의 도서산간 가산시간 변경 → 로컬 저장 + 공유 「설정」 탭 반영 + 표 갱신."""
        self._push_tracking_config(CONFIG_KEY_STALE_REMOTE_BONUS, int(value))

    def _sync_stale_threshold_spinboxes(self):
        """설정 팝업의 분류별 정체기준·도서산간 가산 스핀박스를 현재 값으로 갱신(시그널 차단)."""
        for cat, sb in self._dlg_stale_boxes.items():
            sb.blockSignals(True)
            sb.setValue(self._stale_threshold(cat))
            sb.blockSignals(False)
        if self._dlg_remote_bonus is not None:
            self._dlg_remote_bonus.blockSignals(True)
            self._dlg_remote_bonus.setValue(self._remote_bonus_hours())
            self._dlg_remote_bonus.blockSignals(False)

    def _refresh_key_status(self):
        """저장된 우체국 인증키의 유효성을 백그라운드로 확인해 상태 텍스트만 갱신합니다."""
        key = self._get_kpost_regkey()
        if not key:
            self._set_key_status_text("인증키: 미설정")
            return
        if gspread is None:
            self._set_key_status_text("인증키: 설정됨 (오프라인 — 유효성 미확인)")
            return
        if self._tracking_key_validate_thread is not None:
            return
        self._set_key_status_text("인증키: 확인 중…")
        self._start_key_validation(key, mode="status")

    def _on_tracking_key_edit_clicked(self):
        """관리자용: 작은 입력 다이얼로그로 새 우체국 인증키를 받아 검증 후 저장합니다.
        일반 유저는 평소 키를 보거나 바꿀 필요가 없으므로 입력칸을 노출하지 않습니다."""
        if gspread is None:
            QMessageBox.warning(
                self.host, "인증키 변경", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        if self._tracking_key_validate_thread is not None:
            return
        new_key, ok = QInputDialog.getText(
            self.host, "우체국 OpenAPI 인증키 변경",
            "새 우체국 OpenAPI 인증키(regkey, 30자리)를 입력하세요.\n"
            "종추적 조회로 검증한 뒤 직원 전원에게 적용됩니다.",
            QLineEdit.EchoMode.Password, "",
        )
        if not ok:
            return
        new_key = new_key.strip()
        if not new_key:
            QMessageBox.warning(
                self.host, "인증키 변경", "키가 비어 있습니다.")
            return
        self._set_key_status_text("인증키: 검증 중…")
        self._start_key_validation(new_key, mode="save")

    def _start_key_validation(self, key, mode):
        self._key_validate_mode = mode
        self._key_validate_pending_key = key
        if self._dlg_key_btn is not None:
            self._dlg_key_btn.setEnabled(False)
        thread = TrackingKeyValidateThread(key, self)
        self._tracking_key_validate_thread = thread
        thread.result_ready.connect(self._on_key_validate_finished)
        thread.finished.connect(self._cleanup_key_validate_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _cleanup_key_validate_thread(self):
        self._tracking_key_validate_thread = None
        if self._dlg_key_btn is not None:
            self._dlg_key_btn.setEnabled(True)

    def _on_key_validate_finished(self, payload: dict):
        mode = getattr(self, "_key_validate_mode", "status")
        key = getattr(self, "_key_validate_pending_key", "")
        valid = bool(payload.get("ok") and payload.get("valid"))
        if mode == "save":
            if not valid:
                err = payload.get("error", "유효하지 않은 키")
                QMessageBox.warning(
                    self.host, "인증키 검증 실패",
                    f"이 인증키로 우체국 종추적조회에 실패했습니다.\n\n{err}\n\n"
                    "키를 다시 확인해 주세요. (저장하지 않았습니다)",
                )
                self._set_key_status_text("인증키: 검증 실패 — 저장하지 않음")
                return
            self._commit_kpost_regkey(key)
            self._set_key_status_text("인증키: 유효함 ✓ (저장됨 · 전원 적용)")
        else:  # status
            if valid:
                self._set_key_status_text("인증키: 유효함 ✓")
            else:
                self._set_key_status_text("인증키: 확인 실패 — 「변경」으로 다시 등록")

    def _commit_kpost_regkey(self, key):
        """검증된 우체국 regkey 를 로컬과 공유 「설정」 탭에 저장합니다."""
        self.set_app_setting("kpost_regkey", key)
        self.queue_tracking_config_write({CONFIG_KEY_KPOST_REGKEY: key})

    def queue_tracking_config_write(self, updates):
        """공유 「설정」 탭 쓰기를 대기열에 넣는다.

        쓰기가 진행 중이면 버리지 않고 키별 최신 값으로 모았다가 끝난 뒤 이어서 쓴다
        (스핀박스를 빠르게 바꿔도 마지막 값이 시트에 남는다).
        """
        if gspread is None:
            return
        pending = getattr(self.host, "_tracking_config_pending_writes", None)
        if pending is None:
            pending = self._tracking_config_pending_writes = {}
        pending.update({key: str(value) for key, value in updates.items()})
        self._start_tracking_config_write()

    def _start_tracking_config_write(self):
        pending = getattr(self.host, "_tracking_config_pending_writes", None)
        if self._tracking_config_write_thread is not None or not pending:
            return
        updates = dict(pending)
        pending.clear()
        thread = TrackingConfigWriteThread(updates, self)
        self._tracking_config_write_thread = thread
        thread.result_ready.connect(self._on_tracking_config_write_finished)
        thread.finished.connect(self._cleanup_tracking_config_write_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_config_write_finished(self, payload: dict):
        if payload.get("ok"):
            self._set_tracking_summary("공유 설정을 시트에 저장했습니다. (직원 전원 자동 적용)")
        else:
            print(f"! 공유 설정 저장 실패: {payload.get('error', '')}")

    def _cleanup_tracking_config_write_thread(self):
        self._tracking_config_write_thread = None
        self._start_tracking_config_write()  # 진행 중에 모인 변경을 이어서 쓴다.

    def _begin_tracking_config_read(self):
        """공유 「설정」 탭에서 회사 공통 키를 읽어 입력란·로컬에 반영합니다."""
        if gspread is None or self._tracking_config_read_thread is not None:
            return
        thread = TrackingConfigReadThread(self)
        self._tracking_config_read_thread = thread
        thread.result_ready.connect(self._on_tracking_config_read_finished)
        thread.finished.connect(self._cleanup_tracking_config_read_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_config_read_finished(self, payload: dict):
        if payload.get("ok"):
            regkey = str(payload.get("regkey", "") or "").strip()
            if regkey:
                # 공유 시트의 종추적 키를 로컬에 반영한다.
                self.set_app_setting("kpost_regkey", regkey)
            slack = str(payload.get("slack_webhook", "") or "").strip()
            if slack:
                self.set_app_setting("slack_webhook_url", slack)
            for pkey, akey in (("stale_hub", CONFIG_KEY_STALE_HUB),
                               ("stale_pickup", CONFIG_KEY_STALE_PICKUP),
                               ("stale_transit", CONFIG_KEY_STALE_TRANSIT),
                               ("stale_remote_bonus", CONFIG_KEY_STALE_REMOTE_BONUS)):
                sv = str(payload.get(pkey, "") or "").strip()
                if sv:
                    try:
                        self.set_app_setting(akey, int(sv))
                    except (TypeError, ValueError):
                        pass
            for pkey, akey in (("inquiry_work_start", "inquiry_work_start_hour"),
                               ("inquiry_work_end", "inquiry_work_end_hour")):
                hv = str(payload.get(pkey, "") or "").strip()
                if hv:
                    try:
                        h = int(hv)
                        if 0 <= h <= 23:
                            self.set_app_setting(akey, h)
                    except (TypeError, ValueError):
                        pass
            # 배송추적 자동 새로고침 간격(전원 공유). 빈 값이면 로컬 기본(60) 유지.
            iv = str(payload.get("auto_refresh_min", "") or "").strip()
            if iv:
                try:
                    n = int(iv)
                    if n >= 0:
                        self.set_app_setting(CONFIG_KEY_AUTO_REFRESH_MIN, n)
                except (TypeError, ValueError):
                    pass
        # 공유 설정 수신 후 상태 텍스트·정체기준·알림시간대·자동새로고침·표 갱신
        apply_keys = getattr(self.host, "_apply_shared_key_registration", None)
        if apply_keys is not None:
            apply_keys(payload.get("shared_keys") if payload.get("ok") else None)
        self._refresh_key_status()
        self._refresh_slack_status()
        self._sync_stale_threshold_spinboxes()
        self._sync_inquiry_hours_spinboxes()
        self._sync_auto_refresh_spinbox()
        self._reconfigure_auto_refresh_timer()
        # 목록 로딩 중에는 원본 행이 아직 없으므로 표를 다시 그리지 않는다.
        # 설정 읽기가 먼저 끝난 경우 "표시 0건"으로 깜빡이는 것을 막고,
        # 목록 완료 콜백에서 설정값을 포함해 한 번만 표시한다.
        if self._tracking_list_thread is None:
            self._populate_tracking_table()

    def _cleanup_tracking_config_read_thread(self):
        self._tracking_config_read_thread = None

    def build_admin_slack_controls(self, outer):
        """설정 › 알림 페이지에 배송 위험 Slack 알림 제어를 추가한다."""
        # 위험 일일 알림 토글은 이 PC 앱 설정. Webhook 주소(전원 공유)는 같은 페이지 위 묶음에 있다.
        gb_slack = QGroupBox("배송 위험 일일 알림 · 이 PC")
        sl = QVBoxLayout(gb_slack)
        srow = QHBoxLayout()
        self._dlg_slack_auto = QCheckBox("위험 건이 있으면 슬랙으로 하루 1통 보내기")
        self._dlg_slack_auto.setChecked(bool(self.get_app_setting("slack_auto_notify", False)))
        self._dlg_slack_auto.setToolTip(
            "이 PC에서 전체 새로고침 시 위험 건이 있으면 하루 1통만 슬랙으로 보냅니다(중복 방지)."
        )
        self._dlg_slack_auto.stateChanged.connect(self._on_slack_auto_toggled)
        btn_test = QPushButton("테스트")
        btn_test.setToolTip("슬랙 Webhook으로 테스트 메시지를 보냅니다.")
        btn_test.clicked.connect(self._on_slack_test_clicked)
        srow.addWidget(self._dlg_slack_auto)
        srow.addStretch(1)
        srow.addWidget(btn_test)
        sl.addLayout(srow)
        outer.addWidget(gb_slack)

    def build_admin_risk_controls(self, outer):
        """설정 › 배송추적 페이지에 배송 위험 기준과 자동 갱신 제어를 추가한다."""
        # 위험 판정 기준(전원 공유). 좁은 컬럼이라 2행으로: 정체기준 3종 / 도서산간 가산.
        gb_risk = QGroupBox("위험 판정 기준 · 전원 공유")
        rv = QVBoxLayout(gb_risk)
        caption = QLabel("분류별 정체시간(토·일 제외 영업시간 기준)")
        caption.setProperty("hint", "true")
        rv.addWidget(caption)
        _risk_tip = ("미완료 송장이 이 시간 이상 우체국 이벤트가 없으면 '정체'로 표시합니다. "
                     "토·일은 제외한 영업시간 기준이며, 변경 시 공유 시트에 저장되어 "
                     "전원에게 적용됩니다.")
        # 디자인 개편 4-6: 숫자 칸 4개를 두 칸씩 2줄 격자로(한 줄 3칸이면 좁은 창에서 가로 스크롤이 생겼다).
        grid = QGridLayout()
        grid.setHorizontalSpacing(theme.SPACE_SM)
        grid.setVerticalSpacing(theme.SPACE_SM)
        grid.setColumnStretch(4, 1)

        def put(index, caption, spin):
            row, column = divmod(index, 2)
            label = QLabel(caption)
            label.setProperty("formLabel", True)
            grid.addWidget(label, row, column * 2)
            grid.addWidget(spin, row, column * 2 + 1)

        self._dlg_stale_boxes = {}
        for index, (cat, cap) in enumerate((("허브정체", "허브"), ("수거누락", "수거누락"), ("이동정체", "이동"))):
            sb = QSpinBox()
            sb.setMinimum(1)
            sb.setMaximum(168)
            sb.setValue(self._stale_threshold(cat))
            sb.setSuffix("h")
            sb.setToolTip(_risk_tip)
            sb.valueChanged.connect(
                lambda v, c=cat: self._on_stale_threshold_changed(c, v))
            put(index, cap, sb)
            self._dlg_stale_boxes[cat] = sb
        self._dlg_remote_bonus = QSpinBox()
        self._dlg_remote_bonus.setMinimum(0)
        self._dlg_remote_bonus.setMaximum(168)
        self._dlg_remote_bonus.setValue(self._remote_bonus_hours())
        self._dlg_remote_bonus.setSuffix("h")
        self._dlg_remote_bonus.setToolTip(
            "마지막위치가 제주·울릉 등 도서산간이면 이동정체 기준에 이 시간을 더합니다. "
            "(종추적 API엔 목적지가 없어 현재 위치 텍스트로 판별) 0이면 가산 없음.")
        self._dlg_remote_bonus.valueChanged.connect(self._on_remote_bonus_changed)
        put(3, "도서산간 가산", self._dlg_remote_bonus)
        rv.addLayout(grid)
        outer.addWidget(gb_risk)

        # 배송추적 자동 새로고침 간격(전원 공유). 0이면 자동 새로고침 끔.
        gb_auto = QGroupBox("자동 새로고침 · 전원 공유")
        arow = QHBoxLayout(gb_auto)
        arow.addWidget(QLabel("간격:"))
        self._dlg_auto_refresh = QSpinBox()
        self._dlg_auto_refresh.setRange(0, 720)  # 0=끔 ~ 12시간
        self._dlg_auto_refresh.setSuffix("분")
        self._dlg_auto_refresh.setSpecialValueText("끔")  # 0 → '끔' 표시
        self._dlg_auto_refresh.setValue(self._auto_refresh_interval_min())
        self._dlg_auto_refresh.setToolTip(
            "N분마다 미완료 송장을 우체국으로 백그라운드 조회해 배송상태를 갱신합니다.\n"
            "0이면 자동 새로고침을 끕니다(수동 새로고침만). 기본 60분.\n"
            "여러 대를 켜둬도 공유 시트의 마지막 조회시각으로 중복 조회를 막습니다"
            "(간격이 지난 뒤 먼저 도는 1대만 실제 조회). 변경 시 전원에게 적용됩니다.")
        self._dlg_auto_refresh.valueChanged.connect(self._on_auto_refresh_interval_changed)
        arow.addWidget(self._dlg_auto_refresh)
        arow.addStretch(1)
        outer.addWidget(gb_auto)
