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

try:
    import gspread
except ImportError:
    gspread = None

from PySide6.QtCore import QObject, Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
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
CONFIG_KEY_AUTO_REFRESH_MIN = "tracking_auto_refresh_min"
TRACKING_AUTO_REFRESH_DEFAULT_MIN = 60
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
        if hasattr(self.ui, "comboBox_tracking_filter"):
            self.ui.comboBox_tracking_filter.currentIndexChanged.connect(
                self._on_tracking_filter_changed
            )
        if hasattr(self.ui, "lineEdit_tracking_recipient_search"):
            self.ui.lineEdit_tracking_recipient_search.textChanged.connect(
                self._populate_tracking_table
            )

    def show_help(self):
        """배송추적 사용 안내를 표시한다."""
        QMessageBox.information(
            self.host, "배송추적 사용 안내",
            "우체국 OpenAPI(종추적조회)로 공유 시트(「송장추적」 탭)의 배송 상태를 갱신합니다.\n"
            "송장 불러오기·일괄발송 시 등기번호가 자동으로 누적됩니다.\n"
            "(우편번호 검색과 동일한 우체국 인증키 사용)\n\n"
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
    def _evaluate_risk(self, status, where, done, ref, now_dt, management=""):
        """분류별 영업시간 기준으로 위험 여부와 경과 시간을 계산한다."""
        thresholds = {
            category: self._stale_threshold(category)
            for category in STALE_DEFAULTS
        }
        return evaluate_risk(
            status, where, done, ref, now_dt, thresholds,
            self._remote_bonus_hours(), management)

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

    def _load_tracking_list(self):
        """현재 목록 모드에 필요한 송장추적 행만 백그라운드로 읽어온다."""
        if gspread is None or self._tracking_list_thread is not None:
            return
        if hasattr(self.ui, "label_tracking_count"):
            self.ui.label_tracking_count.setText("목록 불러오는 중…")
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
            if hasattr(self.ui, "label_tracking_count"):
                self.ui.label_tracking_count.setText(f"목록 불러오기 실패: {error}")
            self._set_tracking_table_loading(False)
            return
        self._tracking_list_values = payload.get("values", []) or []
        self._tracking_list_total_rows = int(payload.get("total_rows", 0) or 0)
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
            overlay.setStyleSheet(
                "background-color: rgba(255, 255, 255, 175);"
            )
            message = QLabel("목록 불러오는 중…", overlay)
            message.setAlignment(Qt.AlignmentFlag.AlignCenter)
            message.setFixedSize(220, 56)
            message.setStyleSheet(
                "background-color: rgba(255, 255, 255, 245);"
                "border: 1px solid #d5dbe3; border-radius: 8px;"
                "color: #333; font-size: 13px; font-weight: bold;"
            )
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
        bg_hub = QColor("#ffb3b3")      # 허브 정체 = 빨강(위험)
        bg_warn = QColor("#ffe0b3")     # 수거누락·이동정체 = 주황(주의)
        counts = {"허브정체": 0, "수거누락": 0, "이동정체": 0}

        cols = self._TRACKING_TABLE_COLUMNS
        prev_regino = self._selected_tracking_regino()  # 자동 갱신 시 선택 유지용
        table.setSortingEnabled(False)
        self._populating_tracking_table = True
        try:
            table.clearContents()
            table.setColumnCount(len(cols))
            table.setHorizontalHeaderLabels([label for _, label in cols])
            table.setRowCount(len(rows))
            for r, row in enumerate(rows):
                done = _cell(row, 7).strip().upper() == "Y"
                ref = self._parse_event_dt(_cell(row, 11)) or self._parse_event_dt(_cell(row, 1))
                risk, _elapsed = self._evaluate_risk(
                    _cell(row, 6), _cell(row, 8), done, ref, now_dt,
                    _cell(row, TRACKING_MANAGEMENT_COL))
                if risk:
                    counts[risk] += 1
                bg = bg_hub if risk == "허브정체" else (bg_warn if risk else None)
                for c, (src_idx, _label) in enumerate(cols):
                    item = QTableWidgetItem(_cell(row, src_idx))
                    if src_idx != 10:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    if bg is not None:
                        item.setBackground(bg)
                    table.setItem(r, c, item)
        finally:
            self._populating_tracking_table = False
        table.setSortingEnabled(True)
        table.resizeColumnsToContents()
        header = table.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(True)
        if prev_regino:  # 자동 갱신 후 이전 선택 행 복원
            for r in range(table.rowCount()):
                it = table.item(r, self._tracking_table_col(0))
                if it is not None and it.text().strip() == prev_regino:
                    table.setCurrentCell(r, self._tracking_table_col(0))
                    break
        if hasattr(self.ui, "label_tracking_count"):
            total_risk = sum(counts.values())
            if mode == "전체":
                txt = f"표시 {len(rows)}건 / 전체 {total}건"
            elif mode == "완료":
                txt = f"표시 {len(rows)}건 / 최근 {TRACKING_COMPLETED_LOOKBACK_DAYS}일 완료"
            else:
                txt = f"표시 {len(rows)}건"
            if total_risk:
                txt += (f"  ·  ⚠️ 위험 {total_risk}건"
                        f" (허브 {counts['허브정체']} · 수거누락 {counts['수거누락']}"
                        f" · 이동 {counts['이동정체']})")
                self.ui.label_tracking_count.setStyleSheet("color:#c0392b; font-weight:bold;")
            else:
                self.ui.label_tracking_count.setStyleSheet("")
            self.ui.label_tracking_count.setText(txt)

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
            QMessageBox.information(
                self.host, "택배사 접수목록", "이미 내보내는 중입니다. 잠시만 기다려 주세요.")
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
            self.host.statusBar().showMessage("오늘 접수된 송장이 없습니다.", 4000)
            QMessageBox.information(self.host, "택배사 접수목록", "오늘 송장추적에 등록된 송장이 없습니다.")
            return
        path = payload.get("path")
        self.host.statusBar().showMessage(f"택배사 접수목록 {count}건 저장 완료", 5000)
        result = QMessageBox.question(
            self.host, "택배사 접수목록",
            f"오늘 접수 {count}건을 저장했습니다.\n\n{path}\n\n파일이 있는 폴더를 열까요?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if result != QMessageBox.StandardButton.Yes:
            return
        try:
            folder = os.path.dirname(path)
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.call(("open", folder))
            else:
                subprocess.call(("xdg-open", folder))
        except Exception as error:
            print(f"! 택배사 접수목록 폴더 열기 실패: {error}")

    def _cleanup_courier_export_thread(self):
        thread = self._courier_export_thread
        self._courier_export_thread = None
        if thread is not None:
            thread.deleteLater()

    # ── 송장 배송추적: 새로고침(스마트택배 조회) ──────────────────────────
    def _set_tracking_summary(self, text):
        if hasattr(self.ui, "label_tracking_summary"):
            self.ui.label_tracking_summary.setText(text)

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
        s = str(self.get_app_setting(CONFIG_KEY_AUTO_REFRESH_MIN, "")).strip()
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
        if gspread is not None and url and self._tracking_config_write_thread is None:
            thread = TrackingConfigWriteThread({CONFIG_KEY_SLACK_WEBHOOK: url}, self)
            self._tracking_config_write_thread = thread
            thread.result_ready.connect(self._on_tracking_config_write_finished)
            thread.finished.connect(self._cleanup_tracking_config_write_thread)
            thread.finished.connect(thread.deleteLater)
            thread.start()
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
                QMessageBox.information(
                    self.host, "슬랙 테스트", "전송 성공! 슬랙 채널을 확인해 주세요.")
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

        def _cell(row, idx):
            return row[idx] if idx < len(row) else ""

        out = []
        for row in data:
            done = _cell(row, 7).strip().upper() == "Y"
            ev = self._parse_event_dt(_cell(row, 11))
            ref = ev or self._parse_event_dt(_cell(row, 1))
            risk, elapsed_h = self._evaluate_risk(
                _cell(row, 6), _cell(row, 8), done, ref, now_dt,
                _cell(row, TRACKING_MANAGEMENT_COL))
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
        thresholds = {
            category: self._stale_threshold(category)
            for category in STALE_DEFAULTS
        }
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

    def _stale_threshold(self, category):
        """분류별 정체 판정 기준 시간(영업시간, 공유 설정 동기화된 로컬 값).
        값이 없으면 STALE_DEFAULTS(허브 12 / 수거누락 24 / 이동 48)."""
        default = STALE_DEFAULTS.get(category, 12)
        try:
            v = self.get_app_setting(STALE_CONFIG_KEYS[category], "")
            if str(v).strip():
                return int(v)
        except (KeyError, TypeError, ValueError):
            pass
        return default

    def _remote_bonus_hours(self):
        """도서산간 가산시간(영업시간, 공유 설정 동기화된 로컬 값). 기본 24."""
        try:
            v = self.get_app_setting(CONFIG_KEY_STALE_REMOTE_BONUS, "")
            if str(v).strip():
                return int(v)
        except (TypeError, ValueError):
            pass
        return STALE_REMOTE_BONUS_DEFAULT

    def _push_tracking_config(self, key, value):
        """설정 1건을 로컬 저장 + 표 갱신 + 공유 「설정」 탭에 반영."""
        self.set_app_setting(key, value)
        self._populate_tracking_table()
        if gspread is not None and self._tracking_config_write_thread is None:
            thread = TrackingConfigWriteThread({key: str(value)}, self)
            self._tracking_config_write_thread = thread
            thread.result_ready.connect(self._on_tracking_config_write_finished)
            thread.finished.connect(self._cleanup_tracking_config_write_thread)
            thread.finished.connect(thread.deleteLater)
            thread.start()

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
            self._set_key_status_text("인증키: 미설정 — 「키 변경」으로 등록하세요.")
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
            "(검증 후 직원 전원에게 적용됩니다 · 우편번호 검색에도 함께 사용)",
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
                self._set_key_status_text("인증키: 확인 실패 — 「키 변경」으로 다시 등록")

    def _commit_kpost_regkey(self, key):
        """검증된 우체국 regkey 를 로컬과 공유 「설정」 탭에 저장합니다."""
        self.set_app_setting("kpost_regkey", key)
        if gspread is None or self._tracking_config_write_thread is not None:
            return
        thread = TrackingConfigWriteThread({CONFIG_KEY_KPOST_REGKEY: key}, self)
        self._tracking_config_write_thread = thread
        thread.result_ready.connect(self._on_tracking_config_write_finished)
        thread.finished.connect(self._cleanup_tracking_config_write_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_config_write_finished(self, payload: dict):
        if payload.get("ok"):
            self._set_tracking_summary("인증키를 공유 시트에 저장했습니다. (직원 전원 자동 적용)")
        else:
            print(f"! 공유 설정 저장 실패: {payload.get('error', '')}")

    def _cleanup_tracking_config_write_thread(self):
        self._tracking_config_write_thread = None

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
                # 공유 시트의 회사 공통 키를 로컬에 반영(우편번호 검색도 함께 사용)
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
        """관리자 탭에 배송 위험 Slack 알림 제어를 추가한다."""
        # 슬랙 위험 알림(체크는 PC별, Webhook 설정은 「연동 · API 키」에 있음)
        gb_slack = QGroupBox("슬랙 위험 알림")
        sl = QVBoxLayout(gb_slack)
        srow = QHBoxLayout()
        self._dlg_slack_auto = QCheckBox("이 PC에서 위험 일일 알림(하루 1통)")
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
        """관리자 탭에 배송 위험 기준과 자동 갱신 제어를 추가한다."""
        # 위험 판정 기준(전원 공유). 좁은 컬럼이라 2행으로: 정체기준 3종 / 도서산간 가산.
        gb_risk = QGroupBox("위험 판정 기준 (분류별 정체시간, 영업시간)")
        rv = QVBoxLayout(gb_risk)
        _risk_tip = ("미완료 송장이 이 시간 이상 우체국 이벤트가 없으면 '정체'로 표시합니다. "
                     "토·일은 제외한 영업시간 기준이며, 변경 시 공유 시트에 저장되어 "
                     "전원에게 적용됩니다.")
        rrow = QHBoxLayout()
        self._dlg_stale_boxes = {}
        for cat, cap in (("허브정체", "허브"), ("수거누락", "수거누락"), ("이동정체", "이동")):
            rrow.addWidget(QLabel(f"{cap}:"))
            sb = QSpinBox()
            sb.setMinimum(1)
            sb.setMaximum(168)
            sb.setValue(self._stale_threshold(cat))
            sb.setSuffix("h")
            sb.setToolTip(_risk_tip)
            sb.valueChanged.connect(
                lambda v, c=cat: self._on_stale_threshold_changed(c, v))
            rrow.addWidget(sb)
            self._dlg_stale_boxes[cat] = sb
        rrow.addStretch(1)
        rv.addLayout(rrow)
        brow = QHBoxLayout()
        brow.addWidget(QLabel("도서산간 가산:"))
        self._dlg_remote_bonus = QSpinBox()
        self._dlg_remote_bonus.setMinimum(0)
        self._dlg_remote_bonus.setMaximum(168)
        self._dlg_remote_bonus.setValue(self._remote_bonus_hours())
        self._dlg_remote_bonus.setSuffix("h")
        self._dlg_remote_bonus.setToolTip(
            "마지막위치가 제주·울릉 등 도서산간이면 이동정체 기준에 이 시간을 더합니다. "
            "(종추적 API엔 목적지가 없어 현재 위치 텍스트로 판별) 0이면 가산 없음.")
        self._dlg_remote_bonus.valueChanged.connect(self._on_remote_bonus_changed)
        brow.addWidget(self._dlg_remote_bonus)
        brow.addStretch(1)
        rv.addLayout(brow)
        outer.addWidget(gb_risk)

        # 배송추적 자동 새로고침 간격(전원 공유). 0이면 자동 새로고침 끔.
        gb_auto = QGroupBox("배송추적 자동 새로고침 (전원 공유)")
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
