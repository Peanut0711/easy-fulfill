"""배송추적 탭의 UI 제어와 화면 상태를 담당한다.

이 모듈은 Qt widget 신호를 연결하고 worker 결과를 화면에 반영한다.
Google Sheets/KPOST I/O는 tracking.workers 및 tracking.repository에 위임한다.
"""

from __future__ import annotations

from datetime import date, datetime

try:
    import gspread
except ImportError:
    gspread = None

from PySide6.QtCore import QObject, Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import QMenu, QMessageBox, QTableWidgetItem

from .repository import normalize_tracking_no as _normalize_tracking_no
from .service import (
    STALE_DEFAULTS,
    TRACKING_MANAGEMENT_ACTIVE,
    TRACKING_MANAGEMENT_CANDIDATE,
    TRACKING_MANAGEMENT_COL,
    TRACKING_MANAGEMENT_DISCARDED,
    TRACKING_MANAGEMENT_EXCLUDED,
    TRACKING_MANAGEMENT_MANUAL_STOP,
    evaluate_risk,
)
from .workers import (
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


class TrackingController(QObject):
    """MainWindow에 보관된 상태와 배송추적 UI의 연결을 관리한다."""

    def __init__(self, host):
        super().__init__(host)
        self.host = host
        self.ui = getattr(host, "ui", None)

    def __getattr__(self, name):
        return getattr(self.host, name)

    def __setattr__(self, name, value):
        if name in {"host", "ui"}:
            super().__setattr__(name, value)
        elif name.startswith("_tracking_") or name in {
            "_populating_tracking_table",
            "_slack_notify_after_reload",
        }:
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
                self._populate_tracking_table
            )
        if hasattr(self.ui, "lineEdit_tracking_recipient_search"):
            self.ui.lineEdit_tracking_recipient_search.textChanged.connect(
                self._populate_tracking_table
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
        """공유 시트에서 송장추적 목록을 백그라운드로 읽어옵니다."""
        if gspread is None or self._tracking_list_thread is not None:
            return
        if hasattr(self.ui, "label_tracking_count"):
            self.ui.label_tracking_count.setText("목록 불러오는 중…")
        thread = TrackingListThread(self)
        self._tracking_list_thread = thread
        thread.result_ready.connect(self._on_tracking_list_finished)
        thread.finished.connect(self._cleanup_tracking_list_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_tracking_list_finished(self, payload: dict):
        if not payload.get("ok"):
            if hasattr(self.ui, "label_tracking_count"):
                self.ui.label_tracking_count.setText("목록 불러오기 실패")
            return
        self._tracking_list_values = payload.get("values", []) or []
        self._populate_tracking_table()
        if self._slack_notify_after_reload:
            self._slack_notify_after_reload = False
            self._maybe_send_daily_digest()

    def _cleanup_tracking_list_thread(self):
        self._tracking_list_thread = None

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
            txt = f"표시 {len(rows)}건 / 전체 {total}건"
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
