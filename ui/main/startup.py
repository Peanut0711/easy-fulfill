"""메인 창 시작·종료: 화면 불러오기, 창 위치·크기 기억, 이 PC 앱 설정, 앱 버전 확인·업데이트 안내, 닫을 때 남은 시트 쓰기 마무리."""

import html
import json
import sys
import time
from pathlib import Path

from PySide6.QtCore import QFile, QIODevice, QRect, Qt, QTimer
from PySide6.QtGui import QGuiApplication
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QGridLayout, QLabel, QLayout, QMessageBox, QStyle,
)

from shared import app_log
from shared.lazy_import import preload_modules_in_background
from shared.paths import APP_ROOT
from shared.ui_notice import NoticeBar
from shared.ui_style import set_status
from shared.window_geometry import (
    centered_position,
    default_window_size,
    geometry_to_setting,
    is_restorable,
    parse_saved_geometry,
)
from sheets.version_gate import APP_VERSION, SOURCETREE_UPDATE_ERROR, SOURCETREE_UPDATE_STEPS
from shiboken6 import isValid
from ui.main import host as _host


class StartupMixin:
    """MainWindow 믹스인. 메인 창 시작·종료: 화면 불러오기, 창 위치·크기 기억, 이 PC 앱 설정, 앱 버전 확인·업데이트 안내, 닫을 때 남은 시트 쓰기 마무리."""

    def read_app_settings(self):
        """database/app_settings.json 전체를 한 번 읽어 dict로 반환합니다.

        여러 키를 연달아 쓰는 화면 갱신에서 키마다 파일을 다시 열지 않도록 사용합니다.
        파일이 없거나 손상되면 빈 dict를 반환해 각 키가 기본값을 쓰게 합니다.
        """
        try:
            if self.app_settings_path.exists() and self.app_settings_path.stat().st_size > 0:
                with open(self.app_settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception as e:
            print(f"! 앱 설정 읽기 오류: {e}")
        return {}

    def get_app_setting(self, key, default=None):
        """database/app_settings.json 에서 단일 설정값을 읽습니다."""
        try:
            if self.app_settings_path.exists() and self.app_settings_path.stat().st_size > 0:
                with open(self.app_settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and key in data:
                    return data[key]
        except Exception as e:
            print(f"! 앱 설정 읽기 오류({key}): {e}")
        return default

    def set_app_setting(self, key, value):
        """database/app_settings.json 에 단일 설정값을 병합 저장합니다."""
        try:
            self.app_settings_path.parent.mkdir(parents=True, exist_ok=True)
            data = {}
            if self.app_settings_path.exists() and self.app_settings_path.stat().st_size > 0:
                try:
                    with open(self.app_settings_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except json.JSONDecodeError:
                    data = {}
            if not isinstance(data, dict):
                data = {}
            data[key] = value
            with open(self.app_settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"! 앱 설정 저장 오류({key}): {e}")

    def load_ui(self):
        """UI 파일을 로드합니다."""
        # 옮기기 전 Path(__file__).parent(easy-fulfill.py 폴더)와 같은 곳
        ui_file_path = APP_ROOT / "ui" / "main_window.ui"
        ui_file = QFile(str(ui_file_path))
        
        if not ui_file.open(QIODevice.ReadOnly):
            print(f"UI 파일을 열 수 없음: {ui_file_path}")
            QMessageBox.critical(self, "오류", f"UI 파일을 열 수 없습니다: {ui_file_path}")
            sys.exit(1)
            
        loader = QUiLoader()
        window = loader.load(ui_file)
        ui_file.close()
        
        if not window:
            print(f"UI 로드 실패: {loader.errorString()}")
            QMessageBox.critical(self, "오류", f"UI 파일을 로드할 수 없습니다: {loader.errorString()}")
            sys.exit(1)
            
        # UI 객체를 인스턴스 변수로 저장
        self.ui = window
        self._setup_order_preview()
        self._mount_detail_page_tab()
        self._mount_inquiries_tab()
        
        # 메인 윈도우 설정
        self.setCentralWidget(window.centralwidget)
        self._setup_tab_shortcuts(window.tabWidget)
        # 탭 아래·상태바 위의 알림 띠(완료·주의 알림을 확인 창 대신 보인다). 평소에는 숨어 있다.
        self.notice_bar = NoticeBar(window.centralwidget)
        window.centralwidget.layout().addWidget(self.notice_bar)
        self._apply_tab_order()
        self._setup_navigation_rail(window.tabWidget)
        # self.setMenuBar(window.menubar)
        
        # 툴바 설정
        self.setup_menubar()
        
        self.setStatusBar(window.statusbar)
        self._setup_status_chips()
        self.setWindowTitle(f"{window.windowTitle()} · v{APP_VERSION}")
        self._order_ship_splitter_initialized = False

    def _apply_initial_window_geometry(self):
        """이 PC에 저장한 창 위치를 되살린다. 없거나 화면 구성이 바뀌어 쓸 수 없으면
        작업 영역에 맞춘 기본 크기(FHD 100%에서 900×800)로 주 화면 가운데에 둔다."""
        screen = QGuiApplication.primaryScreen()
        if screen is None:
            return
        hint = self.minimumSizeHint()
        minimum = (hint.width(), hint.height())
        saved = parse_saved_geometry(self.get_app_setting(_host(self).WINDOW_GEOMETRY_SETTING_KEY))
        if saved is not None:
            rect, maximized = saved
            areas = []
            for item in QGuiApplication.screens():
                area = item.availableGeometry()
                areas.append((area.x(), area.y(), area.width(), area.height()))
            if is_restorable(rect, areas, minimum):
                self.setGeometry(QRect(*rect))
                if maximized:
                    self.setWindowState(self.windowState() | Qt.WindowState.WindowMaximized)
                return
        self.setGeometry(self._default_window_rect(screen))

    def _default_window_rect(self, screen):
        """``screen`` 작업 영역에 맞춘 기본 창 크기(FHD 100%에서 900×800)를 가운데에 둔 영역."""
        hint = self.minimumSizeHint()
        area = screen.availableGeometry()
        width, height = default_window_size((area.width(), area.height()), (hint.width(), hint.height()))
        x, y = centered_position((area.x(), area.y(), area.width(), area.height()), (width, height))
        return QRect(x, y, width, height)

    def reset_window_geometry(self):
        """「설정 › 일반」의 「창 크기 기본값으로」: 최대화를 풀고 지금 창이 있는 모니터의 기본 크기로 가운데에 둔 뒤 저장한다."""
        screen = self.screen() or QGuiApplication.primaryScreen()
        if screen is None:
            return
        if self.windowState() & (Qt.WindowState.WindowMaximized | Qt.WindowState.WindowFullScreen
                                 | Qt.WindowState.WindowMinimized):
            self.showNormal()
        rect = self._default_window_rect(screen)
        self.setGeometry(rect)
        self._save_window_geometry()
        self.notify("ok", f"창 크기를 기본값({rect.width()}×{rect.height()})으로 되돌렸습니다")
        self._refresh_window_size_label()

    def _refresh_window_size_label(self):
        """「설정 › 일반」 버튼 옆에 지금 모니터의 기본 크기를 보인다(모니터·배율마다 다름)."""
        label = getattr(self.ui, "label_window_size", None)
        screen = self.screen() or QGuiApplication.primaryScreen()
        if label is None or screen is None:
            return
        rect = self._default_window_rect(screen)
        label.setText(f"이 모니터 기본 {rect.width()}×{rect.height()}")

    def _save_window_geometry(self):
        """창 위치·크기·최대화 여부를 이 PC의 앱 설정에만 저장한다(공유 시트에는 쓰지 않음)."""
        if not getattr(self, "_window_shown_once", False):
            return
        state = self.windowState()
        maximized = bool(state & Qt.WindowState.WindowMaximized)
        rect = self.normalGeometry() if (maximized or state & Qt.WindowState.WindowMinimized) else self.geometry()
        if rect.width() <= 0 or rect.height() <= 0:
            return
        self.set_app_setting(
            _host(self).WINDOW_GEOMETRY_SETTING_KEY,
            geometry_to_setting((rect.x(), rect.y(), rect.width(), rect.height()), maximized),
        )

    def showEvent(self, event):
        super().showEvent(event)
        self._window_shown_once = True
        if not self._order_ship_splitter_initialized:
            QTimer.singleShot(0, self._on_first_show_splitter)
        if not self._startup_order_index_sync_started:
            self._startup_order_index_sync_started = True
            QTimer.singleShot(0, self._kickoff_startup_order_index_sync)
            # 첫 화면을 그린 뒤 엑셀 처리용 pandas를 미리 불러와 첫 주문 엑셀 작업의 대기를 줄인다.
            QTimer.singleShot(
                _host(self).DATA_MODULE_PRELOAD_DELAY_MS,
                lambda: preload_modules_in_background("pandas"),
            )
            # 공유 시트가 요구하는 최소 버전 확인과 이 PC의 버전 기록(시작 동기화 뒤에).
            QTimer.singleShot(_host(self).APP_VERSION_CHECK_DELAY_MS, self._kickoff_app_version_check)

    def _kickoff_app_version_check(self):
        running = getattr(self, "_app_version_check_thread", None)
        if running is not None and isValid(running) and running.isRunning():
            return
        thread = _host(self).AppVersionCheckThread(self)
        thread.result_ready.connect(self._on_app_version_checked)
        thread.finished.connect(thread.deleteLater)
        self._app_version_check_thread = thread
        thread.start()
        if getattr(self, "_app_version_timer", None) is None:
            # 실행 중에도 1시간마다 다른 PC의 새 버전과 공유 최소 버전을 다시 확인한다.
            from sheets.version_gate import VERSION_RECHECK_INTERVAL_MS
            self._app_version_timer = QTimer(self)
            self._app_version_timer.setInterval(VERSION_RECHECK_INTERVAL_MS)
            self._app_version_timer.timeout.connect(self._kickoff_app_version_check)
            self._app_version_timer.start()

    def _on_app_version_checked(self, result):
        if not result.get("ok"):
            print(f"! 앱 버전 확인 실패: {result.get('error')}")
            label = getattr(self.ui, "label_app_version", None)
            if label is not None and not getattr(self, "_version_blocked_notice_shown", False):
                label.setText(f"현재 버전 {APP_VERSION} · 공유 시트 버전 확인 실패")
            return
        self._update_app_version_label(result)
        if result.get("blocked"):
            # 시트 쓰기가 멈춘 상태라 「새 버전 안내 무시」와 관계없이 실행마다 한 번 알린다.
            if not getattr(self, "_version_blocked_notice_shown", False):
                self._version_blocked_notice_shown = True
                self._show_required_update_notice(result.get("minimum", ""))
            return
        app_log.log(f"앱 버전 확인 완료 ({result.get('version')}, 공유 최소 {result.get('minimum')})")
        if result.get("latest"):
            self._maybe_show_update_notice(result["latest"])

    def _update_app_version_label(self, result):
        label = getattr(self.ui, "label_app_version", None)
        if label is None:
            return
        text = f"현재 버전 {result.get('version') or APP_VERSION}"
        if result.get("blocked"):
            text += f" · 공유 최소 버전 {result.get('minimum')} 필요 (시트 쓰기 중지)"
        elif result.get("latest"):
            text += f" · 새 버전 {result['latest']} 확인 (소스트리에서 Fetch → Pull 후 다시 실행)"
        else:
            text += " · 공유 시트에 새 버전 기록 없음"
        label.setText(text)
        chip = getattr(self, "statusChip_version", None)
        if chip is not None:
            version = result.get("version") or APP_VERSION
            if result.get("blocked"):
                set_status(chip, "error", f"v{version} · 업데이트 필요", text)
            elif result.get("latest"):
                set_status(chip, "warn", f"v{version} → {result['latest']}", text)
            else:
                set_status(chip, "idle", f"v{version}", text)
            chip.setToolTip(chip.toolTip() + " · 누르면 「설정 › 프로그램」")

    def _maybe_show_update_notice(self, latest):
        """다른 PC가 쓰는 새 버전을 알린다. 닫으면 이번 실행 동안, 「오늘 더 이상 묻지 않음」이면 오늘 묻지 않는다."""
        from sheets.version_gate import (
            SETTING_IGNORE_UPDATE_NOTICE, SETTING_UPDATE_NOTICE_SNOOZED_DAY, today_text, update_notice_due)
        dismissed = getattr(self, "_version_notice_dismissed", None)
        if dismissed is None:
            dismissed = self._version_notice_dismissed = set()
        if latest in dismissed:
            return
        today = today_text()
        if not update_notice_due(latest,
                                 ignored=bool(self.get_app_setting(SETTING_IGNORE_UPDATE_NOTICE, False)),
                                 snoozed_day=str(self.get_app_setting(SETTING_UPDATE_NOTICE_SNOOZED_DAY, "") or ""),
                                 today=today):
            return
        dismissed.add(latest)
        if self._ask_update_notice(latest) == "snooze":
            self.set_app_setting(SETTING_UPDATE_NOTICE_SNOOZED_DAY, today)

    def _version_update_dialog(self, title, icon, heading, status):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QGridLayout(dialog)
        layout.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)
        layout.setColumnStretch(1, 1)

        icon_label = QLabel()
        icon_label.setPixmap(dialog.style().standardIcon(icon).pixmap(32, 32))
        layout.addWidget(icon_label, 0, 0, Qt.AlignmentFlag.AlignTop)
        summary = QLabel(f'<p><font size="+1"><b>{heading}</b></font></p><p>{status}</p>')
        summary.setWordWrap(True)
        layout.addWidget(summary, 0, 1)

        steps = "<br>".join(
            f"{number}. <b>{html.escape(action)}</b> — {html.escape(detail)}"
            for number, (action, detail) in enumerate(SOURCETREE_UPDATE_STEPS, start=1)
        )
        guide = QLabel(f"<p>{steps}</p><p>{html.escape(SOURCETREE_UPDATE_ERROR)}</p>")
        guide.setWordWrap(True)
        guide.hide()
        layout.addWidget(guide, 2, 0, 1, 2)

        buttons = QDialogButtonBox()
        details = buttons.addButton("자세한 방법 보기", QDialogButtonBox.ButtonRole.ActionRole)
        details.setAutoDefault(False)
        details.setCheckable(True)
        details.toggled.connect(guide.setVisible)
        details.toggled.connect(
            lambda expanded: details.setText("자세한 방법 숨기기" if expanded else "자세한 방법 보기")
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 1, 0, 1, 2)
        return dialog, buttons

    def _show_required_update_notice(self, minimum):
        dialog, buttons = self._version_update_dialog(
            "프로그램 업데이트 필요", QStyle.StandardPixmap.SP_MessageBoxWarning,
            f"시트 쓰기 중지<br>현재 {APP_VERSION} → 최소 버전 {html.escape(minimum)}",
            "업데이트할 때까지 Google 시트는 읽기만 가능합니다.",
        )
        buttons.setStandardButtons(QDialogButtonBox.StandardButton.Ok)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setDefault(True)
        dialog.exec()

    def _ask_update_notice(self, latest):
        """새 버전 안내 창. 반환: 'snooze'(오늘 더 이상 묻지 않음) 또는 'close'."""
        dialog, buttons = self._version_update_dialog(
            "새 버전 안내", QStyle.StandardPixmap.SP_MessageBoxInformation,
            f"현재 {APP_VERSION} → 새 버전 {html.escape(latest)}",
            "업데이트 전에도 계속 사용할 수 있습니다.",
        )
        buttons.addButton("오늘 더 이상 묻지 않음", QDialogButtonBox.ButtonRole.AcceptRole)
        close = buttons.addButton("닫기", QDialogButtonBox.ButtonRole.RejectRole)
        close.setDefault(True)
        return "snooze" if dialog.exec() == QDialog.DialogCode.Accepted else "close"

    def _on_ignore_version_update_toggled(self, checked):
        from sheets.version_gate import SETTING_IGNORE_UPDATE_NOTICE
        self.set_app_setting(SETTING_IGNORE_UPDATE_NOTICE, bool(checked))

    # ── 종료: 대기 중인 공유 시트 반영을 마친 뒤 닫는다 ─────────────────────
    _BACKGROUND_TIMER_NAMES = (
        "_index_sheet_poll_timer", "_tracking_list_timer", "_tracking_auto_refresh_timer",
        "_naver_inquiry_timer", "_app_version_timer",
    )

    def _stop_background_timers(self):
        """종료 중 새 폴링·자동 새로고침이 시작되지 않게 주기 타이머를 멈춘다.

        반영 대기(디바운스) 타이머는 대기 여부 판단에 쓰므로 여기서 멈추지 않고,
        _kick_pending_sheet_writes가 즉시 반영을 시작할 때 멈춘다.
        """
        for name in self._BACKGROUND_TIMER_NAMES:
            timer = getattr(self, name, None)
            if timer is not None:
                timer.stop()

    def _pending_sheet_writes(self):
        """아직 공유 시트에 반영되지 않은 항목 이름 목록(종료 전에 마쳐야 하는 것)."""
        if _host(self).gspread is None:
            return []  # 시트 연동이 없으면 반영할 방법이 없으므로 기다리지 않는다.
        pending = []
        if (self._index_sheet_push_pending
                or self._index_sheet_push_timer.isActive()
                or isinstance(self._index_sheet_op_thread, _host(self).OrderIndexWriteThread)):
            pending.append("주문번호 인덱스")
        if self._tracking_pending_records or self._tracking_op_thread is not None:
            pending.append("송장추적 등록")
        if self._tracking_notes_pending or self._tracking_notes_update_thread is not None:
            pending.append("배송추적 메모")
        return pending

    def _kick_pending_sheet_writes(self):
        """대기 중인 반영을 디바운스 없이 바로 시작한다(실패한 항목은 정해진 횟수까지 재시도)."""
        attempts = self._close_flush_attempts
        if ((self._index_sheet_push_pending or self._index_sheet_push_timer.isActive())
                and self._index_sheet_op_thread is None
                and attempts["index"] < _host(self).CLOSE_FLUSH_MAX_ATTEMPTS):
            attempts["index"] += 1
            self._index_sheet_push_timer.stop()
            self._index_sheet_push_pending = True  # 결과가 성공으로 올 때까지 대기 상태로 본다.
            self._flush_order_indices_to_sheet()
        if (self._tracking_pending_records and self._tracking_op_thread is None
                and attempts["tracking"] < _host(self).CLOSE_FLUSH_MAX_ATTEMPTS):
            attempts["tracking"] += 1
            self._tracking_push_timer.stop()
            self._tracking_controller._flush_tracking_records()
        if (self._tracking_notes_pending and self._tracking_notes_update_thread is None
                and attempts["notes"] < _host(self).CLOSE_FLUSH_MAX_ATTEMPTS):
            attempts["notes"] += 1
            self._tracking_controller._flush_tracking_notes()

    def closeEvent(self, event):
        if getattr(self, "_close_ready", False):
            event.accept()
            return
        if not getattr(self, "_window_geometry_saved", False):
            self._window_geometry_saved = True  # 종료 대기 중 다시 불려도 처음 닫은 위치를 남긴다
            self._save_window_geometry()
        self._stop_background_timers()
        if not self._pending_sheet_writes():
            self._close_ready = True
            event.accept()
            return
        # 종료를 잠시 미루고 대기 중인 반영을 마친다. 끝나면 다시 close()를 호출한다.
        event.ignore()
        if getattr(self, "_close_flush_timer", None) is not None:
            return  # 이미 종료 준비 중
        self._close_flush_attempts = {"index": 0, "tracking": 0, "notes": 0}
        self._close_flush_deadline = time.monotonic() + _host(self).CLOSE_FLUSH_TIMEOUT_MS / 1000
        self._show_busy_processing_overlay(
            "공유 시트에 저장 중…", "주문번호 인덱스·송장추적 반영을 마친 뒤 종료합니다.")
        self.statusBar().showMessage("공유 시트에 저장 중… 완료되면 종료합니다.")
        self._close_flush_timer = QTimer(self)
        self._close_flush_timer.setInterval(100)
        self._close_flush_timer.timeout.connect(self._on_close_flush_tick)
        self._close_flush_timer.start()
        self._kick_pending_sheet_writes()

    def _on_close_flush_tick(self):
        pending = self._pending_sheet_writes()
        retry_left = any(count < _host(self).CLOSE_FLUSH_MAX_ATTEMPTS for count in self._close_flush_attempts.values())
        if pending and time.monotonic() < self._close_flush_deadline and retry_left:
            self._kick_pending_sheet_writes()
            return
        self._close_flush_timer.stop()
        if pending:
            answer = QMessageBox.question(
                self, "공유 시트 저장 확인",
                "다음 항목을 공유 시트에 아직 반영하지 못했습니다.\n\n- " + "\n- ".join(pending) +
                "\n\n지금 종료하면 이 변경이 사라질 수 있습니다(주문번호 인덱스는 다음 실행 때 "
                "시트 값으로 돌아갑니다).\n조금 더 기다리며 다시 시도할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if answer == QMessageBox.StandardButton.Yes:
                self._close_flush_attempts = {"index": 0, "tracking": 0, "notes": 0}
                self._close_flush_deadline = time.monotonic() + _host(self).CLOSE_FLUSH_TIMEOUT_MS / 1000
                self._close_flush_timer.start()
                self._kick_pending_sheet_writes()
                return
        self._close_flush_timer = None
        self._hide_busy_processing_overlay()
        self._close_ready = True
        self.close()
