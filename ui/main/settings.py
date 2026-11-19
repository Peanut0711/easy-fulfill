"""메인 창 「⚙ 설정」 탭: 분류 목록·페이지, 연동(Google 인증, API 키 행, 네이버·쿠팡 키 등록), 알림(문의 알림·근무 시간), 배송추적 설정 묶음, 주문·사은품, DB 동기화, 프로그램(우체국 확인·진단 버튼)."""

import re
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QUrl
from PySide6.QtGui import QBrush, QColor, QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QVBoxLayout,
)

from shared.ui_style import clear_status, COLOR_DANGER, set_band, set_button_role, set_status
from tracking.repository import ConfigValueChangedError
from tracking.service import (
    CONFIG_KEY_INQUIRY_WORK_END,
    CONFIG_KEY_INQUIRY_WORK_START,
    NAVER_INQUIRY_WORK_END_HOUR,
    NAVER_INQUIRY_WORK_START_HOUR,
)
from ui.gift_event_settings import GiftEventSettingsDialog
from ui.gift_history import GiftHistoryDialog
from shared import theme
from shared.ui_components import fixed_width_button
from ui.main import host as _host


SETTINGS_NAV_WIDTH = 140
SETTINGS_PAGE_MAX_WIDTH = 760


class SettingsMixin:
    """MainWindow 믹스인. 메인 창 「⚙ 설정」 탭: 분류 목록·페이지, 연동(Google 인증, API 키 행, 네이버·쿠팡 키 등록), 알림(문의 알림·근무 시간), 배송추적 설정 묶음, 주문·사은품, DB 동기화, 프로그램(우체국 확인·진단 버튼)."""

    def _inquiry_work_hours(self):
        """문의 알림 허용 시작/종료 시각(공유 설정 동기화된 로컬 값). 기본 (10, 19)."""
        def _h(key, default):
            try:
                v = int(self.get_app_setting(key, default))
                return v if 0 <= v <= 23 else default
            except (TypeError, ValueError):
                return default
        start = _h("inquiry_work_start_hour", NAVER_INQUIRY_WORK_START_HOUR)
        end = _h("inquiry_work_end_hour", NAVER_INQUIRY_WORK_END_HOUR)
        if start >= end:
            return NAVER_INQUIRY_WORK_START_HOUR, NAVER_INQUIRY_WORK_END_HOUR
        return start, end

    def _sync_inquiry_hours_spinboxes(self):
        """설정 팝업의 알림 시간대 스핀박스를 현재 값으로 갱신(시그널 차단)."""
        start, end = self._inquiry_work_hours()
        for box, val in ((getattr(self, "_dlg_inquiry_start", None), start),
                         (getattr(self, "_dlg_inquiry_end", None), end)):
            if box is not None:
                box.blockSignals(True)
                box.setValue(val)
                box.blockSignals(False)

    def _on_inquiry_work_hours_changed(self, _value=None):
        """알림 시간대(시작/종료) 변경 → 로컬 저장 + 공유 「설정」 탭 반영(전원 적용).
        종료는 시작보다 최소 1시간 뒤가 되도록 보정한다."""
        if self._dlg_inquiry_start is None or self._dlg_inquiry_end is None:
            return
        start = int(self._dlg_inquiry_start.value())
        end = int(self._dlg_inquiry_end.value())
        if end <= start:
            end = min(start + 1, 23)
            self._dlg_inquiry_end.blockSignals(True)
            self._dlg_inquiry_end.setValue(end)
            self._dlg_inquiry_end.blockSignals(False)
        self.set_app_setting("inquiry_work_start_hour", start)
        self.set_app_setting("inquiry_work_end_hour", end)
        self._tracking_controller.queue_tracking_config_write({
            CONFIG_KEY_INQUIRY_WORK_START: str(start),
            CONFIG_KEY_INQUIRY_WORK_END: str(end),
        })

    def _set_key_status_text(self, text):
        t = (text or "").strip()
        for p in ("인증키: ", "인증키:"):
            if t.startswith(p):
                t = t[len(p):].strip()
                break
        self._key_status_text = t
        self._update_status_displays()

    def _key_status_short(self):
        t = self._key_status_text
        if "유효함" in t:
            return "유효함 ✓"
        if "미설정" in t:
            return "미설정"
        if "확인 중" in t or "검증 중" in t:
            return "확인 중…"
        if "실패" in t:
            return "확인 실패"
        return t

    def _update_status_displays(self):
        """밖(작은 한 줄)·팝업 라벨·설정 탭 API키 상태를 함께 갱신."""
        if hasattr(self.ui, "label_tracking_status_compact"):
            self.ui.label_tracking_status_compact.setText(
                f"우체국 API: {self._key_status_short()}  ·  슬랙: {self._slack_status_text}"
            )
        if self._dlg_key_status is not None:
            self._dlg_key_status.setText(f"인증키: {self._key_status_text}")
        if self._dlg_slack_status is not None:
            self._dlg_slack_status.setText(f"슬랙 알림: {self._slack_status_text}")
        label_slack = getattr(self.ui, "label_slack_webhook_status", None)
        if label_slack is not None:
            if "설정됨" in self._slack_status_text:
                set_status(label_slack, "ok", "설정됨")
            else:
                set_status(label_slack, "warn", "미설정 — 위험 알림·문의 알림을 보내지 않습니다")
        self._update_api_key_rows()
        self._update_settings_nav_badges()

    # 「설정 › 연동 › API 키」 행: key → (상태 라벨, 버튼)
    API_KEY_ROWS = {
        "kpost": ("label_api_status_kpost", "pushButton_cfg_kpost_key"),
        "postcode": ("label_api_status_postcode", "pushButton_cfg_kpost_postcode_key"),
        "naver": ("label_api_status_naver", "pushButton_cfg_naver_key"),
        "coupang": ("label_api_status_coupang", "pushButton_cfg_coupang_key"),
    }

    def _api_key_row_status(self, key):
        """키 행 상태 ``(kind, text, detail)``. 종추적은 기존 키 확인 결과, 나머지는 공유 「설정」 탭에
        값이 있는지(시작할 때 이미 읽은 값)와 이 PC에서 방금 한 검증 결과로 정한다."""
        if key == "kpost":
            text = self._key_status_text.replace("✓", "").replace("  ", " ").strip()
            return self._key_status_kind(), text, ""
        return getattr(self, "_shared_key_status", {}).get(key, ("idle", "확인 중…", ""))

    def _update_api_key_rows(self, keys=None):
        if getattr(self, "ui", None) is None:  # 화면을 만들기 전(또는 화면 없는 테스트)
            return
        for key, (label_name, button_name) in self.API_KEY_ROWS.items():
            if keys is not None and key not in keys:
                continue
            kind, text, detail = self._api_key_row_status(key)
            label = getattr(self.ui, label_name, None)
            if label is not None:
                set_status(label, kind, text, detail)
            button = getattr(self.ui, button_name, None)
            if button is not None:
                button.setText("등록" if text.startswith("미설정") else "변경")

    def _shared_key_status_store(self):
        if "_shared_key_status" not in self.__dict__:
            self._shared_key_status = {}
        return self._shared_key_status

    def _set_shared_key_status(self, key, kind, text, detail=""):
        self._shared_key_status_store()[key] = (kind, text, detail)
        self._update_api_key_rows((key,))

    def _apply_shared_key_registration(self, registered):
        """공유 「설정」 탭을 읽은 결과(키별 값 유무)를 키 행에 반영한다. ``None``이면 읽기 실패."""
        self._shared_key_status_store()
        if registered is not None:
            self._shared_keys_known = True
        for key in ("postcode", "naver", "coupang"):
            if registered is None:
                if self._shared_key_status.get(key, ("idle",))[0] == "idle":
                    self._shared_key_status[key] = ("idle", "확인 못 함 (공유 시트를 읽지 못함)", "")
                continue
            if registered.get(key):
                self._shared_key_status[key] = ("ok", "등록됨", "")
            elif key == "postcode" and str(self.get_app_setting("kpost_regkey", "") or "").strip():
                self._shared_key_status[key] = (
                    "warn", "미설정 — 종추적 키로 조회", "주소 검색 전용 키가 없어 종추적 키로 우편번호를 조회합니다.")
            else:
                self._shared_key_status[key] = ("warn", "미설정", "")
        self._update_api_key_rows(("postcode", "naver", "coupang"))

    def _key_status_kind(self):
        """우체국 종추적 키 상태 문구를 상태 종류(ok·warn·error·idle)로 바꾼다."""
        t = self._key_status_text
        if "실패" in t:
            return "error"
        if "확인 중" in t or "검증 중" in t:
            return "idle"
        if "유효함" in t:
            return "ok"
        return "warn"  # 미설정, 오프라인(유효성 미확인)

    def _build_admin_tracking_settings(self):
        """슬랙·문의 알림은 「설정 › 알림」, 위험 판정 기준·자동 새로고침은 「설정 › 배송추적」에 구성.
        시작 시 1회 호출. 묶음 제목 뒤 「· 이 PC」(이 PC 앱 설정)·「· 전원 공유」(공유 시트)로 적용 범위를 구분한다."""
        host = getattr(self.ui, "widget_settings_alerts_host", None)
        tracking_host = getattr(self.ui, "widget_settings_tracking_host", None)
        if (host is None or host.layout() is None or tracking_host is None
                or tracking_host.layout() is None or self._admin_tracking_built):
            return
        outer = host.layout()
        self._admin_tracking_built = True

        self._tracking_controller.build_admin_slack_controls(outer)

        # 미답변 문의 알림: 자동 조회 토글은 이 PC, 알림 시간대는 전원 공유라 묶음을 나눈다.
        gb_naver = QGroupBox("미답변 문의 알림 · 이 PC")
        ndl = QVBoxLayout(gb_naver)
        self._dlg_naver_status = QLabel()
        self._dlg_naver_status.setObjectName("label_inquiry_alert_status")
        self._dlg_naver_status.setWordWrap(True)  # 긴 실패 문구가 페이지 폭을 넓히지 않게
        ndl.addWidget(self._dlg_naver_status)
        nrow = QHBoxLayout()
        self._dlg_naver_enable = QCheckBox("자동 조회(5분마다 네이버·쿠팡 → 슬랙)")
        self._dlg_naver_enable.setChecked(bool(self.get_app_setting("naver_inquiry_notify", False)))
        _ws, _we = self._inquiry_work_hours()
        self._dlg_naver_enable.setToolTip(
            "켜면 이 PC가 5분마다 네이버·쿠팡 상품·고객문의(미답변)를 조회합니다.\n"
            f"새 미답변은 즉시, 처리 안 된 건은 60분마다 다시 슬랙으로 알립니다"
            f"(평일 {_ws}~{_we}시, 아래에서 변경).\n"
            "누군가 답변하면 자동으로 멈춥니다. 전원 공유 시트(「문의알림」)로 중복을 막으니\n"
            "여러 대를 동시에 켜둬도 됩니다(호출을 줄이려면 일부만 켜도 무방)."
        )
        self._dlg_naver_enable.stateChanged.connect(self._on_naver_inquiry_toggled)
        btn_npoll = QPushButton("지금 확인")
        btn_npoll.setToolTip(
            "토글과 무관하게 지금 한 번 조회합니다. "
            "네이버·쿠팡 키 입력은 「설정 › 연동」의 API 키에서 합니다.")
        btn_npoll.clicked.connect(self._on_naver_inquiry_manual_poll)
        nrow.addWidget(self._dlg_naver_enable)
        nrow.addStretch(1)
        nrow.addWidget(btn_npoll)
        ndl.addLayout(nrow)
        outer.addWidget(gb_naver)
        # 알림 허용 시간대(평일, 시작~종료 시각). 변경 시 전원 공유.
        gb_hours = QGroupBox("문의 알림 시간대 · 전원 공유")
        hrow = QHBoxLayout(gb_hours)
        hrow.addWidget(QLabel("평일"))
        _ih_start, _ih_end = self._inquiry_work_hours()
        self._dlg_inquiry_start = QSpinBox()
        self._dlg_inquiry_start.setRange(0, 23)
        self._dlg_inquiry_start.setSuffix("시")
        self._dlg_inquiry_start.setValue(_ih_start)
        self._dlg_inquiry_end = QSpinBox()
        self._dlg_inquiry_end.setRange(0, 23)
        self._dlg_inquiry_end.setSuffix("시")
        self._dlg_inquiry_end.setValue(_ih_end)
        _ih_tip = ("이 시간대(평일)에만 미답변 문의 알림을 보냅니다. 그 외 시간/주말은 "
                   "조용히 누적했다가 다음 근무 시작 시각에 한 번에 환기합니다. "
                   "변경 시 공유 시트에 저장되어 전원에게 적용됩니다. 기본 10~19시.")
        self._dlg_inquiry_start.setToolTip(_ih_tip)
        self._dlg_inquiry_end.setToolTip(_ih_tip)
        self._dlg_inquiry_start.valueChanged.connect(self._on_inquiry_work_hours_changed)
        self._dlg_inquiry_end.valueChanged.connect(self._on_inquiry_work_hours_changed)
        hrow.addWidget(self._dlg_inquiry_start)
        hrow.addWidget(QLabel("~"))
        hrow.addWidget(self._dlg_inquiry_end)
        hrow.addStretch(1)
        outer.addWidget(gb_hours)

        self._tracking_controller.build_admin_risk_controls(tracking_host.layout())

        self._update_naver_inquiry_status_label()
        self._update_status_displays()

    # ── 네이버 문의 알림(상품문의·고객문의 → 슬랙) ──────────────────────────
    def _init_naver_inquiry_polling(self):
        """이 PC에서 문의 자동 알림이 켜져 있으면 5분 폴링을 시작합니다.
        시작 직후 1회 조회(미답변을 누적·환기; 알림은 평일 10~19시에만 발송)."""
        if _host(self).gspread is None:
            return
        if bool(self.get_app_setting("naver_inquiry_notify", False)):
            self._naver_inquiry_timer.start()
            # 시작 직후 다른 로딩과 겹치지 않게 약간 지연 후 첫 조회.
            QTimer.singleShot(8000, self._on_naver_inquiry_poll)

    def _on_naver_inquiry_toggled(self, state):
        enabled = bool(state)
        self.set_app_setting("naver_inquiry_notify", enabled)
        if enabled:
            self._naver_inquiry_timer.start()
            self._set_naver_inquiry_status("조회 중…", "idle")
            self._on_naver_inquiry_poll()
        else:
            self._naver_inquiry_timer.stop()
            self._set_naver_inquiry_status("자동 조회 꺼짐", "idle")

    def _on_naver_inquiry_manual_poll(self):
        self._set_naver_inquiry_status("조회 중…", "idle")
        self._on_naver_inquiry_poll()

    def _on_naver_inquiry_poll(self):
        if _host(self).gspread is None or self._naver_inquiry_thread is not None:
            return
        thread = _host(self).NaverInquiryPollThread(self)
        self._naver_inquiry_thread = thread
        thread.result_ready.connect(self._on_naver_inquiry_result)
        thread.finished.connect(self._cleanup_naver_inquiry_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_naver_inquiry_result(self, payload: dict):
        t = datetime.now().strftime("%H:%M")
        if not payload.get("ok"):
            err = str(payload.get("error", ""))
            print(f"! 문의 조회 실패: {err}")
            self._set_naver_inquiry_status(f"실패: {err[:60]} · {t}", "error", err)
            return
        if payload.get("skipped_other_pc"):
            self._set_naver_inquiry_status(f"다른 PC가 방금 확인함 · {t}", "ok")
            return
        open_cnt = payload.get("open", 0)
        new = payload.get("new", 0)
        reminded = payload.get("reminded", 0)
        sent = payload.get("sent", 0)
        off_hours = payload.get("off_hours", False)
        kind = "warn" if open_cnt else "ok"
        if off_hours:
            _ws, _we = self._inquiry_work_hours()
            self._set_naver_inquiry_status(
                f"미답변 {open_cnt}건 · 알림 시간대 아님(평일 {_ws}~{_we}시) · {t}", kind)
        elif new or reminded:
            self._set_naver_inquiry_status(
                f"미답변 {open_cnt}건 · 알림 {sent}건(신규 {new}/리마인더 {reminded}) · {t}", kind)
        else:
            self._set_naver_inquiry_status(f"미답변 {open_cnt}건 · 새 알림 없음 · {t}", kind)
        for e in payload.get("errors", []) or []:
            print(f"! 문의 알림: {e}")

    def _cleanup_naver_inquiry_thread(self):
        self._naver_inquiry_thread = None

    def _set_naver_inquiry_status(self, text, kind="idle", detail=""):
        """문의 알림 상태. kind: ok(정상)·warn(미답변 있음)·error(실패)·idle(조회 중·꺼짐)."""
        self._naver_inquiry_status_text = text
        self._naver_inquiry_status_kind = kind
        self._naver_inquiry_status_detail = detail
        self._update_naver_inquiry_status_label()

    def _update_naver_inquiry_status_label(self):
        if self._dlg_naver_status is not None:
            set_status(self._dlg_naver_status, self._naver_inquiry_status_kind,
                       self._naver_inquiry_status_text,
                       getattr(self, "_naver_inquiry_status_detail", ""))

    def _on_naver_creds_edit_clicked(self):
        """관리자용: client_id/secret 을 입력받아 토큰 발급으로 검증 후 공유 시트에 저장."""
        if _host(self).gspread is None:
            QMessageBox.warning(self, "키 설정", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        if self._naver_creds_validate_thread is not None:
            return
        cid, ok = QInputDialog.getText(
            self, "네이버 커머스API · client_id",
            "client_id 를 입력하세요.\n(검증 후 직원 전원 공유 시트에 저장됩니다)",
            QLineEdit.EchoMode.Normal, "",
        )
        if not ok:
            return
        cid = cid.strip()
        if not cid:
            QMessageBox.warning(self, "키 설정", "client_id 가 비어 있습니다.")
            return
        csec, ok = QInputDialog.getText(
            self, "네이버 커머스API · client_secret",
            "client_secret 을 입력하세요.",
            QLineEdit.EchoMode.Password, "",
        )
        if not ok:
            return
        csec = csec.strip()
        if not csec:
            QMessageBox.warning(self, "키 설정", "client_secret 이 비어 있습니다.")
            return
        self._naver_pending_creds = (cid, csec)
        self._set_shared_key_status("naver", "idle", "검증 중…")
        thread = _host(self).NaverCredsValidateThread(cid, csec, self)
        self._naver_creds_validate_thread = thread
        thread.result_ready.connect(self._on_naver_creds_validated)
        thread.finished.connect(self._cleanup_naver_creds_validate_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _cleanup_naver_creds_validate_thread(self):
        self._naver_creds_validate_thread = None

    def _on_naver_creds_validated(self, payload: dict):
        cid, csec = self._naver_pending_creds
        if not (payload.get("ok") and payload.get("valid")):
            err = payload.get("error", "토큰 발급 실패")
            QMessageBox.warning(
                self, "키 검증 실패",
                f"이 client_id/secret 으로 토큰 발급에 실패했습니다.\n\n{err}\n\n"
                "키를 다시 확인해 주세요. (저장하지 않았습니다)",
            )
            self._set_shared_key_status("naver", "error", "입력한 키 검증 실패 — 저장하지 않음", str(err))
            return
        # 검증 통과 → 공유 「설정」 탭에 저장(전원 적용)
        self._tracking_controller.queue_tracking_config_write(
            {_host(self).CONFIG_KEY_NAVER_CLIENT_ID: cid, _host(self).CONFIG_KEY_NAVER_CLIENT_SECRET: csec})
        self._set_shared_key_status("naver", "ok", "등록됨 (확인함 · 전원 적용)")

    def _on_coupang_creds_edit_clicked(self):
        """관리자용: 쿠팡 vendorId/accessKey/secretKey 를 입력받아 1회 조회로 검증 후 저장."""
        if _host(self).gspread is None:
            QMessageBox.warning(self, "키 설정", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        if self._coupang_creds_validate_thread is not None:
            return
        vid, ok = QInputDialog.getText(
            self, "쿠팡 WING OpenAPI · vendorId",
            "업체코드(vendorId, 예: A00012345)를 입력하세요.\n"
            "(검증 후 직원 전원 공유 시트에 저장됩니다)",
            QLineEdit.EchoMode.Normal, "",
        )
        if not ok:
            return
        vid = vid.strip()
        if not vid:
            QMessageBox.warning(self, "키 설정", "vendorId 가 비어 있습니다.")
            return
        ak, ok = QInputDialog.getText(
            self, "쿠팡 WING OpenAPI · accessKey",
            "accessKey 를 입력하세요.",
            QLineEdit.EchoMode.Normal, "",
        )
        if not ok:
            return
        ak = ak.strip()
        if not ak:
            QMessageBox.warning(self, "키 설정", "accessKey 가 비어 있습니다.")
            return
        sk, ok = QInputDialog.getText(
            self, "쿠팡 WING OpenAPI · secretKey",
            "secretKey 를 입력하세요.",
            QLineEdit.EchoMode.Password, "",
        )
        if not ok:
            return
        sk = sk.strip()
        if not sk:
            QMessageBox.warning(self, "키 설정", "secretKey 가 비어 있습니다.")
            return
        self._coupang_pending_creds = (vid, ak, sk)
        self._set_shared_key_status("coupang", "idle", "검증 중…")
        thread = _host(self).CoupangCredsValidateThread(vid, ak, sk, self)
        self._coupang_creds_validate_thread = thread
        thread.result_ready.connect(self._on_coupang_creds_validated)
        thread.finished.connect(self._cleanup_coupang_creds_validate_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _cleanup_coupang_creds_validate_thread(self):
        self._coupang_creds_validate_thread = None

    def _on_coupang_creds_validated(self, payload: dict):
        vid, ak, sk = self._coupang_pending_creds
        if not (payload.get("ok") and payload.get("valid")):
            err = payload.get("error", "조회 실패")
            QMessageBox.warning(
                self, "키 검증 실패",
                f"이 쿠팡 키로 문의 조회에 실패했습니다.\n\n{err}\n\n"
                "vendorId/accessKey/secretKey 를 다시 확인해 주세요. (저장하지 않았습니다)",
            )
            self._set_shared_key_status("coupang", "error", "입력한 키 검증 실패 — 저장하지 않음", str(err))
            return
        # 검증 통과 → 공유 「설정」 탭에 저장(전원 적용)
        self._tracking_controller.queue_tracking_config_write({
            _host(self).CONFIG_KEY_COUPANG_VENDOR_ID: vid,
            _host(self).CONFIG_KEY_COUPANG_ACCESS_KEY: ak,
            _host(self).CONFIG_KEY_COUPANG_SECRET_KEY: sk,
        })
        self._set_shared_key_status("coupang", "ok", "등록됨 (확인함 · 전원 적용)")

    def _open_gift_event_settings(self):
        def open_worksheet():
            from sheets.google_sheets_oauth import get_authorized_gspread_client
            return _host(self)._standalone_open_config_ws(get_authorized_gspread_client())

        dialog = GiftEventSettingsDialog(open_worksheet, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.statusBar().showMessage("사은품 행사 공유 설정을 저장했습니다.", 5000)
        dialog.deleteLater()

    def _open_gift_history(self):
        self._gift_history_sync.flush()
        dialog = GiftHistoryDialog(self._gift_history_sync.read_report, self)
        dialog.exec()
        dialog.deleteLater()

    def _invalidate_google_sheets_client_and_caches(self):
        self._gspread_client = None
        self._spreadsheet_product_code_maps = {}
        self._coupang_option_to_vp_product_no = {}

    def _refresh_google_auth_status_ui(self):
        """「설정 › 연동」의 Google 상태(토큰 파일 기준 + 실행 중 난 인증 오류)와 분류 목록 배지를 갱신한다."""
        try:
            from sheets.google_sheets_oauth import get_oauth_status

            kind, text, detail = get_oauth_status()
        except Exception as e:
            kind, text, detail = "error", "상태를 표시할 수 없습니다.", str(e)
        runtime_error = getattr(self, "_google_auth_runtime_error", "")
        if runtime_error and kind == "ok":
            kind, text, detail = "error", "인증 오류: 「재인증」이 필요합니다.", runtime_error
        self._google_auth_kind = kind
        label = getattr(self.ui, "label_google_auth_status", None)
        if label is not None:
            set_status(label, kind, text, detail)
        chip = getattr(self, "statusChip_google", None)
        if chip is not None:
            set_status(chip, kind, "Google", f"Google Sheets: {text} · 누르면 「설정 › 연동」")
        tip = getattr(self.ui, "label_google_auth_tip", None)
        if tip is not None:
            # 「연결 해제 후 재인증」 안내는 인증 오류일 때만 보인다.
            tip.setVisible(kind == "error")
        self._update_settings_nav_badges()

    def _mark_google_auth_runtime_error(self, err_text):
        """시트 작업 중 Google 인증 오류가 나면 기록해 「설정 › 연동」과 분류 목록에 ✗로 보인다."""
        self._google_auth_runtime_error = str(err_text or "").strip() or "Google 인증 오류"
        self._refresh_google_auth_status_ui()

    def _clear_google_auth_runtime_error(self):
        self._google_auth_runtime_error = ""

    def _settings_integration_problem(self):
        """연동에 문제가 있으면(Google 인증 실패, 종추적 키 확인 실패) True."""
        if getattr(self, "_google_auth_kind", "") == "error":
            return True
        return self._key_status_kind() == "error"

    def _update_settings_nav_badges(self):
        """「설정」 분류 목록에서 문제가 있는 분류 이름 뒤에 ✗를 붙인다(다른 탭에서 와도 보이게)."""
        nav = getattr(self.ui, "listWidget_settings_nav", None)
        stack = getattr(self.ui, "stackedWidget_settings", None)
        if nav is None or stack is None:
            return
        problems = {"scrollArea_settings_integration": self._settings_integration_problem()}
        for row in range(min(nav.count(), stack.count())):
            item = nav.item(row)
            base = item.data(Qt.ItemDataRole.UserRole)
            if base is None:
                base = item.text()
                item.setData(Qt.ItemDataRole.UserRole, base)
            bad = problems.get(stack.widget(row).objectName(), False)
            text = f"{base}  ✗" if bad else base
            if item.text() != text:
                item.setText(text)
                item.setForeground(QColor(COLOR_DANGER) if bad else QBrush())
                item.setToolTip("연동에 문제가 있습니다. 눌러서 확인하세요." if bad else "")

    def _setup_settings_tab(self):
        """「설정」 탭: 왼쪽 분류 목록과 오른쪽 페이지(QStackedWidget)를 잇는다."""
        nav = getattr(self.ui, "listWidget_settings_nav", None)
        stack = getattr(self.ui, "stackedWidget_settings", None)
        if nav is None or stack is None:
            return
        # 페이지(스크롤 영역)는 탭 배경색을 그대로 쓴다(내용 폭 밖이 다른 색으로 보이지 않게).
        for index in range(stack.count()):
            area = stack.widget(index)
            if isinstance(area, QScrollArea):
                area.viewport().setAutoFillBackground(False)
                if area.widget() is not None:
                    area.widget().setAutoFillBackground(False)
        # 디자인 개편 4-6: 분류 목록은 사이드바 모양(테마 #listWidget_settings_nav 규칙), 페이지 최대 폭 620→760.
        tab_layout = getattr(self.ui, "horizontalLayout_settings_tab", None)
        if tab_layout is not None:
            tab_layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, 0, 0)
            tab_layout.setSpacing(theme.SPACE_SM)
        nav.setFixedWidth(SETTINGS_NAV_WIDTH)
        slack_status = getattr(self.ui, "label_slack_webhook_status", None)
        if slack_status is not None:  # 긴 미설정 안내가 좁은 창에서 버튼을 밀어내지 않게 줄바꿈
            slack_status.setWordWrap(True)
        slack_button = getattr(self.ui, "pushButton_cfg_slack_webhook", None)
        if slack_button is not None:
            fixed_width_button(slack_button)
        separator = getattr(self.ui, "line_settings_nav", None)
        if separator is not None:  # 분류 목록 오른쪽 경계선은 테마가 그린다(선이 두 줄로 보이지 않게)
            separator.hide()
        for index in range(stack.count()):
            area = stack.widget(index)
            page = area.widget() if isinstance(area, QScrollArea) else None
            if page is not None and page.maximumWidth() < SETTINGS_PAGE_MAX_WIDTH:
                page.setMaximumWidth(SETTINGS_PAGE_MAX_WIDTH)
        nav.currentRowChanged.connect(self._on_settings_nav_changed)
        nav.setCurrentRow(0)
        stack.setCurrentIndex(0)
        self._setup_epost_diagnostic_buttons()

    def _epost_diagnostic_actions(self):
        """「설정 › 프로그램」에 버튼으로 두는 우체국 확인·진단 동작(실제 접수·인쇄·발송 없음). 작업 순서대로."""
        names = (
            "act_epost_test_receipt",              # testYn=Y 테스트 접수
            "act_epost_portal_diagnostic",         # 출력 화면 컨트롤 읽기 전용 점검
            "act_epost_portal_lookup",             # 신규출력 목록 대조(읽기 전용)
            "act_epost_portal_print_popup",        # 운송장출력 팝업, 인쇄 전에서 멈춤
            "act_epost_portal_oz_viewer",          # OZ Viewer 열림만 확인
            "act_epost_portal_output_confirm",     # 출력여부 읽기 전용 확인
            "act_epost_portal_reprint_popup",      # 단건 재출력 팝업, 인쇄 전에서 멈춤
            "act_epost_portal_reprint_oz_viewer",  # 단건 재출력 OZ Viewer 열림만 확인
            "act_epost_oz_print_dialog",           # Windows 인쇄 창 열림만 확인
        )
        return [(name, getattr(self, name)) for name in names if getattr(self, name, None) is not None]

    DEV_FEATURES_SETTING_KEY = "dev_features_enabled"

    def _dev_entry_actions(self):
        """개발 중 기능의 진입점(실제 반영 없음): API로 주문 불러오기, 포털 로그인 연결, 출력 대상 확인."""
        names = ("act_api_load", "act_epost_portal_login", "act_epost_print_targets")
        return [(name, getattr(self, name)) for name in names if getattr(self, name, None) is not None]

    def _dev_feature_actions(self):
        """「개발 중 기능」 스위치가 켜고 끄는 동작 전부(진입점, 확인·진단, 실제 동작)."""
        real = [(name, getattr(self, name)) for name in (
            "act_dispatch", "act_epost_real_receipt", "act_epost_actual_print", "act_epost_batch_actual_print")
            if getattr(self, name, None) is not None]
        return self._dev_entry_actions() + self._epost_diagnostic_actions() + real

    def _setup_epost_diagnostic_buttons(self):
        """「설정 › 프로그램」의 「개발 중 기능」 묶음(검토 리스트 R7).

        API·우체국 포털 기능은 개발 중이라 메뉴바에 두지 않고 여기에 모은다. 버튼은 같은 QAction을 부르고,
        「개발 중 기능 켜기(이 PC)」를 끄면 QAction 자체가 꺼져 어떤 경로로도 실행되지 않는다.
        """
        grid = getattr(self.ui, "gridLayout_epost_diagnostics", None)
        box_layout = getattr(self.ui, "verticalLayout_epost_diagnostics", None)
        if grid is None or box_layout is None or grid.count():
            return
        host = self.ui.groupBox_epost_diagnostics
        host.setTitle("개발 중 기능 (API·우체국) · 이 PC")
        intro = getattr(self.ui, "label_epost_diagnostics_intro", None)
        if intro is not None:
            intro.setText("실무에서는 쓰지 않는 개발 중 기능입니다. 개발·시험하는 PC에서만 켜세요. "
                          "실제 동작(접수·인쇄·발송처리)은 실행 전 확인 창이 뜹니다.")
            intro.setProperty("hint", "true")

        def caption(text):
            label = QLabel(text, host)
            label.setProperty("formLabel", True)
            return label

        def button_for(name, action, prefix):
            button = QPushButton(action.text().removeprefix("우체국 "), host)
            button.setObjectName(prefix + name.removeprefix("act_"))
            button.setToolTip(action.statusTip() or action.text())
            if not action.icon().isNull():
                button.setIcon(action.icon())
            button.clicked.connect(action.trigger)
            action.changed.connect(lambda b=button, a=action: b.setEnabled(a.isEnabled()))
            return button

        self.checkBox_dev_features = QCheckBox("개발 중 기능 켜기 (이 PC)", host)
        self.checkBox_dev_features.setObjectName("checkBox_dev_features")
        self.checkBox_dev_features.setToolTip("이 PC에만 저장됩니다. 끄면 아래 버튼과 동작이 모두 비활성입니다.")
        index = box_layout.indexOf(grid) if box_layout.indexOf(grid) >= 0 else box_layout.count()
        box_layout.insertWidget(index, self.checkBox_dev_features)

        entry_grid = QGridLayout()
        entry_grid.setHorizontalSpacing(theme.SPACE_SM)
        entry_grid.setVerticalSpacing(theme.SPACE_SM)
        for i, (name, action) in enumerate(self._dev_entry_actions()):
            entry_grid.addWidget(button_for(name, action, "pushButton_dev_"), i // 2, i % 2)
        box_layout.insertWidget(index + 1, caption("진입점"))
        box_layout.insertLayout(index + 2, entry_grid)
        box_layout.insertWidget(index + 3, caption("확인·진단 (실제 동작 없음)"))
        for i, (name, action) in enumerate(self._epost_diagnostic_actions()):
            grid.addWidget(button_for(name, action, "pushButton_diag_"), i // 2, i % 2)
        real_grid = QGridLayout()
        real_grid.setHorizontalSpacing(theme.SPACE_SM)
        real_grid.setVerticalSpacing(theme.SPACE_SM)
        for i, action in enumerate(self._real_api_actions()):
            name = next(n for n, a in self._dev_feature_actions() if a is action)
            button = button_for(name, action, "pushButton_real_")
            set_button_role(button, "danger")
            real_grid.addWidget(button, i // 2, i % 2)
        box_layout.addWidget(caption("실제 동작 — 실행 전 확인 창이 뜹니다"))
        box_layout.addLayout(real_grid)

        enabled = bool(self.get_app_setting(self.DEV_FEATURES_SETTING_KEY, False))
        self.checkBox_dev_features.setChecked(enabled)
        self._apply_dev_features_enabled(enabled)
        self.checkBox_dev_features.toggled.connect(self._on_dev_features_toggled)

    def _apply_dev_features_enabled(self, enabled):
        for _name, action in self._dev_feature_actions():
            action.setEnabled(bool(enabled))

    def _on_dev_features_toggled(self, checked):
        self.set_app_setting(self.DEV_FEATURES_SETTING_KEY, bool(checked))
        self._apply_dev_features_enabled(checked)

    def _on_settings_nav_changed(self, row):
        stack = getattr(self.ui, "stackedWidget_settings", None)
        if stack is None or row < 0 or row >= stack.count():
            return
        stack.setCurrentIndex(row)
        self._on_settings_page_shown(row)

    def show_settings_page(self, key):
        """다른 곳에서 「설정」의 특정 분류로 이동한다. key: integration·alerts·tracking·orders·db_sync·program."""
        stack = getattr(self.ui, "stackedWidget_settings", None)
        nav = getattr(self.ui, "listWidget_settings_nav", None)
        tab = getattr(self.ui, "tab_settings", None)
        if stack is None or nav is None or tab is None:
            return
        for index in range(stack.count()):
            if stack.widget(index).objectName() == f"scrollArea_settings_{key}":
                self.ui.tabWidget.setCurrentWidget(tab)
                nav.setCurrentRow(index)
                return

    def _on_settings_page_shown(self, index):
        """분류 페이지를 열 때 그 페이지의 상태를 다시 읽는다(페이지 객체 이름으로 분기)."""
        stack = getattr(self.ui, "stackedWidget_settings", None)
        page = stack.widget(index) if stack is not None else None
        name = page.objectName() if page is not None else ""
        if name == "scrollArea_settings_integration":
            self._refresh_google_auth_status_ui()
            settings_visible = (getattr(self.ui, "tabWidget", None) is not None
                                and self.ui.tabWidget.currentWidget() is getattr(self.ui, "tab_settings", None))
            if settings_visible and not getattr(self, "_shared_keys_known", False) and _host(self).gspread is not None:
                # 키 행 상태는 공유 「설정」 탭 읽기(배송추적 탭 진입 때 하던 것) 결과로 정한다.
                # 아직 한 번도 읽지 못했으면 여기서 같은 읽기를 시작한다(진행 중이면 다시 시작하지 않음).
                self._begin_tracking_config_read()
        elif name == "scrollArea_settings_db_sync":
            self._refresh_db_sheet_sync_path_labels()

    def _on_db_sync_refresh_paths_clicked(self):
        self._refresh_db_sheet_sync_path_labels(reset_manual_overrides=True)

    def _db_sync_start_dir(self, config_db_dir: str) -> str:
        p = Path(config_db_dir).resolve()
        return str(p) if p.is_dir() else str(Path.cwd())

    def _on_db_sync_naver_browse_clicked(self):
        if not hasattr(self.ui, "label_db_sync_naver_file"):
            return
        try:
            from sheets.db_sheet_sync import NAVER_CONFIG
        except ImportError:
            QMessageBox.warning(self, "DB동기화", "db_sheet_sync 모듈을 불러올 수 없습니다.")
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "네이버 상품 DB (CSV) 선택",
            self._db_sync_start_dir(NAVER_CONFIG["db_dir"]),
            "CSV (*.csv);;모든 파일 (*.*)",
        )
        if not path:
            return
        self._db_sync_naver_path_override = Path(path)
        self._refresh_db_sheet_sync_path_labels()

    def _on_db_sync_coupang_browse_clicked(self):
        if not hasattr(self.ui, "label_db_sync_coupang_file"):
            return
        try:
            from sheets.db_sheet_sync import COUPANG_CONFIG
        except ImportError:
            QMessageBox.warning(self, "DB동기화", "db_sheet_sync 모듈을 불러올 수 없습니다.")
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "쿠팡 가격·재고 DB (Excel) 선택",
            self._db_sync_start_dir(COUPANG_CONFIG["db_dir"]),
            "Excel (*.xlsx);;모든 파일 (*.*)",
        )
        if not path:
            return
        self._db_sync_coupang_path_override = Path(path)
        self._refresh_db_sheet_sync_path_labels()

    def _update_db_sync_test_mode_ui(self, *_args):
        """테스트 모드면 주의 색 띠를 보이고 실행 버튼을 「테스트 실행」으로, 끄면 띠를 숨기고 「실행」으로."""
        box = getattr(self.ui, "checkBox_db_sync_preview", None)
        if box is None:
            return
        test_mode = box.isChecked()
        band = getattr(self.ui, "label_db_sync_test_mode_band", None)
        if band is not None:
            set_band(band, "warn" if test_mode else None)
        button = getattr(self.ui, "pushButton_db_sync_run", None)
        if button is not None:
            button.setText("테스트 실행" if test_mode else "실행")
            button.setToolTip(
                "공유 시트에 쓰지 않고 추가될 행만 로그로 보입니다." if test_mode else
                "공유 시트의 네이버·쿠팡 탭에 없는 상품을 실제로 추가합니다(실행 전 확인 창).")

    def _set_db_sync_path_controls_busy(self, busy: bool):
        w = self.ui
        for name in (
            "pushButton_db_sync_refresh_paths",
            "pushButton_db_sync_run",
            "pushButton_db_sync_naver_browse",
            "pushButton_db_sync_coupang_browse",
        ):
            b = getattr(w, name, None)
            if b is not None:
                b.setEnabled(not busy)

    def _refresh_db_sheet_sync_path_labels(self, reset_manual_overrides=False):
        if not hasattr(self.ui, "label_db_sync_naver_file"):
            return
        try:
            from sheets.db_sheet_sync import (
                COUPANG_CONFIG,
                NAVER_CONFIG,
                format_db_sync_label_line,
                get_latest_file_from_pattern,
                get_latest_file_from_patterns,
            )
        except ImportError as e:
            set_status(self.ui.label_db_sync_naver_file, "error", "네이버 DB · 모듈 오류", str(e))
            set_status(self.ui.label_db_sync_coupang_file, "error", "쿠팡 DB · 모듈 오류", str(e))
            return
        if reset_manual_overrides:
            self._db_sync_naver_path_override = None
            self._db_sync_coupang_path_override = None
        n = self._db_sync_naver_path_override
        if n is None:
            n = get_latest_file_from_patterns(NAVER_CONFIG["db_dir"], NAVER_CONFIG["file_patterns"])
        c = self._db_sync_coupang_path_override
        if c is None:
            c = get_latest_file_from_pattern(COUPANG_CONFIG["db_dir"], COUPANG_CONFIG["file_pattern"])
        for label, title, path, manual in (
            (self.ui.label_db_sync_naver_file, "네이버 DB", n, self._db_sync_naver_path_override is not None),
            (self.ui.label_db_sync_coupang_file, "쿠팡 DB", c, self._db_sync_coupang_path_override is not None),
        ):
            text = format_db_sync_label_line(title, path) + (" · 직접 지정" if manual else "")
            if path is None:
                set_status(label, "warn", text, mark=False)  # 파일 없음은 주의 색
            else:
                clear_status(label)
                label.setText(text)
                label.setToolTip(str(path))  # 전체 경로

    def _on_db_sheet_sync_run_clicked(self):
        if not hasattr(self.ui, "plainTextEdit_db_sync_log"):
            return
        if self._db_sheet_sync_thread is not None and self._db_sheet_sync_thread.isRunning():
            self.statusBar().showMessage("DB동기화: 이미 실행 중입니다.", 5000)
            return
        if _host(self).gspread is None:
            QMessageBox.warning(self, "DB동기화", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        do_naver = self.ui.checkBox_db_sync_naver.isChecked()
        do_coupang = self.ui.checkBox_db_sync_coupang.isChecked()
        if not do_naver and not do_coupang:
            QMessageBox.warning(self, "DB동기화", "네이버 또는 쿠팡 중 하나 이상 선택하세요.")
            return
        test_mode = self.ui.checkBox_db_sync_preview.isChecked()
        test_count = 0
        verbose_log = (
            hasattr(self.ui, "checkBox_db_sync_verbose_log")
            and self.ui.checkBox_db_sync_verbose_log.isChecked()
        )
        if not test_mode:
            reply = QMessageBox.question(
                self,
                "DB동기화",
                "테스트 모드가 꺼져 있습니다. 스프레드시트에 신규 행을 실제로 추가합니다. 계속할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        self._show_busy_processing_overlay(
            "DB 동기화 작업 중…",
            "시트와 로컬 DB 파일을 처리하고 있습니다. 잠시만 기다려 주세요.",
        )
        QApplication.processEvents()
        self._set_db_sync_path_controls_busy(True)
        self.ui.plainTextEdit_db_sync_log.setPlainText("실행 중… 잠시만 기다려 주세요.\n")
        params = {
            "spreadsheet_id": _host(self).SPREADSHEET_ID,
            "do_naver": do_naver,
            "do_coupang": do_coupang,
            "test_mode": test_mode,
            "test_count": test_count,
            "naver_path": (
                str(self._db_sync_naver_path_override)
                if self._db_sync_naver_path_override is not None
                else None
            ),
            "coupang_path": (
                str(self._db_sync_coupang_path_override)
                if self._db_sync_coupang_path_override is not None
                else None
            ),
            "verbose_log": verbose_log,
        }
        self._db_sheet_sync_thread = _host(self).DbSheetSyncThread(params)
        self._db_sheet_sync_thread.result_ready.connect(self._on_db_sheet_sync_finished)
        self._db_sheet_sync_thread.start()

    def _on_db_sheet_sync_finished(self, result: dict):
        self._hide_busy_processing_overlay()
        self._set_db_sync_path_controls_busy(False)
        self._db_sheet_sync_thread = None
        lines = result.get("logs") or []
        self.ui.plainTextEdit_db_sync_log.setPlainText("\n".join(lines))
        if result.get("error"):
            err_text = str(result["error"])
            if _host(self)._is_likely_google_sheets_oauth_error(RuntimeError(err_text)):
                handled = self._prompt_google_reauth_for_oauth_error(
                    title="DB동기화",
                    err_text=err_text,
                )
                if not handled:
                    QMessageBox.warning(self, "DB동기화", err_text)
            else:
                QMessageBox.warning(self, "DB동기화", err_text)
        wrote = False
        for key in ("naver", "coupang"):
            ch = result.get(key)
            if ch and ch.get("appended"):
                wrote = True
        if not result.get("test_mode") and wrote and result.get("ok"):
            self._invalidate_google_sheets_client_and_caches()

        done_ch = []
        skipped_ch = []
        for key, label in (("naver", "네이버"), ("coupang", "쿠팡")):
            ch = result.get(key)
            if not ch:
                continue
            if ch.get("skipped_missing_file"):
                skipped_ch.append(label)
            elif ch.get("ok"):
                done_ch.append(label)

        if result.get("error"):
            self.statusBar().showMessage("DB동기화 중 오류가 있습니다.", 8000)
        elif result.get("ok"):
            if skipped_ch and done_ch:
                self.statusBar().showMessage(
                    "DB동기화 완료 · 일부 채널은 로컬 파일 없음으로 건너뜀", 7000
                )
            elif skipped_ch and not done_ch:
                self.statusBar().showMessage(
                    "DB동기화: 선택한 채널에 해당하는 로컬 파일이 없습니다", 7000
                )
            else:
                self.statusBar().showMessage("DB동기화 완료", 5000)
        else:
            self.statusBar().showMessage("DB동기화 중 오류가 있습니다.", 8000)

    def _on_google_disconnect_clicked(self):
        reply = QMessageBox.question(
            self,
            "연결 해제",
            "이 PC에 저장된 Google 로그인 정보(token.json)를 삭제할까요?\n"
            "다음에 스프레드시트 연동 시 다시 브라우저 로그인이 필요합니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        from sheets.google_sheets_oauth import delete_oauth_token_file

        deleted = delete_oauth_token_file()
        self._invalidate_google_sheets_client_and_caches()
        self._clear_google_auth_runtime_error()
        self._refresh_google_auth_status_ui()
        if deleted:
            self.notify("ok", "Google 연결을 해제했습니다(이 PC의 저장된 로그인 정보를 삭제).")
        else:
            self.notify("warn", "삭제할 Google 로그인 정보(token.json)가 없었습니다.")

    def _on_google_reauth_clicked(self, ask_confirm: bool = True):
        if ask_confirm:
            reply = QMessageBox.question(
                self,
                "재인증",
                "브라우저가 열리면 Google 계정으로 로그인·권한 승인을 완료해 주세요.\n"
                "진행 시 기존 token.json은 삭제됩니다.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return False

        if _host(self).gspread is None:
            QMessageBox.warning(self, "재인증", "gspread 패키지가 필요합니다. (pip install gspread)")
            return False
        try:
            from sheets.google_sheets_oauth import delete_oauth_token_file, get_authorized_gspread_client
        except ImportError:
            QMessageBox.warning(self, "재인증", "google-auth-oauthlib 패키지가 필요합니다.")
            return False

        if getattr(self, "_google_reauth_thread", None) is not None:
            self.statusBar().showMessage("재인증: Google 로그인이 이미 진행 중입니다. 열린 브라우저에서 마쳐 주세요.", 5000)
            return True
        delete_oauth_token_file()
        self._invalidate_google_sheets_client_and_caches()
        # 브라우저 로그인은 수 분 걸릴 수 있으므로 백그라운드에서 기다리고 화면은 계속 쓸 수 있게 둔다.
        button = getattr(self.ui, "pushButton_google_reauth", None)
        if button is not None:
            button.setEnabled(False)
        self.statusBar().showMessage("브라우저에서 Google 로그인·권한 승인을 마쳐 주세요. (최대 5분)")
        thread = _host(self).GoogleReauthThread(self)
        self._google_reauth_thread = thread
        thread.result_ready.connect(self._on_google_reauth_finished)
        thread.finished.connect(self._cleanup_google_reauth_thread)
        thread.finished.connect(thread.deleteLater)
        thread.start()
        return True

    def _on_google_reauth_finished(self, payload):
        if payload.get("ok"):
            self._gspread_client = payload.get("client")
            self._clear_google_auth_runtime_error()
            self._refresh_google_auth_status_ui()
            self.statusBar().showMessage("Google Sheets 연동이 완료되었습니다.", 5000)
            self.notify("ok", "Google Sheets 연동이 완료되었습니다.")
        else:
            self._gspread_client = None
            self._refresh_google_auth_status_ui()
            self.statusBar().showMessage("Google 재인증 실패", 5000)
            QMessageBox.critical(self, "재인증 실패", f"Google 로그인에 실패했습니다.\n\n{payload.get('error', '')}")

    def _cleanup_google_reauth_thread(self):
        self._google_reauth_thread = None
        button = getattr(self.ui, "pushButton_google_reauth", None)
        if button is not None:
            button.setEnabled(True)

    def _on_google_open_oauth_folder_clicked(self):
        from sheets.google_sheets_oauth import GOOGLE_AUTH_DIR

        GOOGLE_AUTH_DIR.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(GOOGLE_AUTH_DIR.resolve())))

    @staticmethod
    def _oauth_error_dialog_hint() -> str:
        return (
            "「⚙ 설정」탭 → 「연동」에서 "
            "「연결 해제」 후 「재인증」을 눌러 보세요."
        )

    def _prompt_google_reauth_for_oauth_error(self, *, title: str, err_text: str) -> bool:
        self._mark_google_auth_runtime_error(err_text)
        reply = QMessageBox.question(
            self,
            title,
            "Google Sheets 로그인 토큰이 만료되었거나 철회되었습니다.\n\n"
            f"{err_text}\n\n"
            "지금 저장된 토큰을 삭제하고 재인증하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return False
        return bool(self._on_google_reauth_clicked(ask_confirm=False))

    def _on_kpost_postcode_key_edit_clicked(self):
        title = "주소 검색용 우편번호 API 키"
        regkey, accepted = QInputDialog.getText(
            self, title,
            "우체국 우편번호 조회에 등록된 인증키(regkey)를 입력해 주세요.\n"
            "주소 조회에 성공하면 공유 시트에 저장합니다.",
            QLineEdit.EchoMode.Password,
        )
        if not accepted:
            return
        regkey = regkey.strip()
        if not regkey:
            QMessageBox.warning(self, title, "인증키를 입력해 주세요.")
            return

        try:
            config = _host(self).run_in_background(_host(self)._read_shared_postcode_config)
        except Exception:
            self._show_kpost_api_error(self, (_host(self).KPOST_ERROR_CONFIG_READ, ""))
            return
        items, error = self._request_kpost_postnew(regkey, _host(self).KPOST_POSTCODE_TEST_ADDRESS)
        if error:
            self._show_kpost_api_error(self, error)
            return
        if not any(item["address"] and re.fullmatch(r"[0-9]{5}", item["postcd"]) for item in items):
            QMessageBox.warning(self, title, "검증 주소의 우편번호를 확인하지 못했습니다. 키와 우체국 서비스 상태를 확인해 주세요.")
            return

        try:
            original = config.get(_host(self).CONFIG_KEY_KPOST_POSTCODE_REGKEY, "")
            _host(self).run_in_background(_host(self)._save_shared_postcode_key, regkey, original)
        except ConfigValueChangedError:
            QMessageBox.warning(self, title, "다른 PC에서 공유 키를 변경했습니다. 확인 후 다시 시도해 주세요.")
            return
        except Exception:
            QMessageBox.warning(self, title, "공유 시트에 저장하지 못했습니다. Google Sheets 연결과 편집 권한을 확인해 주세요.")
            return
        self._set_shared_key_status("postcode", "ok", "등록됨 (확인함 · 전원 적용)")
        self.notify("ok", "주소 검색용 우편번호 API 키를 공유 시트에 저장했습니다. 다른 PC에서도 다음 주소 검색부터 적용됩니다.")
