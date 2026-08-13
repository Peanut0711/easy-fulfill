import os
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QFile, QIODevice, QObject
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QLineEdit,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from tracking.service import (
    STALE_DEFAULTS,
    TRACKING_MANAGEMENT_COL,
    TRACKING_MANAGEMENT_EXCLUDED,
    TRACKING_MANAGEMENT_MANUAL_STOP,
    build_risk_digest_text,
    business_elapsed_hours,
    evaluate_risk,
    risk_bucket,
    select_tracking_list_row_numbers,
    tracking_management_state,
)
from tracking.repository import (
    TRACKING_DETAIL_READ_BATCH_SIZE,
    read_tracking_list_metadata,
    read_tracking_rows,
    update_tracking_management,
    update_tracking_notes,
)
from tracking.controller import TrackingController
from tracking.workers import (
    DigestSendThread,
    SlackSendThread,
    TrackingManagementUpdateThread,
    TrackingNotesUpdateThread,
    TrackingKeyValidateThread,
    TrackingRefreshThread,
)


row = [""] * 13
row[TRACKING_MANAGEMENT_COL] = TRACKING_MANAGEMENT_MANUAL_STOP

assert TRACKING_MANAGEMENT_MANUAL_STOP in TRACKING_MANAGEMENT_EXCLUDED
assert tracking_management_state(row, "배달완료") == TRACKING_MANAGEMENT_MANUAL_STOP
assert risk_bucket("운송장출력", "", False, TRACKING_MANAGEMENT_MANUAL_STOP) is None
assert select_tracking_list_row_numbers(
    ["2026-07-01 09:00:00", "2026-08-13 09:00:00", "2026-08-12 09:00:00"],
    ["N", "Y", "Y"],
    ["2026.07.01 10:00", "2026.08.13 11:00", "2026.07.02 10:00"],
    ["추적중", "추적중", "추적중"],
    "배송중",
    datetime(2026, 8, 13, 12),
) == [2]
assert select_tracking_list_row_numbers(
    ["2026-07-01 09:00:00", "2026-08-13 09:00:00", "2026-08-12 09:00:00"],
    ["N", "Y", "Y"],
    ["2026.07.01 10:00", "2026.08.13 11:00", "2026.07.02 10:00"],
    ["추적중", "추적중", "수동 중지"],
    "완료",
    datetime(2026, 8, 13, 12),
) == [3]
assert business_elapsed_hours(
    datetime(2026, 8, 7, 18), datetime(2026, 8, 10, 10)
) == 16.0

risk, elapsed_h = evaluate_risk(
    "운송장출력",
    "",
    False,
    datetime(2026, 8, 3, 9),
    datetime(2026, 8, 5, 10),
    STALE_DEFAULTS,
    24,
)
assert risk == "수거누락"
assert elapsed_h == 49.0

digest = build_risk_digest_text(
    [{
        "regino": "123",
        "name": "홍길동",
        "status": "운송장출력",
        "where": "",
        "event_time": "",
        "elapsed_h": 49.0,
        "category": "수거누락",
    }],
    STALE_DEFAULTS,
    datetime(2026, 8, 5, 10),
)
assert digest == (
    "⚠️ 배송 위험 1건 (평일 기준 무이동 · 2026-08-05 10:00)\n"
    "[🟠 수거 누락 의심 · 기준 24h+]\n"
    "• 123 홍길동 — 운송장출력 @ 위치미상 (49시간째 무이동)"
)


class _TrackingWorksheetStub:
    def __init__(self):
        self.batch_calls = []

    def get_all_values(self):
        return [["등기번호"], ["123"]]

    def batch_update(self, updates, value_input_option):
        assert value_input_option == "RAW"
        self.batch_calls.append(updates)


worksheet = _TrackingWorksheetStub()
assert update_tracking_management(worksheet, ["123"], "수동 중지") == {
    "updated": 1, "missing": 0,
}
assert worksheet.batch_calls[-1] == [{"range": "M2", "values": [["수동 중지"]]}]
assert update_tracking_notes(worksheet, {"123": "테스트 메모"}) == 1
assert worksheet.batch_calls[-1] == [{"range": "K2", "values": [["테스트 메모"]]}]


class _TrackingListWorksheetStub:
    def __init__(self):
        self.ranges = []

    def batch_get(self, ranges):
        self.ranges.append(ranges)
        if ranges[0] == "A1:M1":
            return [
                [["등기번호", "등록일시"]],
                [["2026-08-13 09:00:00"], ["2026-08-12 09:00:00"]],
                [["N"], ["Y"]],
                [["2026.08.13 10:00"], ["2026.08.12 10:00"]],
                [["추적중"], ["수동 중지"]],
            ]
        return [[["123", "2026-08-13 09:00:00"]], [["456", "2026-08-12 09:00:00"]]]


list_worksheet = _TrackingListWorksheetStub()
header, metadata = read_tracking_list_metadata(list_worksheet)
assert header == ["등기번호", "등록일시"]
assert metadata[0] == ["2026-08-13 09:00:00", "2026-08-12 09:00:00"]
assert read_tracking_rows(list_worksheet, [2, 3]) == [
    ["123", "2026-08-13 09:00:00"],
    ["456", "2026-08-12 09:00:00"],
]
assert list_worksheet.ranges == [
    ["A1:M1", "B2:B", "H2:H", "L2:L", "M2:M"],
    ["A2:M2", "A3:M3"],
]


class _TrackingListBatchWorksheetStub:
    def __init__(self):
        self.ranges = []

    def batch_get(self, ranges):
        self.ranges.append(ranges)
        return [[[item]] for item in ranges]


batch_worksheet = _TrackingListBatchWorksheetStub()
row_numbers = list(range(2, TRACKING_DETAIL_READ_BATCH_SIZE + 4))
assert read_tracking_rows(batch_worksheet, row_numbers) == [
    [f"A{row}:M{row}"] for row in row_numbers
]
assert [len(ranges) for ranges in batch_worksheet.ranges] == [
    TRACKING_DETAIL_READ_BATCH_SIZE, 2,
]

assert hasattr(TrackingRefreshThread, "progress")
assert hasattr(TrackingManagementUpdateThread, "result_ready")
assert hasattr(TrackingNotesUpdateThread, "result_ready")
assert hasattr(SlackSendThread, "result_ready")
assert hasattr(DigestSendThread, "result_ready")
assert hasattr(TrackingKeyValidateThread, "result_ready")
assert TrackingController._tracking_table_col(10) == 10
assert TrackingController._parse_event_dt("2026.08.05 10:00") == datetime(2026, 8, 5, 10)
assert hasattr(TrackingController, "_maybe_send_daily_digest")
assert hasattr(TrackingController, "_begin_tracking_config_read")

app = QApplication.instance() or QApplication([])


class _TrackingControllerHost(QObject):
    def __init__(self):
        super().__init__()
        self.ui = object()
        self._tracking_list_thread = None

    def get_app_setting(self, _key, default=None):
        return default


admin_host = _TrackingControllerHost()
admin_widget = QWidget()
admin_layout = QVBoxLayout(admin_widget)
admin_controller = TrackingController(admin_host)
admin_controller.build_admin_slack_controls(admin_layout)
admin_controller.build_admin_risk_controls(admin_layout)
assert admin_layout.count() == 3
assert admin_host._dlg_slack_auto is not None
assert set(admin_host._dlg_stale_boxes) == {"허브정체", "수거누락", "이동정체"}
assert admin_host._dlg_auto_refresh is not None


loading_host = _TrackingControllerHost()
loading_host.ui = SimpleNamespace(tableWidget_tracking=QTableWidget())
loading_host.ui.tableWidget_tracking.resize(400, 200)
loading_controller = TrackingController(loading_host)
loading_controller.bind_ui()
loading_controller._set_tracking_table_loading(True)
overlay = loading_host._tracking_loading_overlay
assert overlay.parentWidget() is loading_host.ui.tableWidget_tracking.viewport()
assert loading_host._tracking_loading_message.text() == "목록 불러오는 중…"
assert overlay.geometry() == loading_host.ui.tableWidget_tracking.viewport().rect()
loading_controller._set_tracking_table_loading(False)
assert overlay.isHidden()


def _config_read_render_calls(list_thread):
    host = _TrackingControllerHost()
    host._tracking_list_thread = list_thread
    controller = TrackingController(host)
    calls = []
    controller._refresh_key_status = lambda: None
    controller._refresh_slack_status = lambda: None
    controller._sync_stale_threshold_spinboxes = lambda: None
    controller._sync_inquiry_hours_spinboxes = lambda: None
    controller._sync_auto_refresh_spinbox = lambda: None
    controller._reconfigure_auto_refresh_timer = lambda: None
    controller._populate_tracking_table = lambda: calls.append("populate")
    controller._on_tracking_config_read_finished({"ok": False})
    return calls


# 설정 Sheet가 목록보다 먼저 끝나도 빈 표를 그리지 않아야 한다.
assert _config_read_render_calls(object()) == []
# 목록이 이미 준비된 뒤 설정을 읽었다면 위험도 기준을 다시 반영한다.
assert _config_read_render_calls(None) == ["populate"]

file = QFile(str(Path(__file__).with_name("ui") / "main_window.ui"))
assert file.open(QIODevice.OpenModeFlag.ReadOnly)
ui = QUiLoader().load(file)
combo = ui.findChild(QComboBox, "comboBox_tracking_filter")
assert [combo.itemText(i) for i in range(combo.count())] == [
    "배송중", "오늘", "추적 중지", "폐기 후보", "폐기", "완료", "전체",
]
search = ui.findChild(QLineEdit, "lineEdit_tracking_recipient_search")
assert search.placeholderText() == "수취인명 검색"
assert search.isClearButtonEnabled()
table = ui.findChild(QTableWidget, "tableWidget_tracking")
assert table.editTriggers() == QAbstractItemView.EditTrigger.DoubleClicked
