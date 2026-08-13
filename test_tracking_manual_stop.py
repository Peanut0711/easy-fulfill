import os
from datetime import datetime
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QFile, QIODevice
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication, QAbstractItemView, QComboBox, QLineEdit, QTableWidget

from tracking.service import (
    STALE_DEFAULTS,
    TRACKING_MANAGEMENT_COL,
    TRACKING_MANAGEMENT_EXCLUDED,
    TRACKING_MANAGEMENT_MANUAL_STOP,
    build_risk_digest_text,
    business_elapsed_hours,
    evaluate_risk,
    risk_bucket,
    tracking_management_state,
)


row = [""] * 13
row[TRACKING_MANAGEMENT_COL] = TRACKING_MANAGEMENT_MANUAL_STOP

assert TRACKING_MANAGEMENT_MANUAL_STOP in TRACKING_MANAGEMENT_EXCLUDED
assert tracking_management_state(row, "배달완료") == TRACKING_MANAGEMENT_MANUAL_STOP
assert risk_bucket("운송장출력", "", False, TRACKING_MANAGEMENT_MANUAL_STOP) is None
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

app = QApplication.instance() or QApplication([])
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
