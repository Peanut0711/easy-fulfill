"""배송추적 표 보기(알림·상태바 계획 D-1)를 검증한다.

- 열 숨김(기본 「최근조회」·「스토어」, 머리 메뉴로 바꾸면 이 PC 설정에 저장)
- 보이는 열을 표 폭 안에 맞춤(가로 스크롤 없음, 열 머리 글자 폭보다는 줄이지 않음)
- 위험 칩(건수, 누르면 그 위험만, 다시 누르면 모두, 위험이 없어지면 거르기 해제)
- 건수 줄의 마지막 조회 시각·자동 새로고침 간격, 빈 목록 안내
"""

import os
import unittest
from datetime import datetime, timedelta
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QObject
from PySide6.QtWidgets import QApplication, QHBoxLayout, QLabel, QLineEdit, QTableWidget, QWidget

from tracking.controller import (
    SETTING_TRACKING_HIDDEN_COLUMNS, TRACKING_LAST_COLUMN_MIN_WIDTH, TrackingController,
)
from tracking.service import STALE_CONFIG_KEYS, TRACKING_MANAGEMENT_COL


def make_row(index, *, risky=False, location="서울강남우체국", checked=None, name=None):
    row = [""] * 14
    row[0] = f"60812345{index:05d}"
    row[1] = (datetime.now() - timedelta(days=5 if risky else 0)).strftime("%Y-%m-%d %H:%M:%S")
    row[2] = "쿠팡"
    row[3] = f"2026111{index:07d}"
    row[4] = name or f"수취인{index}"
    row[6] = "운송장출력" if risky else "배달준비"
    row[7] = "N"
    row[8] = location
    row[9] = checked or ""
    row[11] = "" if risky else datetime.now().strftime("%Y.%m.%d %H:%M")
    row[TRACKING_MANAGEMENT_COL] = "추적중"
    return row


class Host(QObject):
    def __init__(self, settings=None):
        super().__init__()
        self._tracking_list_thread = None
        self.settings = dict(settings or {})
        self.panel = QWidget()
        layout = QHBoxLayout(self.panel)
        count = QLabel()
        layout.addWidget(count)
        layout.addWidget(QLabel())  # 진행 문구 자리
        self.table = QTableWidget()
        self.table.resize(860, 300)  # 기본 창(900 폭)의 표 폭
        self.ui = SimpleNamespace(
            tableWidget_tracking=self.table, label_tracking_count=count,
            horizontalLayout_tracking_filter=layout, lineEdit_tracking_recipient_search=QLineEdit(),
        )

    def read_app_settings(self):
        return dict(self.settings)

    def get_app_setting(self, key, default=None):
        return self.settings.get(key, default)

    def set_app_setting(self, key, value):
        self.settings[key] = value


class TrackingViewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def controller(self, rows, settings=None):
        host = Host({STALE_CONFIG_KEYS["수거누락"]: 1, **(settings or {})})
        controller = TrackingController(host)
        controller.bind_ui()
        host._tracking_list_values = [["등기번호"] + [""] * 13] + rows
        host._tracking_list_loaded = True
        host.table.show()
        controller._populate_tracking_table()
        return host, controller

    def visible_labels(self, host):
        table = host.table
        return [table.horizontalHeaderItem(c).text() for c in range(table.columnCount())
                if not table.isColumnHidden(c)]

    def test_default_hidden_columns_and_header_menu_saves(self):
        host, controller = self.controller([make_row(1)])
        self.assertNotIn("최근조회", self.visible_labels(host))
        self.assertNotIn("스토어", self.visible_labels(host))
        controller._set_tracking_column_visible("스토어", True)
        self.assertIn("스토어", self.visible_labels(host))
        self.assertEqual(host.settings[SETTING_TRACKING_HIDDEN_COLUMNS], ["최근조회"])
        controller._set_tracking_column_visible("주문번호", False)
        controller._populate_tracking_table()  # 다시 그려도 저장한 대로
        self.assertNotIn("주문번호", self.visible_labels(host))
        controller._reset_tracking_columns()
        self.assertEqual(sorted(host.settings[SETTING_TRACKING_HIDDEN_COLUMNS]), ["스토어", "최근조회"])
        # 수취인명은 숨길 수 없다(설정에 있어도 무시).
        host2, _ = self.controller([make_row(1)], {SETTING_TRACKING_HIDDEN_COLUMNS: ["수취인명"]})
        self.assertIn("수취인명", self.visible_labels(host2))

    def test_columns_fit_table_width(self):
        long_place = "동서울우편집중국(소포) 발송 대기 구역 " * 3
        host, _ = self.controller([make_row(i, location=long_place) for i in range(8)])
        self.app.processEvents()  # 행 번호 열이 생겨 표시 폭이 바뀌면 크기 변경 감시가 다시 맞춘다
        table = host.table
        visible = [c for c in range(table.columnCount()) if not table.isColumnHidden(c)]
        fixed = sum(table.columnWidth(c) for c in visible[:-1])
        self.assertLessEqual(fixed + TRACKING_LAST_COLUMN_MIN_WIDTH, table.viewport().width())
        header = table.horizontalHeader()
        self.assertEqual(header.length(), table.viewport().width())
        self.assertFalse(table.horizontalScrollBar().isVisible())
        for c in visible[:-1]:
            self.assertGreaterEqual(table.columnWidth(c), min(header.sectionSizeHint(c), 220), c)
        place = next(c for c in visible if table.horizontalHeaderItem(c).text() == "마지막위치")
        self.assertEqual(table.item(0, place).toolTip(), long_place)  # 잘린 글자는 툴팁으로

    def test_risk_chips_filter(self):
        rows = [make_row(1, risky=True), make_row(2), make_row(3, risky=True), make_row(4)]
        host, controller = self.controller(rows)
        chips = host._tracking_risk_chips
        self.assertEqual(chips["all"].text(), "⚠ 위험 2")
        self.assertFalse(chips["all"].isHidden())
        self.assertEqual(chips["수거누락"].text(), "수거누락 2")
        self.assertTrue(chips["허브정체"].isHidden())  # 0건 칩은 숨긴다
        self.assertEqual(host.table.rowCount(), 4)
        chips["all"].click()
        self.assertEqual(host.table.rowCount(), 2)
        self.assertTrue(chips["all"].isChecked())
        self.assertIn("(위험만)", host.ui.label_tracking_count.text())
        chips["all"].click()
        self.assertEqual(host.table.rowCount(), 4)
        chips["수거누락"].click()
        self.assertEqual(host.table.rowCount(), 2)
        # 위험이 없어지면 거르기를 풀고 모두 보인다.
        host._tracking_list_values = [host._tracking_list_values[0], make_row(2), make_row(4)]
        controller._populate_tracking_table()
        self.assertEqual(host.table.rowCount(), 2)
        self.assertIsNone(host._tracking_risk_filter)
        self.assertTrue(chips["all"].isHidden())

    def test_count_line_shows_last_check_and_auto_interval(self):
        checked = datetime.now().replace(hour=10, minute=20, second=5)
        rows = [make_row(1, checked=checked.strftime("%Y-%m-%d %H:%M:%S")),
                make_row(2, checked=(checked - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S"))]
        host, _ = self.controller(rows)
        text = host.ui.label_tracking_count.text()
        self.assertIn("표시 2건", text)
        self.assertIn("마지막 조회 10:20", text)
        self.assertIn("자동 60분", text)
        host2, _ = self.controller(rows, {"tracking_auto_refresh_min": "0"})
        self.assertIn("자동 꺼짐", host2.ui.label_tracking_count.text())

    def test_empty_list_message(self):
        host, controller = self.controller([make_row(1, name="김민수")])
        label = host._tracking_empty_label
        self.assertTrue(label.isHidden())
        host.ui.lineEdit_tracking_recipient_search.setText("없는사람")
        controller._populate_tracking_table()
        self.assertFalse(label.isHidden())
        self.assertIn("「없는사람」 수취인을 찾지 못했습니다", label.text())
        host.ui.lineEdit_tracking_recipient_search.setText("")
        host._tracking_list_values = [host._tracking_list_values[0]]
        controller._populate_tracking_table()
        self.assertEqual(label.text(), "배송중인 송장이 없습니다.")
        # 목록을 아직 못 읽었으면(읽기 실패 등) 「없습니다」로 보이지 않는다.
        host._tracking_list_loaded = False
        controller._populate_tracking_table()
        self.assertTrue(label.isHidden())


if __name__ == "__main__":
    unittest.main()
