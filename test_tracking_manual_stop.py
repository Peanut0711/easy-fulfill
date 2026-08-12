from pathlib import Path
from runpy import run_path
from PySide6.QtCore import QFile, QIODevice
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication, QComboBox, QLineEdit, QAbstractItemView, QTableWidget


module = run_path(Path(__file__).with_name("easy-fulfill.py"), run_name="easy_fulfill_test")
row = [""] * 13
row[module["TRACKING_MANAGEMENT_COL"]] = module["TRACKING_MANAGEMENT_MANUAL_STOP"]

assert module["TRACKING_MANAGEMENT_MANUAL_STOP"] in module["TRACKING_MANAGEMENT_EXCLUDED"]
assert module["_tracking_management_state"](row, "배달완료") == module["TRACKING_MANAGEMENT_MANUAL_STOP"]
assert [label for _, label in module["MainWindow"]._TRACKING_TABLE_COLUMNS] == [
    "수취인명", "배송상태", "마지막위치", "최근이벤트", "관리상태", "완료",
    "주문번호", "등기번호", "스토어", "최근조회", "메모",
]
assert module["MainWindow"]._tracking_table_col(0) == 7

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
