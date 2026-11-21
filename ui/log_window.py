"""「도움말 › 프로그램 로그」 창: shared/app_log에 모인 줄을 시각과 함께 보여 준다.

모달이 아니어서 창을 띄워 둔 채 작업할 수 있고, 열려 있는 동안 새 줄을 주기적으로 붙인다.
"""

from PySide6.QtCore import QTimer
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QPlainTextEdit, QPushButton, QVBoxLayout

from shared import app_log
from shared.dialog_geometry import apply_dialog_geometry, save_dialog_geometry

LOG_WINDOW_GEOMETRY_KEY = "log_window_geometry"
LOG_WINDOW_BASE_SIZE = (760, 480)
LOG_WINDOW_REFRESH_MS = 300


def format_entry(entry):
    _seq, stamp, text, is_error = entry
    return f"{stamp}  {'[오류] ' if is_error and text.strip() else ''}{text}"


class LogWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dialog_program_log")
        self.setWindowTitle("프로그램 로그")
        self.setModal(False)
        self._last_seq = 0

        layout = QVBoxLayout(self)
        self.text = QPlainTextEdit(self)
        self.text.setObjectName("plainTextEdit_program_log")
        self.text.setReadOnly(True)
        self.text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.text.setMaximumBlockCount(app_log.MAX_ENTRIES)
        layout.addWidget(self.text)

        buttons = QDialogButtonBox(self)
        self.button_copy = QPushButton("전체 복사")
        self.button_copy.setToolTip("로그 전체를 클립보드에 복사(문제 문의할 때 붙여 넣기)")
        self.button_copy.clicked.connect(self.copy_all)
        buttons.addButton(self.button_copy, QDialogButtonBox.ButtonRole.ActionRole)
        self.button_clear = QPushButton("지우기")
        self.button_clear.setToolTip("지금까지 모인 로그를 지운다(파일로 저장된 것은 없음)")
        self.button_clear.clicked.connect(self.clear_log)
        buttons.addButton(self.button_clear, QDialogButtonBox.ButtonRole.ActionRole)
        close = buttons.addButton("닫기", QDialogButtonBox.ButtonRole.RejectRole)
        close.clicked.connect(self.close)
        layout.addWidget(buttons)

        self._timer = QTimer(self)
        self._timer.setInterval(LOG_WINDOW_REFRESH_MS)
        self._timer.timeout.connect(self.refresh)
        apply_dialog_geometry(self, LOG_WINDOW_GEOMETRY_KEY, LOG_WINDOW_BASE_SIZE)

    def refresh(self):
        entries = app_log.entries_since(self._last_seq)
        if not entries:
            return
        self._last_seq = entries[-1][0]
        bar = self.text.verticalScrollBar()
        at_bottom = bar.value() >= bar.maximum() - 2
        self.text.appendPlainText("\n".join(format_entry(entry) for entry in entries))
        if at_bottom:
            bar.setValue(bar.maximum())

    def copy_all(self):
        QGuiApplication.clipboard().setText(self.text.toPlainText())

    def clear_log(self):
        app_log.clear()
        self.text.clear()

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh()
        self._timer.start()

    def hideEvent(self, event):
        self._timer.stop()
        save_dialog_geometry(self, LOG_WINDOW_GEOMETRY_KEY)
        super().hideEvent(event)
