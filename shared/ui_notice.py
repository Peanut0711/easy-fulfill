"""화면 안 알림 띠(확인 창 대신 쓰는 완료·주의 알림).

docs/gui-feedback-tabs-plan.md C-1.

- 탭 영역 아래·상태바 위에 한 줄로 보인다. 종류 표시(✓ ! ✗) · 문구 · 버튼 최대 2개 · 닫기(×).
- 새 알림은 이전 알림을 바꾼다(쌓지 않는다).
- 성공(ok)은 ``OK_AUTO_HIDE_MS`` 뒤 저절로 숨고, 마우스를 올리면 멈춘다. 주의·오류는 닫을 때까지 둔다.
- 색은 ``shared/ui_style.py``의 앱 스타일시트(``QFrame[notice=…]``)에서 정한다.
"""

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QSizePolicy, QToolButton

from shared.ui_style import STATUS_MARKS, _repolish, set_status

NOTICE_KINDS = ("ok", "warn", "error")
OK_AUTO_HIDE_MS = 10_000
MAX_ACTIONS = 2


class NoticeBar(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("frame_notice")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 6, 8, 6)
        layout.setSpacing(8)
        self._mark = QLabel()
        self._mark.setObjectName("label_notice_mark")
        self._text = QLabel()
        self._text.setObjectName("label_notice_text")
        # 긴 파일 이름이 창을 넓히지 않게: 폭은 남는 만큼만 쓰고 넘치면 말줄임(…), 원문은 툴팁.
        self._text.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self._text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self._full_text = ""
        layout.addWidget(self._mark)
        layout.addWidget(self._text, 1)
        self._buttons = []
        for index in range(MAX_ACTIONS):
            button = QPushButton()
            button.setObjectName(f"pushButton_notice_action_{index}")
            button.hide()
            button.clicked.connect(lambda _=False, i=index: self._run_action(i))
            layout.addWidget(button)
            self._buttons.append(button)
        self._close = QToolButton()
        self._close.setObjectName("toolButton_notice_close")
        self._close.setText("×")
        self._close.setAutoRaise(True)
        self._close.setToolTip("알림 닫기")
        self._close.clicked.connect(self.hide)
        layout.addWidget(self._close)
        self._actions = []
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)
        self.kind = ""
        self.hide()

    def show_notice(self, kind, text, actions=(), detail=""):
        """``actions``: ``[(버튼 글자, 함수), …]``(최대 2개). ``detail``은 툴팁(없으면 문구 원문)."""
        if kind not in NOTICE_KINDS:
            raise ValueError(f"알 수 없는 알림: {kind}")
        actions = list(actions)[:MAX_ACTIONS]
        if self.property("notice") != kind:
            self.setProperty("notice", kind)
            _repolish(self)
        self.kind = kind
        set_status(self._mark, kind, STATUS_MARKS[kind], mark=False)
        self._full_text = str(text or "").strip()
        self._text.setToolTip(detail or self._full_text)
        self._actions = actions
        for button, (label, _callback) in zip(self._buttons, actions):
            button.setText(label)
            button.show()
        for button in self._buttons[len(actions):]:
            button.hide()
        self.show()
        self._elide()
        self._timer.stop()
        if kind == "ok":
            self._timer.start(OK_AUTO_HIDE_MS)

    def text(self):
        return self._full_text

    def action_labels(self):
        return [label for label, _callback in self._actions]

    def _run_action(self, index):
        if index < len(self._actions):
            self._actions[index][1]()

    def _elide(self):
        width = max(self._text.width(), 0)
        metrics = self._text.fontMetrics()
        self._text.setText(metrics.elidedText(self._full_text, Qt.TextElideMode.ElideMiddle, width)
                           if width else self._full_text)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._elide()

    def enterEvent(self, event):
        # 읽는 동안에는 숨기지 않는다.
        self._timer.stop()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self.kind == "ok" and self.isVisible():
            self._timer.start(OK_AUTO_HIDE_MS)
        super().leaveEvent(event)
