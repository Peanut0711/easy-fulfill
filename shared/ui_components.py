"""화면 공통 부품(docs/ui-design-guide.md 「부품」). 모양은 테마(shared/theme.py)의 동적 속성 규칙으로 그린다.

- ``SectionCard``: 구획 하나(흰 카드, 제목 줄, 오른쪽 보조 동작). QGroupBox·파란 번호 제목을 대신한다.
- ``ButtonRow``: 왼쪽 초기화·위험 동작, 오른쪽 보조 → 주 동작. 버튼은 글자 폭보다 늘리지 않는다.
- ``FormGrid``: 라벨 열 · 입력 열 · 보조 버튼 열, 필수 표시, 칸 아래 오류 문구.
- ``EmptyState``: 빈 표·목록 자리에 무엇을 하면 채워지는지 한 줄과 주 동작 하나.
- ``setup_table``: 표 머리글·행 높이·선택·격자·말줄임 공통 설정. ``align_numeric``: 숫자 칸 오른쪽 맞춤.
- ``confirm_action``: 「예/아니요」 대신 동작 이름 버튼을 쓰는 확인 창(되돌릴 수 없는 동작은 「취소」가 기본).
"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from shared import theme
from shared.ui_style import set_button_role


def _repolish(widget):
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


def fixed_width_button(button):
    """버튼이 레이아웃에서 글자 폭보다 늘어나지 않게 한다(폭 나누기 금지)."""
    policy = button.sizePolicy()
    policy.setHorizontalPolicy(QSizePolicy.Policy.Fixed)
    button.setSizePolicy(policy)
    return button


class SectionCard(QFrame):
    """구획 카드. ``body``(QVBoxLayout)에 내용을 넣는다. 제목·단계·동작이 없으면 제목 줄을 숨긴다."""

    def __init__(self, title="", step=None, parent=None, actions=()):
        super().__init__(parent)
        self.setProperty("card", True)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
        outer.setSpacing(theme.SPACE_SM)

        self.header_widget = QWidget()
        self.header = QHBoxLayout(self.header_widget)
        self.header.setContentsMargins(0, 0, 0, 0)
        self.header.setSpacing(theme.SPACE_SM)
        self._step = step
        self.step_badge = QLabel("" if step is None else str(step))
        self.step_badge.setProperty("stepBadge", True)
        self.step_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.step_badge.setFixedSize(22, 22)
        self.step_badge.setVisible(step is not None)
        self.title_label = QLabel(title)
        self.title_label.setProperty("sectionTitle", True)
        self.header.addWidget(self.step_badge)
        self.header.addWidget(self.title_label)
        self.header.addStretch(1)
        self._actions = []
        outer.addWidget(self.header_widget)

        self.body = QVBoxLayout()
        self.body.setSpacing(theme.SPACE_SM)
        outer.addLayout(self.body)
        for action in actions:
            self.add_action(action)
        self._sync_header()

    def _sync_header(self):
        self.header_widget.setVisible(bool(self.title_label.text()) or self._step is not None or bool(self._actions))

    def add_action(self, widget):
        """제목 줄 오른쪽에 보조 동작(버튼 등)을 붙인다."""
        if isinstance(widget, QPushButton):
            fixed_width_button(widget)
        self.header.addWidget(widget)
        self._actions.append(widget)
        self._sync_header()
        return widget

    def set_title(self, text):
        self.title_label.setText(text)
        self._sync_header()

    def set_step_done(self, done):
        """단계 배지를 완료(✓, 초록) 또는 원래 번호로 보인다."""
        if self._step is None:
            return
        self.step_badge.setText("✓" if done else str(self._step))
        state = "done" if done else ""
        if (self.step_badge.property("stepState") or "") != state:
            self.step_badge.setProperty("stepState", state)
            _repolish(self.step_badge)


class ButtonRow(QWidget):
    """버튼 줄: ``left``(초기화·위험)는 왼쪽, ``right``(보조 → 주 동작)는 오른쪽 끝."""

    def __init__(self, left=(), right=(), parent=None):
        super().__init__(parent)
        self.layout_ = QHBoxLayout(self)
        self.layout_.setContentsMargins(0, 0, 0, 0)
        self.layout_.setSpacing(theme.SPACE_SM)
        for button in left:
            self.layout_.addWidget(fixed_width_button(button))
        self.layout_.addStretch(1)
        for button in right:
            self.layout_.addWidget(fixed_width_button(button))

    def buttons(self):
        items = (self.layout_.itemAt(i).widget() for i in range(self.layout_.count()))
        return [w for w in items if w is not None]


class FormGrid(QGridLayout):
    """라벨 · 입력 · 보조 버튼 3열 격자. ``add_row``가 행 번호를 돌려준다."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHorizontalSpacing(theme.SPACE_SM)
        self.setVerticalSpacing(theme.SPACE_SM)
        self.setColumnStretch(1, 1)
        self._rows = {}

    def add_row(self, label_text, field, extra=None, required=False):
        row = self.rowCount() if self.count() else 0
        label = QLabel(f"{label_text} *" if required else label_text)
        label.setProperty("formLabel", True)
        if required:
            label.setToolTip("필수 입력")
        label.setBuddy(field)
        error = QLabel()
        error.setProperty("fieldError", True)
        error.setWordWrap(True)
        error.setVisible(False)
        self.addWidget(label, row, 0, Qt.AlignmentFlag.AlignVCenter)
        if extra is None:
            self.addWidget(field, row, 1, 1, 2)
        else:
            self.addWidget(field, row, 1)
            if isinstance(extra, QPushButton):
                fixed_width_button(extra)
            self.addWidget(extra, row, 2)
        self.addWidget(error, row + 1, 1, 1, 2)
        self._rows[field] = (label, error, required)
        return row

    def label_for(self, field):
        return self._rows[field][0]

    def set_missing(self, field, missing):
        """필수 칸이 비었으면 라벨을 주의 색으로 보인다."""
        label = self._rows[field][0]
        state = "warn" if missing else ""
        if (label.property("status") or "") != state:
            label.setProperty("status", state)
            _repolish(label)

    def set_error(self, field, text=""):
        """칸 아래 오류 문구를 보이거나(``text``) 숨긴다."""
        error = self._rows[field][1]
        error.setText(text)
        error.setVisible(bool(text))


class EmptyState(QWidget):
    """빈 표·목록 자리 안내. ``action``(QPushButton)은 주 동작으로 보인다."""

    def __init__(self, text, action=None, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(theme.SPACE_LG, theme.SPACE_LG, theme.SPACE_LG, theme.SPACE_LG)
        layout.setSpacing(theme.SPACE_MD)
        layout.addStretch(1)
        self.label = QLabel(text)
        self.label.setProperty("emptyText", True)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setWordWrap(True)
        layout.addWidget(self.label)
        self.action = action
        if action is not None:
            set_button_role(action, "primary")
            row = QHBoxLayout()
            row.addStretch(1)
            row.addWidget(fixed_width_button(action))
            row.addStretch(1)
            layout.addLayout(row)
        layout.addStretch(1)

    def set_text(self, text):
        self.label.setText(text)


NUMERIC_ALIGNMENT = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter


def align_numeric(item):
    """표 항목(QTableWidgetItem·QStandardItem)을 숫자 칸 정렬(오른쪽)로 한다."""
    item.setTextAlignment(NUMERIC_ALIGNMENT)
    return item


def setup_table(table, *, select_rows=True, single=True, stretch_last=True):
    """표 공통 설정: 행 단위 선택, 세로 격자선 없음, 행 높이, 끝 말줄임, 행 번호 숨김."""
    table.setShowGrid(False)
    table.setProperty("dataTable", True)
    table.setWordWrap(False)
    table.setTextElideMode(Qt.TextElideMode.ElideRight)
    if select_rows:
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    if single:
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    vertical = getattr(table, "verticalHeader", None)
    if vertical is not None:
        vertical().setVisible(False)
        vertical().setDefaultSectionSize(theme.CONTROL_HEIGHT_COMPACT + 4)
    header = getattr(table, "horizontalHeader", None)
    if header is not None:
        header().setStretchLastSection(stretch_last)
        header().setHighlightSections(False)
    return table


def confirm_action(parent, title, text, action_text, *, destructive=False, detail="", cancel_text="취소"):
    """동작 이름 버튼(``action_text``)을 누르면 True. 되돌릴 수 없는 동작이면 위험 버튼, 기본은 「취소」."""
    box = QMessageBox(parent)
    box.setWindowTitle(title)
    box.setIcon(QMessageBox.Icon.Warning if destructive else QMessageBox.Icon.Question)
    box.setText(text)
    if detail:
        box.setInformativeText(detail)
    action = box.addButton(action_text, QMessageBox.ButtonRole.AcceptRole)
    cancel = box.addButton(cancel_text, QMessageBox.ButtonRole.RejectRole)
    set_button_role(action, "danger" if destructive else "primary")
    box.setDefaultButton(cancel if destructive else action)
    box.setEscapeButton(cancel)
    box.exec()
    return box.clickedButton() is action
