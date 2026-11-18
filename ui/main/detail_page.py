"""메인 창 「상세페이지」 탭: 화면 구성, 쿠팡 WING 로그인 세션, 상세 HTML 만들기·미리보기·편집, 쿠팡 상품에 HTML 채우기, 작업 로그."""

import json
import re
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QProcess, QProcessEnvironment, Qt, QTimer, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressDialog,
    QPushButton,
    QScrollArea,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from shared.paths import APP_ROOT
from shared import theme
from shared.ui_components import ButtonRow, FormGrid, SectionCard, fixed_width_button
from shared.ui_style import set_button_role
from ui.main import host as _host


class DetailPageMixin:
    """MainWindow 믹스인. 메인 창 「상세페이지」 탭: 화면 구성, 쿠팡 WING 로그인 세션, 상세 HTML 만들기·미리보기·편집, 쿠팡 상품에 HTML 채우기, 작업 로그."""

    def _mount_detail_page_tab(self):
        """단건 상세 전환만 제공한다. 비밀번호·일괄 저장은 의도적으로 제외한다."""
        tabs = self.ui.tabWidget
        page = QWidget()
        page.setObjectName("tab_detail_page")
        # 내용 최소 높이(약 600px)가 창 크기 예산보다 커서 스크롤 영역에 넣는다(창이 커지지 않게).
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea(page)
        scroll.setObjectName("scrollArea_detail_page")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        content = QWidget()
        content.setObjectName("widget_detail_page_content")
        scroll.setWidget(content)
        # 스크롤 영역이 창 배경색(회색)을 칠해 다른 탭(흰색)과 달라 보이지 않게 한다(설정·문서 탭과 같은 방식).
        scroll.viewport().setAutoFillBackground(False)
        content.setAutoFillBackground(False)
        page_layout.addWidget(scroll)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
        layout.setSpacing(theme.SPACE_MD)

        # 디자인 개편 4-5: 단계마다 카드(SectionCard). 끝난 단계는 번호 배지가 ✓(초록)로 바뀐다
        # (_set_detail_step_done). 입력 초기화·상품번호 변경 때 ②③을 지운다.
        self._detail_step_titles = {
            1: "쿠팡 WING 연결",
            2: "상품 상세 HTML 만들기",
            3: "쿠팡 상품 HTML 적용",
        }
        self._detail_steps_done = set()
        self._detail_steps_product = ""

        def hint(text):
            label = QLabel(text)
            label.setWordWrap(True)
            label.setProperty("hint", "true")
            return label

        session_box = SectionCard(self._detail_step_titles[1], step=1)
        self.label_detail_session = QLabel(self._detail_session_label())
        self.label_detail_session.setWordWrap(True)
        self.pushButton_detail_login = QPushButton("쿠팡 로그인 연결/갱신")
        self.pushButton_detail_clear_session = QPushButton("로그인 세션 삭제")
        set_button_role(self.pushButton_detail_clear_session, "danger")
        session_row = QHBoxLayout()
        session_row.setSpacing(theme.SPACE_SM)
        session_row.addWidget(self.label_detail_session, 1)
        session_row.addWidget(fixed_width_button(self.pushButton_detail_clear_session))
        session_row.addWidget(fixed_width_button(self.pushButton_detail_login))
        session_box.body.addLayout(session_row)

        # ② 대상 상품: ②(만들기)와 ③(적용)이 같은 네이버 상품번호 칸 하나를 읽는다. 처음 입력하는 ② 카드 안에 둔다.
        create_box = SectionCard(self._detail_step_titles[2], step=2)
        self.lineEdit_detail_naver = QLineEdit()
        self.lineEdit_detail_naver.setObjectName("lineEdit_detail_naver")
        self.lineEdit_detail_naver.setPlaceholderText("네이버 스마트스토어 상품번호 또는 상품 주소")
        create_grid = FormGrid()
        create_grid.add_row("네이버 상품번호", self.lineEdit_detail_naver)
        create_grid.label_for(self.lineEdit_detail_naver).setToolTip("③ 적용에서도 이 상품을 씁니다.")
        create_box.body.addLayout(create_grid)
        self.pushButton_detail_create = QPushButton("HTML 생성")
        self.pushButton_detail_open_preview = QPushButton("미리보기")
        self.pushButton_detail_open_html = QPushButton("HTML 편집")
        self.pushButton_detail_open_folder = QPushButton("결과 폴더 열기")
        set_button_role(self.pushButton_detail_create, "primary")
        create_box.body.addWidget(ButtonRow(
            left=[self.pushButton_detail_open_folder],
            right=[self.pushButton_detail_open_preview, self.pushButton_detail_open_html,
                   self.pushButton_detail_create]))
        create_box.body.addWidget(hint("쿠팡 신규 등록 화면에는 생성된 HTML을 사용자가 직접 붙여 넣습니다."))
        self.detail_create_progress = None

        replace_box = SectionCard(self._detail_step_titles[3], step=3)
        self.pushButton_detail_clear_inputs = QPushButton("입력 초기화")
        self.lineEdit_detail_vendor_inventory = QLineEdit()
        self.lineEdit_detail_vendor_inventory.setPlaceholderText("쿠팡 등록상품 ID")
        replace_grid = FormGrid()
        replace_grid.add_row("쿠팡 등록상품 ID", self.lineEdit_detail_vendor_inventory)
        replace_box.body.addLayout(replace_grid)
        self.pushButton_detail_stage = QPushButton("WING에 HTML 채우기")
        self.pushButton_detail_finish = QPushButton("WING 작업 종료")
        self.pushButton_detail_finish.setEnabled(False)
        set_button_role(self.pushButton_detail_stage, "primary")
        replace_box.body.addWidget(ButtonRow(
            left=[self.pushButton_detail_clear_inputs],
            right=[self.pushButton_detail_finish, self.pushButton_detail_stage]))
        replace_box.body.addWidget(hint(
            "기본 등록 → HTML 작성 전환과 HTML 입력까지만 자동 수행합니다. 저장은 WING에서 직접 누르세요."))

        # 작업 로그 카드: 기본 접힘. 오류가 기록되면 자동으로 펼친다(_detail_log(error=True)).
        log_box = SectionCard()
        self.toolButton_detail_log_toggle = QToolButton()
        self.toolButton_detail_log_toggle.setObjectName("toolButton_detail_log_toggle")
        self.toolButton_detail_log_toggle.setCheckable(True)
        self.toolButton_detail_log_toggle.setAutoRaise(True)
        self.toolButton_detail_log_toggle.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toolButton_detail_log_toggle.setProperty("sectionTitle", True)
        log_box.header.insertWidget(0, self.toolButton_detail_log_toggle)
        self.pushButton_detail_log_clear = QPushButton("로그 지우기")
        self.pushButton_detail_log_copy = QPushButton("로그 복사")
        self.pushButton_detail_log_save = QPushButton("로그 저장")
        for button in (self.pushButton_detail_log_clear, self.pushButton_detail_log_copy,
                       self.pushButton_detail_log_save):
            log_box.add_action(button)
        self.textEdit_detail_log = QPlainTextEdit()
        self.textEdit_detail_log.setReadOnly(True)
        self.textEdit_detail_log.setPlaceholderText("상세페이지 작업 로그")
        self.textEdit_detail_log.setMinimumHeight(150)
        log_box.body.addWidget(self.textEdit_detail_log, 1)
        self._detail_step_boxes = {1: session_box, 2: create_box, 3: replace_box}
        layout.addWidget(session_box)
        layout.addWidget(create_box)
        layout.addWidget(replace_box)
        layout.addWidget(log_box)
        layout.addStretch(1)  # 로그를 접었을 때 빈 카드가 늘어나지 않게 남는 공간은 아래에 둔다
        tabs.addTab(page, "상세페이지")
        self.toolButton_detail_log_toggle.toggled.connect(self._set_detail_log_expanded)
        self._set_detail_log_expanded(False)

        self.pushButton_detail_login.clicked.connect(self._on_detail_login_clicked)
        self.pushButton_detail_clear_session.clicked.connect(self._on_detail_clear_session_clicked)
        self.pushButton_detail_create.clicked.connect(self._on_detail_create_clicked)
        self.pushButton_detail_open_preview.clicked.connect(self._on_detail_open_preview_clicked)
        self.pushButton_detail_open_html.clicked.connect(self._on_detail_open_html_clicked)
        self.pushButton_detail_open_folder.clicked.connect(self._on_detail_open_folder_clicked)
        self.pushButton_detail_stage.clicked.connect(self._on_detail_stage_clicked)
        self.pushButton_detail_finish.clicked.connect(self._on_detail_finish_clicked)
        self.pushButton_detail_clear_inputs.clicked.connect(self._on_detail_clear_inputs_clicked)
        self.lineEdit_detail_naver.textChanged.connect(
            lambda value: self._normalize_detail_id_input(
                self.lineEdit_detail_naver, value, "naver"
            )
        )
        self.lineEdit_detail_naver.textChanged.connect(self._on_detail_target_changed)
        self.lineEdit_detail_vendor_inventory.textChanged.connect(
            lambda value: self._normalize_detail_id_input(
                self.lineEdit_detail_vendor_inventory, value, "coupang"
            )
        )
        self.pushButton_detail_log_clear.clicked.connect(self.textEdit_detail_log.clear)
        self.pushButton_detail_log_copy.clicked.connect(self._on_detail_log_copy_clicked)
        self.pushButton_detail_log_save.clicked.connect(self._on_detail_log_save_clicked)

    DETAIL_LOG_ERROR_MARKERS = ("[오류]", "Traceback (most recent call last)")

    def _detail_log(self, text, error=False):
        self.textEdit_detail_log.appendPlainText(text.rstrip())
        if error or any(marker in text for marker in self.DETAIL_LOG_ERROR_MARKERS):
            self._set_detail_log_expanded(True)

    def _set_detail_log_expanded(self, expanded):
        button = self.toolButton_detail_log_toggle
        if button.isChecked() != expanded:
            button.blockSignals(True)
            button.setChecked(expanded)
            button.blockSignals(False)
        button.setText("작업 로그 ▾" if expanded else "작업 로그 ▸")
        button.setToolTip("작업 로그 접기" if expanded else "작업 로그 펼치기")
        self.textEdit_detail_log.setVisible(expanded)

    def _set_detail_step_done(self, step, done=True):
        """단계 카드의 번호 배지를 ✓로 바꾸거나 되돌린다(1: WING 연결, 2: HTML 만들기, 3: WING에 채우기)."""
        if done:
            self._detail_steps_done.add(step)
            if step in (2, 3):
                text = self.lineEdit_detail_naver.text().strip()
                self._detail_steps_product = self._detail_id_from_input(text, "naver") or text
        else:
            self._detail_steps_done.discard(step)
        self._detail_step_boxes[step].set_step_done(step in self._detail_steps_done)

    def _on_detail_target_changed(self, _value):
        # 다른 상품으로 바꾸면 그 상품의 ②③은 아직 하지 않은 것이다. 주소를 붙여 넣으면 같은 신호 안에서
        # 번호로 바뀌므로, 신호 값이 아니라 칸의 현재 값을 번호로 맞춰 비교한다.
        text = self.lineEdit_detail_naver.text().strip()
        current = self._detail_id_from_input(text, "naver") or text
        if current != self._detail_steps_product:
            self._set_detail_step_done(2, False)
            self._set_detail_step_done(3, False)

    def _on_detail_clear_inputs_clicked(self):
        self.lineEdit_detail_naver.clear()
        self.lineEdit_detail_vendor_inventory.clear()
        self._set_detail_step_done(2, False)
        self._set_detail_step_done(3, False)

    @staticmethod
    def _detail_upload_progress(output):
        matches = re.findall(r"\[(\d+)/(\d+)\]", output)
        return tuple(map(int, matches[-1])) if matches else None

    def _detail_session_label(self):
        profile = APP_ROOT / "output" / "coupang-browser-profile"
        if not (profile / "easy-fulfill-auth-state.bin").exists():
            return "저장된 로그인 세션 없음"
        seller_path = profile / "seller-name.txt"
        seller = seller_path.read_text(encoding="utf-8").strip() if seller_path.exists() else ""
        return (
            f"저장된 로그인 정보: {seller} (작업 시 권한 확인)"
            if seller else "저장된 로그인 정보 있음 (작업 시 권한 확인)"
        )

    def _on_detail_log_copy_clicked(self):
        QApplication.clipboard().setText(self.textEdit_detail_log.toPlainText())
        self._detail_log("[작업 로그] 클립보드에 복사했습니다.")

    def _on_detail_log_save_clicked(self):
        log_dir = APP_ROOT / "logs" / "detail-page"
        log_dir.mkdir(parents=True, exist_ok=True)
        path = log_dir / f"detail-page-{datetime.now():%Y%m%d-%H%M%S}.log"
        path.write_text(self.textEdit_detail_log.toPlainText(), encoding="utf-8")
        self._detail_log(f"[작업 로그] 저장했습니다: {path}")
        # 확인 창 대신 알림 띠의 「로그 열기」·「폴더 열기」(계획 C-3).
        self.notify(
            "ok", f"작업 로그를 저장했습니다: {path.name}",
            [("로그 열기", lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))),
             ("폴더 열기", lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(log_dir))))],
            detail=str(path),
        )

    @staticmethod
    def _detail_id_from_input(value, kind):
        value = value.strip()
        if value.isdigit():
            return value
        if kind == "naver":
            match = re.fullmatch(
                r"https?://(?:www\.)?smartstore\.naver\.com/[^/?#]+/products/(\d+)(?:[/?#].*)?",
                value,
                re.IGNORECASE,
            )
        else:
            valid_url = re.fullmatch(
                r"https?://wing\.coupang\.com/tenants/seller-web/vendor-inventory/modify\?[^#]+",
                value,
                re.IGNORECASE,
            )
            match = re.search(r"(?:[?&])vendorInventoryId=(\d+)(?:[&#]|$)", value, re.IGNORECASE) if valid_url else None
        return match.group(1) if match else None

    def _detail_product_no(self, line_edit, label, kind):
        value = line_edit.text().strip()
        product_no = self._detail_id_from_input(value, kind)
        if not product_no:
            QMessageBox.warning(self, "상세페이지", f"{label}를 확인하세요.")
            return None
        line_edit.setText(product_no)
        return product_no

    def _normalize_detail_id_input(self, line_edit, value, kind):
        parsed = self._detail_id_from_input(value, kind)
        if parsed and parsed != value:
            line_edit.setText(parsed)

    def _detail_output_dir(self, product_no):
        return APP_ROOT / "output" / "detail-preview" / product_no

    def _detail_create_progress_dialog(self):
        if self.detail_create_progress is None:
            progress = QProgressDialog(self)
            progress.setWindowTitle("HTML 생성")
            progress.setWindowModality(Qt.WindowModality.NonModal)
            progress.setAutoClose(False)
            progress.setAutoReset(False)
            progress.setCancelButton(None)
            self.detail_create_progress = progress
        return self.detail_create_progress

    def _start_detail_process(self, title, arguments, hold_wing=False):
        if self._detail_process is not None and self._detail_process.state() != QProcess.ProcessState.NotRunning:
            QMessageBox.warning(self, "상세페이지", "다른 상세페이지 작업이 진행 중입니다.")
            return
        process = QProcess(self)
        process.setWorkingDirectory(str(APP_ROOT))
        process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        environment = QProcessEnvironment.systemEnvironment()
        environment.insert("PYTHONIOENCODING", "utf-8")
        environment.insert("PYTHONUNBUFFERED", "1")
        process.setProcessEnvironment(environment)
        process.readyReadStandardOutput.connect(self._on_detail_process_output)
        process.finished.connect(self._on_detail_process_finished)
        self._detail_process = process
        self._detail_process_holding_wing = hold_wing
        self._detail_process_title = title
        self._detail_process_output = ""
        if title == "상품 상세 HTML 생성":
            progress = self._detail_create_progress_dialog()
            progress.setRange(0, 0)
            progress.setLabelText("HTML 생성 준비 중...")
            progress.show()
        self._detail_log(f"[{title}] 시작")
        process.start(sys.executable, arguments)

    def _on_detail_process_output(self):
        if self._detail_process is None:
            return
        text = bytes(self._detail_process.readAllStandardOutput()).decode("utf-8", errors="replace")
        self._detail_process_output += text
        self._detail_log(text)
        if self._detail_process_title == "기존 쿠팡 상품 HTML 채우기" and "준비 완료: 쿠팡 상품" in text:
            self._set_detail_step_done(3)
        if self._detail_process_title == "상품 상세 HTML 생성":
            progress = self._detail_upload_progress(self._detail_process_output)
            if progress:
                current, total = progress
                dialog = self._detail_create_progress_dialog()
                dialog.setRange(0, total)
                dialog.setValue(current)
                dialog.setLabelText(
                    f"쿠팡 이미지 업로드 {current}/{total} ({current * 100 // total}%)"
                )

    def _on_detail_process_finished(self, exit_code, _status):
        self._on_detail_process_output()
        output = self._detail_process_output
        title = self._detail_process_title
        wing_closed = (
            title == "기존 쿠팡 상품 HTML 채우기"
            and "[WING 창 닫힘 감지]" in output
        )
        self._detail_process = None
        self._detail_process_holding_wing = False
        if wing_closed:
            self._on_detail_finish_clicked()
            self._detail_log(
                "[기존 쿠팡 상품 HTML 채우기] WING 창 닫힘을 감지해 WING 작업 종료와 입력 초기화를 처리했습니다."
            )
        else:
            self.pushButton_detail_finish.setEnabled(False)
        if title == "상품 상세 HTML 생성":
            progress = self._detail_create_progress_dialog()
            progress.setRange(0, 100)
            progress.setValue(100 if exit_code == 0 else 0)
            progress.setLabelText(
                "HTML 생성 완료 (100%)" if exit_code == 0 else "HTML 생성 실패 — 작업 로그를 확인하세요."
            )
            if exit_code == 0:
                QTimer.singleShot(2_000, lambda: progress.hide() if self._detail_process is None else None)
        if title == "쿠팡 로그인 연결" and exit_code == 0:
            try:
                result = json.loads(output.splitlines()[-1])
                seller = (result.get("seller") or "").strip()
                if seller:
                    profile = APP_ROOT / "output" / "coupang-browser-profile"
                    profile.mkdir(parents=True, exist_ok=True)
                    (profile / "seller-name.txt").write_text(seller, encoding="utf-8")
                self.label_detail_session.setText(self._detail_session_label())
            except (json.JSONDecodeError, IndexError):
                self.label_detail_session.setText(self._detail_session_label())
            self._set_detail_step_done(1)
        elif title == "쿠팡 로그인 세션 삭제" and exit_code == 0:
            self.label_detail_session.setText("저장된 로그인 세션 없음")
            self._set_detail_step_done(1, False)
        self._detail_log(f"[{title}] {'완료' if exit_code == 0 else f'실패 ({exit_code})'}", error=exit_code != 0)
        if title == "상품 상세 HTML 생성" and exit_code == 0 and self._detail_created_product_no:
            self.lineEdit_detail_naver.setText(self._detail_created_product_no)
            self._set_detail_step_done(2)
            self._open_detail_editor(self._detail_created_product_no)

    def _on_detail_login_clicked(self):
        self._start_detail_process("쿠팡 로그인 연결", ["-m", "product_detail.coupang_session", "--login"])

    def _on_detail_clear_session_clicked(self):
        if QMessageBox.question(self, "로그인 세션 삭제", "이 PC에 저장된 쿠팡 WING 로그인 세션을 삭제할까요?") != QMessageBox.StandardButton.Yes:
            return
        self.label_detail_session.setText("저장된 로그인 세션 삭제 중")
        self._start_detail_process("쿠팡 로그인 세션 삭제", ["-m", "product_detail.coupang_session", "--clear"])

    def _on_detail_create_clicked(self):
        product_no = self._detail_product_no(self.lineEdit_detail_naver, "네이버 상품번호", "naver")
        if product_no:
            existing_html = self._detail_output_dir(product_no) / "coupang-paste.html"
            if existing_html.exists():
                message = QMessageBox(self)
                message.setWindowTitle("HTML 다시 생성")
                message.setIcon(QMessageBox.Icon.Warning)
                message.setText("이미 생성된 HTML이 있습니다.")
                message.setInformativeText(
                    "다시 생성하면 기존 HTML 편집 내용이 초기화될 수 있습니다.\n다시 생성할까요?"
                )
                regenerate = message.addButton("다시 생성", QMessageBox.ButtonRole.DestructiveRole)
                cancel = message.addButton("취소", QMessageBox.ButtonRole.RejectRole)
                message.setDefaultButton(cancel)
                message.setEscapeButton(cancel)
                message.exec()
                if message.clickedButton() is not regenerate:
                    return
            self._detail_created_product_no = product_no
            self._start_detail_process("상품 상세 HTML 생성", ["-m", "product_detail.naver_to_coupang_html", product_no, "--upload"])

    def _open_detail_file(self, product_no, name):
        path = self._detail_output_dir(product_no) / name
        if not path.exists():
            QMessageBox.warning(self, "상세페이지", f"파일이 없습니다. 먼저 HTML을 생성하세요.\n{path}")
            return
        if not hasattr(self, "_detail_preview_servers"):
            self._detail_preview_servers = {}
        key = path.resolve()
        server = self._detail_preview_servers.get(key)
        if server is None or server.closed:
            try:
                server = _host(self).DetailPreviewServer(path)
            except OSError as error:
                QMessageBox.warning(self, "상세페이지", f"미리보기를 시작하지 못했습니다.\n{error}")
                return
            self._detail_preview_servers[key] = server
            QApplication.instance().aboutToQuit.connect(server.close)
            self.destroyed.connect(server.close)
        QDesktopServices.openUrl(QUrl(server.url))

    def _on_detail_open_preview_clicked(self):
        product_no = self._detail_product_no(self.lineEdit_detail_naver, "네이버 상품번호", "naver")
        if product_no:
            self._open_detail_file(product_no, "coupang-paste.html")

    def _on_detail_open_html_clicked(self):
        product_no = self._detail_product_no(self.lineEdit_detail_naver, "네이버 상품번호", "naver")
        if product_no:
            self._open_detail_editor(product_no)

    def _open_detail_editor(self, product_no):
        path = self._detail_output_dir(product_no) / "coupang-paste.html"
        if not path.exists():
            QMessageBox.warning(self, "상세페이지", f"파일이 없습니다. 먼저 HTML을 생성하세요.\n{path}")
            return
        _host(self).DetailHtmlEditorDialog(path, self).exec()

    def _on_detail_open_folder_clicked(self):
        product_no = self._detail_product_no(self.lineEdit_detail_naver, "네이버 상품번호", "naver")
        if product_no:
            path = self._detail_output_dir(product_no)
            if not path.exists():
                QMessageBox.warning(self, "상세페이지", "결과 폴더가 없습니다. 먼저 HTML을 생성하세요.")
                return
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _on_detail_stage_clicked(self):
        product_no = self._detail_product_no(self.lineEdit_detail_naver, "네이버 상품번호", "naver")
        vendor_id = self._detail_product_no(self.lineEdit_detail_vendor_inventory, "쿠팡 등록상품 ID", "coupang")
        if product_no and vendor_id:
            html_path = self._detail_output_dir(product_no) / "coupang-paste.html"
            if not html_path.exists():
                QMessageBox.warning(
                    self,
                    "상세페이지",
                    "저장된 HTML이 없습니다.\n「② 상품 상세 HTML 만들기」에서 HTML을 생성한 뒤 편집·저장해 주세요.",
                )
                return
            self._detail_log(f"[쿠팡 상품 HTML 적용] 저장된 HTML 사용: {html_path}")
            self.pushButton_detail_finish.setEnabled(True)
            self._start_detail_process("기존 쿠팡 상품 HTML 채우기", ["-m", "product_detail.coupang_detail_replace", product_no, vendor_id], hold_wing=True)

    def _on_detail_finish_clicked(self):
        if self._detail_process is not None and self._detail_process_holding_wing:
            self._detail_process.write(b"\n")
            self._detail_log("[기존 쿠팡 상품 HTML 채우기] 저장 없이 WING 작업을 종료합니다.")
        self.pushButton_detail_finish.setEnabled(False)
        self._on_detail_clear_inputs_clicked()
