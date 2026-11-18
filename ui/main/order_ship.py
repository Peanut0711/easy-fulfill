"""메인 창 「주문·발송」 탭 화면: ①② 패널 분할 비율, 주문 미리보기 표, 패널별 버튼 줄과 다음 동작 강조, 끌어다 놓기, 진행 요약, 파일 상태 라벨, 스토어 로고."""

from pathlib import Path

from PySide6.QtCore import QEvent, QSize, Qt
from PySide6.QtWidgets import QApplication, QFrame, QMessageBox, QVBoxLayout, QWidget

from ui.order_preview import (
    CODE_MISSING_BACKGROUND,
    NON_PARCEL_BACKGROUND,
    setup_preview_view,
    show_preview,
    summary_counts,
)
from ui.order_ship_panel import (
    apply_primary,
    FileDropTarget,
    next_stage,
    order_summary_html,
    ship_summary_html,
    stabilize_widths,
    STAGE_INTAKE,
    STAGE_INVOICE,
    STAGE_ORDER,
    STAGE_RECEIPT,
    store_badge_pixmap,
    STORE_NAMES,
)

from shared import theme
from shared.ui_components import fixed_width_button, make_step_badge, set_props, set_step_badge_done


class OrderShipMixin:
    """MainWindow 믹스인. 메인 창 「주문·발송」 탭 화면: ①② 패널 분할 비율, 주문 미리보기 표, 패널별 버튼 줄과 다음 동작 강조, 끌어다 놓기, 진행 요약, 파일 상태 라벨, 스토어 로고."""

    ORDER_SHIP_BUTTON_NAMES = (
        "pushButton_load_order", "pushButton_copy_order", "pushButton_intake_excel", "pushButton_clear_order",
        "pushButton_load_invoice", "pushButton_courier_receipt",
    )

    def _on_first_show_splitter(self):
        # 창을 만들 때는 앱 스타일시트가 아직 다 적용되지 않아 버튼 폭이 작게 잡힌다. 처음 보일 때 다시 잰다.
        stabilize_widths(getattr(self.ui, name, None) for name in self.ORDER_SHIP_BUTTON_NAMES)
        self._apply_order_ship_splitter_sizes()
        self._order_ship_splitter_initialized = True

    def _apply_order_ship_splitter_sizes(self):
        """주문·발송 탭 스플리터 초기 비율 (레이아웃 적용 후 한 프레임 뒤에 호출)."""
        splitter = getattr(self.ui, "splitter_order_ship", None)
        if splitter is None:
            return

        def _stretch_last_column(panel: QWidget):
            lay = panel.layout()
            if isinstance(lay, QVBoxLayout) and lay.count() > 0:
                last = lay.count() - 1
                for i in range(lay.count()):
                    lay.setStretch(i, 1 if i == last else 0)

        op = getattr(self.ui, "widget_order_panel", None)
        spanel = getattr(self.ui, "widget_ship_panel", None)
        if op:
            _stretch_last_column(op)
        if spanel:
            _stretch_last_column(spanel)

        w = splitter.width()
        if w < 200:
            w = max(self.width() - 48, 1200)
        left = int(w * 0.52)
        splitter.setSizes([left, max(w - left, 200)])

    def _setup_order_preview(self):
        """주문 미리보기 표(펼치면 상품 목록)를 준비한다."""
        view = getattr(self.ui, "treeView_order_preview", None)
        if view is None:
            return
        setup_preview_view(view, copy_callback=self.copy_to_clipboard)
        legend = getattr(self.ui, "label_order_legend", None)
        if legend is not None:
            legend.setVisible(False)

    def _show_order_preview(self, markdown_text, rows):
        """주문 미리보기를 갱신한다. markdown_text는 복사(Ctrl+C)에 그대로 쓰고, rows는 표에 보여 준다."""
        self._order_markdown_text = markdown_text or ""
        self._order_preview_rows = list(rows or [])
        view = getattr(self.ui, "treeView_order_preview", None)
        if view is not None:
            show_preview(view, self._order_preview_rows)
        stack = getattr(self.ui, "stackedWidget_order_preview", None)
        if stack is not None:
            stack.setCurrentIndex(1 if self._order_preview_rows else 0)
        legend = getattr(self.ui, "label_order_legend", None)
        if legend is not None:
            counts = summary_counts(self._order_preview_rows)
            parts = []
            if counts["code_missing"]:
                parts.append(
                    f'<span style="background:{CODE_MISSING_BACKGROUND};">&nbsp;&nbsp;&nbsp;&nbsp;</span>'
                    f' 상품코드 매칭 실패 {counts["code_missing"]}건')
            if counts["non_parcel"]:
                parts.append(
                    f'<span style="background:{NON_PARCEL_BACKGROUND};">&nbsp;&nbsp;&nbsp;&nbsp;</span>'
                    f' 택배 외 배송 {counts["non_parcel"]}건')
            legend.setText("&nbsp;&nbsp;&nbsp;".join(parts))
            legend.setVisible(bool(parts))

    def _apply_order_ship_design(self):
        """주문·발송 탭을 디자인 가이드 모양으로 짠다(계획 4-1).

        - ①② 패널을 카드로, 단계 문자(①②)를 완료가 보이는 단계 배지로 바꾼다.
        - 「초기화」는 작업 순서 버튼 줄에서 빼 ① 카드 제목 줄 오른쪽에 둔다(버튼 줄이 짧아져
          좁은 창에서도 겹치지 않는다).
        - 버튼은 글자 폭보다 늘리지 않고, 빈 미리보기 안내는 빈 상태 글자로 보인다.
        """
        ui = self.ui
        tab_layout = getattr(ui, "verticalLayout_tab_order_ship", None)
        if tab_layout is not None:
            tab_layout.setContentsMargins(theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
        splitter = getattr(ui, "splitter_order_ship", None)
        if splitter is not None:
            splitter.setHandleWidth(theme.SPACE_MD)
        for name in ("widget_order_panel", "widget_ship_panel"):
            panel = getattr(ui, name, None)
            if panel is None:
                continue
            panel.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            set_props(panel, card=True)
            panel.setMinimumWidth(0)  # .ui의 320 고정 대신 레이아웃이 계산한 최소 폭(버튼 줄 등)을 따른다
            if panel.layout() is not None:
                panel.layout().setContentsMargins(
                    theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD, theme.SPACE_MD)
                panel.layout().setSpacing(theme.SPACE_SM)
        self._order_step_badges = {}
        for step, label_name, layout_name, title in (
                (1, "label_step_order", "horizontalLayout_order_title", "주문 불러오기"),
                (2, "label_step_ship", "horizontalLayout_ship_title", "송장·발송 생성")):
            label, layout = getattr(ui, label_name, None), getattr(ui, layout_name, None)
            if label is None or layout is None:
                continue
            label.setText(title)
            set_props(label, sectionTitle=True)
            badge = make_step_badge(step)
            layout.insertWidget(0, badge)
            layout.setSpacing(theme.SPACE_SM)
            self._order_step_badges[step] = badge
        clear = getattr(ui, "pushButton_clear_order", None)
        buttons_row = getattr(ui, "horizontalLayout_order_buttons", None)
        title_row = getattr(ui, "horizontalLayout_order_title", None)
        if clear is not None and buttons_row is not None and title_row is not None:
            buttons_row.removeWidget(clear)
            title_row.addWidget(clear)
        for row_name in ("horizontalLayout_order_buttons", "horizontalLayout_ship_buttons"):
            row = getattr(ui, row_name, None)
            if row is None:
                continue
            row.setSpacing(theme.SPACE_SM)
            for i in range(row.count()):
                widget = row.itemAt(i).widget()
                if widget is not None:
                    fixed_width_button(widget)
            row.addStretch(1)
        if clear is not None:
            fixed_width_button(clear)
        for name in ("filePathLabel", "label_invoice"):  # 파일 칸은 긴 이름을 말줄임으로 보이므로 좁아져도 된다
            drop = getattr(ui, name, None)
            if drop is not None:
                drop.setMinimumWidth(120)
        invoice_caption = getattr(ui, "label", None)  # 「송장 정보」: 파일 칸 하나뿐이라 라벨 없이 둔다
        if invoice_caption is not None:
            invoice_caption.hide()
        empty = getattr(ui, "label_order_empty", None)
        if empty is not None:
            set_props(empty, emptyText=True)
        stack = getattr(ui, "stackedWidget_order_preview", None)
        if stack is not None:
            stack.setFrameShape(QFrame.Shape.NoFrame)
        empty_page = getattr(ui, "page_order_empty", None)
        if empty_page is not None:  # 비어 있을 때는 끌어다 놓을 수 있는 점선 영역으로 보인다
            empty_page.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            set_props(empty_page, emptyPanel=True)
        legend = getattr(ui, "label_order_legend", None)
        if legend is not None:
            set_props(legend, hint="true")

    def _setup_order_ship_buttons(self):
        """주문·발송 탭 패널별 버튼 줄: 아이콘·툴팁을 메뉴 동작과 맞추고 연결한다."""
        self._apply_order_ship_design()
        specs = (
            ("pushButton_load_order", self.act_load, "주문 엑셀 불러오기 (Ctrl+O)", None),
            ("pushButton_copy_order", self.act_copy, "주문 정보를 마크다운으로 복사 (Ctrl+C)", self.act_copy.trigger),
            ("pushButton_intake_excel", self.act_export,
             "우체국 접수용 하이제니스 폼 엑셀 만들기 — 메뉴 「송장 출력」과 같음 (Ctrl+E)", self.act_export.trigger),
            ("pushButton_clear_order", self.act_clear, "주문·송장 정보 초기화 (Ctrl+R)", self.act_clear.trigger),
            ("pushButton_load_invoice", self.act_load_invoice,
             "우체국 송장 엑셀 불러오기 — 일괄발송 파일을 자동으로 만듭니다", None),
            ("pushButton_courier_receipt", self.act_courier_export,
             "오늘 등록된 송장으로 택배사 접수목록 엑셀 내보내기", self.act_courier_export.trigger),
        )
        for name, action, tip, handler in specs:
            button = getattr(self.ui, name, None)
            if button is None:
                continue
            button.setIcon(action.icon())
            button.setIconSize(QSize(16, 16))
            button.setToolTip(tip)
            if handler is not None:
                button.clicked.connect(lambda _checked=False, h=handler: h())
        stabilize_widths(getattr(self.ui, name, None) for name, *_ in specs)

    def _setup_order_ship_drop_targets(self):
        """주문·발송 패널에 엑셀 끌어다 놓기를 붙인다(파일 대화상자 없이 같은 검증으로 불러오기)."""
        order_panel = getattr(self.ui, "widget_order_panel", None)
        ship_panel = getattr(self.ui, "widget_ship_panel", None)
        if order_panel is not None:
            self._order_drop_target = FileDropTarget(
                order_panel, self._on_order_file_dropped, "여기에 놓으면 주문 엑셀을 불러옵니다")
        if ship_panel is not None:
            self._invoice_drop_target = FileDropTarget(
                ship_panel, self._on_invoice_file_dropped, "여기에 놓으면 우체국 송장 엑셀을 불러옵니다")
        # 송장 미리보기 글 상자가 파일 경로를 글자로 받아 넣지 않게 하고, 놓기는 패널이 받는다.
        if hasattr(self.ui, "plainTextEdit_invoice"):
            self.ui.plainTextEdit_invoice.setAcceptDrops(False)

    def _order_ship_busy(self):
        overlay = getattr(self, "_busy_overlay", None)
        return overlay is not None and overlay.isVisible()

    def _on_order_file_dropped(self, file_path):
        if self._order_ship_busy():
            self.statusBar().showMessage("처리 중입니다. 끝난 뒤 다시 놓아 주세요.", 3000)
            return
        print(f"\n[주문 엑셀 끌어다 놓기] {file_path}")
        self._load_order_file(file_path)

    def _on_invoice_file_dropped(self, file_path):
        if self._order_ship_busy():
            self.statusBar().showMessage("처리 중입니다. 끝난 뒤 다시 놓아 주세요.", 3000)
            return
        print(f"\n[송장 엑셀 끌어다 놓기] {file_path}")
        self._load_invoice_path(file_path)

    def _refresh_order_ship_panels(self):
        """주문·발송 탭: 다음에 할 동작 하나만 강조하고, 쓸 수 없는 버튼은 끈다."""
        if not hasattr(self, "ui"):
            return
        has_orders = bool(self.orders) and bool(self.is_order_file_valid)
        stage = next_stage(
            has_orders=has_orders,
            intake_done=self._intake_excel_done,
            batch_done=self._ship_batch_done,
        )
        self._order_ship_stage = stage
        badges = getattr(self, "_order_step_badges", {})
        if 1 in badges:
            set_step_badge_done(badges[1], has_orders)
        if 2 in badges:
            set_step_badge_done(badges[2], bool(self._ship_batch_done))
        apply_primary({
            STAGE_ORDER: getattr(self.ui, "pushButton_load_order", None),
            STAGE_INTAKE: getattr(self.ui, "pushButton_intake_excel", None),
            STAGE_INVOICE: getattr(self.ui, "pushButton_load_invoice", None),
            STAGE_RECEIPT: getattr(self.ui, "pushButton_courier_receipt", None),
        }, stage)
        if hasattr(self.ui, "pushButton_copy_order"):
            self.ui.pushButton_copy_order.setEnabled(bool(self._order_markdown_text))
        if hasattr(self.ui, "pushButton_intake_excel"):
            self.ui.pushButton_intake_excel.setEnabled(bool(self.orders))
        if hasattr(self.ui, "label_order_summary"):
            counts = summary_counts(self._order_preview_rows) if has_orders else None
            self.ui.label_order_summary.setText(order_summary_html(self.store_type, counts))
        if hasattr(self.ui, "label_ship_summary"):
            self.ui.label_ship_summary.setText(ship_summary_html(
                batch=self._ship_batch_summary if self._ship_batch_done else None,
                invoice_rows=self._invoice_row_count))
            batch = self._ship_batch_summary if self._ship_batch_done else None
            if batch and batch[3] is not None:
                self.ui.label_ship_summary.setToolTip(
                    f"송장 {batch[3]}건 중 {batch[2]}건을 주문과 매칭해 일괄발송 파일을 만들었습니다.")
            else:
                self.ui.label_ship_summary.setToolTip("")

    def _mark_ship_batch_done(self, store_type, order_count):
        """일괄발송 파일을 만든 묶음을 ② 요약에 남긴다(① 패널은 생성 뒤 비워진다)."""
        matched, total = self._last_invoice_match or (None, None)
        self._ship_batch_summary = (store_type, order_count, matched, total)
        self._ship_batch_done = True
        self._refresh_order_ship_panels()

    def copy_to_clipboard(self):
        """불러온 주문의 마크다운(불러올 때 자동 복사한 것과 같은 글)을 클립보드에 복사합니다."""
        try:
            text = self._order_markdown_text
            if not text:
                self.statusBar().showMessage("복사할 주문 정보가 없습니다.", 2000)
                return
            QApplication.clipboard().setText(text)
            self.statusBar().showMessage("주문 정보가 클립보드에 복사되었습니다.", 2000)
        except Exception as e:
            QMessageBox.warning(self, "오류", f"복사 중 오류가 발생했습니다: {str(e)}")

    def _set_status_label(self, label, text, ok=False):
        """주문·발송 탭 상태 라벨 텍스트+색. ok=True 초록(완료/불러옴), 기본 회색(대기/없음).
        가운데 정렬 + 폭에 맞춰 긴 파일명은 끝을 '…'로 줄인다(전체는 툴팁)."""
        if label is None:
            return
        label.setProperty("_full_text", text)
        label.setAlignment(Qt.AlignCenter)
        # 파일 상태 칸은 끌어다 놓을 수 있는 곳임을 점선 테두리로 보여 준다(모양은 테마의 dropZone 규칙).
        state = "ok" if ok else "idle"
        if label.property("fileState") != state or not label.property("dropZone"):
            label.setProperty("dropZone", True)
            label.setProperty("fileState", state)
            style = label.style()
            style.unpolish(label)
            style.polish(label)
        # 리사이즈 시 다시 줄이도록 이벤트 필터 1회 등록
        if not hasattr(self, "_elide_labels"):
            self._elide_labels = set()
        if label not in self._elide_labels:
            self._elide_labels.add(label)
            label.installEventFilter(self)
        self._elide_status_label(label)

    def _elide_status_label(self, label):
        """라벨의 현재 폭에 맞춰 보관된 전체 텍스트를 끝 생략(…)으로 표시."""
        full = label.property("_full_text")
        if full is None:
            return
        elided = label.fontMetrics().elidedText(
            full, Qt.TextElideMode.ElideRight, max(0, label.width() - 8))
        label.setText(elided)
        label.setToolTip(full if elided != full else "")

    def eventFilter(self, obj, event):
        """상태 라벨 리사이즈 시 끝 생략(…) 재계산."""
        if event.type() == QEvent.Type.Resize and obj in getattr(self, "_elide_labels", ()):
            self._elide_status_label(obj)
        return super().eventFilter(obj, event)

    _STORE_NAMES = STORE_NAMES

    def _set_store_label(self, store_type):
        """주문 스토어를 image/ 폴더의 로고 배지로 표시. 미선택이면 지원 스토어 로고를 흐리게 보여 준다."""
        lbl = getattr(self.ui, "label_logo", None)
        if lbl is None:
            return
        name = self._STORE_NAMES.get(store_type or "", "")
        lbl.clear()  # 이전 픽스맵 제거
        lbl.setFrameShape(QFrame.Shape.NoFrame)  # 버튼처럼 보이던 패널 테두리 제거(라벨로)
        lbl.setAlignment(Qt.AlignCenter)
        try:
            pixmap = store_badge_pixmap(
                store_type if name else None, Path("image"), (140, 40), lbl.devicePixelRatioF())
        except Exception as e:  # 로고를 그리지 못해도 글자로는 보여 준다.
            print(f"! 스토어 로고 배지 표시 실패: {e}")
            pixmap = None
        if pixmap is not None:
            lbl.setPixmap(pixmap)
        else:
            lbl.setText(name or "스토어 미선택")
        lbl.setToolTip(f"{name} 주문" if name else "스토어 미선택 — 주문 엑셀을 불러오면 해당 스토어가 표시됩니다")
        lbl.setAccessibleName(name or "스토어 미선택")

    def _reset_idle_logo(self):
        """주문 미로드 시 스토어 라벨을 옅은 기본값으로 되돌림."""
        self._set_store_label(None)
