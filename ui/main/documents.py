"""메인 창 「문서 만들기」 탭: 견적서·거래명세서 화면(documents/quotation_statement_widget.py)을 탭에 붙인다."""

from PySide6.QtWidgets import QFrame, QLabel, QScrollArea

from shared.paths import APP_ROOT
from ui.main import host as _host


class DocumentsMixin:
    """MainWindow 믹스인. 메인 창 「문서 만들기」 탭: 견적서·거래명세서 화면(documents/quotation_statement_widget.py)을 탭에 붙인다."""

    def _mount_document_widget(self):
        """문서 탭 첫 진입 때 위젯을 장착하고 실패하면 탭 안에 오류를 표시한다."""
        if self._document_widget_mount_attempted:
            return
        layout = getattr(self.ui, "verticalLayout_documents", None)
        if layout is None:
            return
        self._document_widget_mount_attempted = True
        try:
            from documents.quotation_statement_widget import QuotationStatementWidget

            self._document_widget = QuotationStatementWidget(
                product_search_worker=_host(self).run_naver_product_search_worker,
                order_lookup_worker=_host(self).run_naver_order_lookup_worker,
                coupang_order_lookup_worker=_host(self).run_coupang_order_lookup_worker,
                on_created=self.show_excel_created_message,
                base_dir=APP_ROOT,
                parent=self,
            )
            # 문서 화면의 최소 높이(약 770px)가 창보다 크면 첫 진입 때 창이 커지거나
            # 표와 「계산 요약」이 겹쳐 그려진다. 스크롤 영역에 넣어 창 크기를 유지한다.
            scroll = QScrollArea()
            scroll.setObjectName("scrollArea_documents")
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.Shape.NoFrame)
            scroll.setWidget(self._document_widget)
            # 스크롤 영역이 창 배경색(회색)을 칠해 다른 탭(흰색)과 달라 보이지 않게 한다(설정 탭과 같은 방식).
            scroll.viewport().setAutoFillBackground(False)
            self._document_widget.setAutoFillBackground(False)
            layout.setContentsMargins(6, 6, 6, 6)
            layout.addWidget(scroll)
        except Exception as exc:
            self._document_widget = None
            error = QLabel(
                "문서 만들기 기능을 불러오지 못했습니다. 기존 기능은 계속 사용할 수 있습니다.\n"
                f"오류: {exc}"
            )
            error.setWordWrap(True)
            layout.addWidget(error)
            print(f"! 문서 만들기 탭 초기화 실패: {exc}")
