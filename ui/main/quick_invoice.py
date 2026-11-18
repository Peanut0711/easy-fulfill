"""메인 창 「단건 송장」 탭: 스토어 선택·수동 입력 격자, 우체국 우편번호·주소 검색, 클립보드 붙여 넣기 해석, 송장 엑셀 만들기, 오늘 만든 단건 송장 목록."""

import json
import os
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

import xml.etree.ElementTree as ET
from orders.quick import (
    detect_quick_store,
    extract_zip_code,
    parse_coupang_quick_clipboard,
    parse_gmarket_quick_clipboard,
    parse_naver_quick_clipboard,
    save_invoice_excel,
)
from shared.ui_style import set_button_role
from tracking.workers import CONFIG_KEY_KPOST_REGKEY
from ui.main import host as _host


class QuickInvoiceMixin:
    """MainWindow 믹스인. 메인 창 「단건 송장」 탭: 스토어 선택·수동 입력 격자, 우체국 우편번호·주소 검색, 클립보드 붙여 넣기 해석, 송장 엑셀 만들기, 오늘 만든 단건 송장 목록."""

    def _iter_quick_excel_manual_widgets(self):
        for name in (
            "label_quick_manual_name",
            "lineEdit_quick_manual_name",
            "label_quick_manual_phone",
            "lineEdit_quick_manual_phone",
            "label_quick_manual_address",
            "lineEdit_quick_manual_address",
            "label_quick_manual_postcode",
            "lineEdit_quick_manual_postcode",
            "pushButton_quick_manual_postcode_lookup",
            "pushButton_quick_manual_address_search",
            "label_quick_manual_detail",
            "lineEdit_quick_manual_detail",
            "label_quick_manual_product",
            "lineEdit_quick_manual_product",
        ):
            if hasattr(self.ui, name):
                yield getattr(self.ui, name)

    def _get_kpost_regkey(self) -> str:
        """종추적과 기존 겸용 키에 쓰이는 우체국 regkey."""
        for env_name in ("KPOST_REGKEY", "EPOST_REGKEY"):
            v = os.environ.get(env_name, "").strip()
            if v:
                return v
        try:
            if self.app_settings_path.exists() and self.app_settings_path.stat().st_size > 0:
                with open(self.app_settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    raw = data.get("kpost_regkey") or data.get("epost_regkey")
                    if isinstance(raw, str) and raw.strip():
                        return raw.strip()
        except Exception as e:
            print(f"! kpost regkey 로드 오류: {e}")
        return ""

    def _show_kpost_api_error(self, parent_widget, err_pair):
        code, msg = err_pair
        code = (code or "—").strip()
        msg = (msg or "사유를 확인할 수 없습니다.").strip()
        parent = parent_widget if parent_widget is not None else self

        if code == _host(self).KPOST_ERROR_NO_REGKEY:
            QMessageBox.warning(
                parent,
                "우편번호 API 키가 없습니다",
                "「⚙ 설정 › 연동」의 API 키에서 「우체국 주소 검색」을 등록해 주세요.\n\n"
                "주소와 우편번호를 직접 입력해 엑셀을 만들 수도 있습니다.",
            )
            return

        if code == _host(self).KPOST_ERROR_CONFIG_READ:
            QMessageBox.warning(
                parent, "공유 키 확인 실패",
                "공유 시트에서 우편번호 API 키를 읽지 못했습니다.\n"
                "「⚙ 설정 › 연동」에서 Google Sheets 연결과 시트 접근 권한을 확인해 주세요.",
            )
            return

        if code == "ERR-123":
            msg = (
                "우체국에서 우편번호 조회용 인증키를 확인하지 못했습니다.\n"
                "우편번호 조회에 등록된 키를 「⚙ 설정 › 연동」의 API 키 「우체국 주소 검색」에 입력해 주세요."
            )

        QMessageBox.warning(
            parent,
            "우편번호 조회 오류",
            f"오류 코드: {code}\n\n원인:\n{msg}",
        )

    def _kpost_parse_postnew_response(self, xml: bytes):
        """
        Returns:
            (items, None) — items: [{postcd, address, addrjibun}, ...]
            (None, (error_code, message)) — API 오류 XML
        """
        try:
            root = ET.fromstring(xml)
        except ET.ParseError:
            return None, ("XML 오류", "응답을 XML로 해석할 수 없습니다.")

        for el in root.iter():
            if _host(self)._xml_local_lower(el.tag) != "error":
                continue
            code, msg = "", ""
            for ch in el:
                ln = _host(self)._xml_local_lower(ch.tag)
                if ln == "error_code":
                    code = (ch.text or "").strip()
                elif ln == "message":
                    msg = (ch.text or "").strip()
            return None, (code or "오류", msg or "알 수 없는 오류입니다.")

        items = []
        for el in root.iter():
            if _host(self)._xml_local_lower(el.tag) != "item":
                continue
            row = {}
            for ch in el:
                ln = _host(self)._xml_local_lower(ch.tag)
                if ln in ("postcd", "address", "addrjibun"):
                    row[ln] = (ch.text or "").strip()
            if row.get("postcd") or row.get("address"):
                items.append(
                    {
                        "postcd": row.get("postcd", ""),
                        "address": row.get("address", ""),
                        "addrjibun": row.get("addrjibun", ""),
                    }
                )

        return items, None

    def _kpost_postnew_lookup_items(self, query: str):
        """postNew 통합검색. (items, None) 또는 (None, (code, msg))."""
        q = (query or "").strip()
        if len(q) < 2:
            return None, (
                "ERR-121",
                "검색어는 2자 이상 입력해야 합니다.",
            )

        try:
            config = _host(self).run_in_background(_host(self)._read_shared_postcode_config)
        except Exception:
            # Google Sheets 연결 없이 쓰던 PC는 기존 로컬 겸용 키로 검색할 수 있다.
            regkey = self._get_kpost_regkey()
            if not regkey:
                return None, (_host(self).KPOST_ERROR_CONFIG_READ, "")
        else:
            regkey = (config.get(_host(self).CONFIG_KEY_KPOST_POSTCODE_REGKEY)
                      or self._get_kpost_regkey()
                      or config.get(CONFIG_KEY_KPOST_REGKEY, ""))
            if not regkey:
                return None, (_host(self).KPOST_ERROR_NO_REGKEY, "")
        return self._request_kpost_postnew(regkey, q)

    def _request_kpost_postnew(self, regkey: str, query: str):
        params = {
            "regkey": regkey,
            "target": "postNew",
            "query": query,
            "countPerPage": 20,
            "currentPage": 1,
        }
        try:
            # 최대 15초 걸릴 수 있어 작업 스레드에서 기다린다(화면은 계속 그린다).
            resp = _host(self).run_in_background(
                _host(self).requests.get, _host(self).KPOST_OPENAPI2_URL, params=params, timeout=15, allow_redirects=False,
            )
            resp.raise_for_status()
        except _host(self).requests.RequestException as e:
            # 예외 문구에는 regkey가 담긴 요청 URL이 들어 있으므로 원인 종류만 보여 준다.
            status = getattr(getattr(e, "response", None), "status_code", None)
            reason = f"HTTP {status}" if status else type(e).__name__
            return None, ("HTTP 오류", f"우편번호 검색 요청 실패({reason})")

        items, error = self._kpost_parse_postnew_response(resp.content)
        if error:
            code, message = error
            return None, (code, message.replace(regkey, "[인증키]"))
        return items, None

    def _on_quick_manual_postcode_lookup_clicked(self):
        """주소란 검색어로 우편번호만 조회해 첫 결과를 입력합니다."""
        if not hasattr(self.ui, "lineEdit_quick_manual_address"):
            return
        q = self.ui.lineEdit_quick_manual_address.text().strip()
        if len(q) < 2:
            QMessageBox.warning(
                self,
                "입력",
                "기본주소 칸에 도로명·번지 등 검색에 쓸 내용을 2자 이상 입력해 주세요.",
            )
            return
        items, err = self._kpost_postnew_lookup_items(q)
        if err:
            self._show_kpost_api_error(self, err)
            return
        if not items:
            QMessageBox.information(
                self,
                "결과 없음",
                "우편번호를 찾지 못했습니다.\n"
                "「주소 검색」으로 목록에서 고르거나 주소를 더 구체적으로 입력해 보세요.",
            )
            return
        if hasattr(self.ui, "lineEdit_quick_manual_postcode"):
            self.ui.lineEdit_quick_manual_postcode.setText(items[0]["postcd"])
        hint = items[0].get("address", "")
        self.statusBar().showMessage(
            f"우편번호 입력됨 (첫 번째 결과) — {hint[:60]}{'…' if len(hint) > 60 else ''}",
            5000,
        )

    def _on_quick_manual_address_search_clicked(self):
        dlg = _host(self).KpostAddressSearchDialog(self, self)
        if dlg.exec() != QDialog.Accepted:
            return
        if hasattr(self.ui, "lineEdit_quick_manual_address"):
            self.ui.lineEdit_quick_manual_address.setText(dlg.selected_address)
        if hasattr(self.ui, "lineEdit_quick_manual_postcode"):
            self.ui.lineEdit_quick_manual_postcode.setText(dlg.selected_postcd)
        if hasattr(self.ui, "lineEdit_quick_manual_detail"):
            self.ui.lineEdit_quick_manual_detail.setText(dlg.detail_address)

    def _refresh_quick_excel_manual_ui(self):
        """수동 모드일 때만 이름·전화·주소 입력란을 표시합니다.
        그룹 크기는 _mount_quick_excel_inline이 붙인 격자 레이아웃이 맞춘다."""
        if not hasattr(self.ui, "comboBox_store_select"):
            return
        manual = self.ui.comboBox_store_select.currentText().strip() == "수동"
        for w in self._iter_quick_excel_manual_widgets():
            w.setVisible(manual)
        if hasattr(self.ui, "label_gift_3"):
            self.ui.label_gift_3.setVisible(not manual)

    def _mount_quick_excel_inline(self):
        """퀵엑셀 그룹(.ui의 절대좌표 groupBox)을 단건 송장 탭에 상시 인라인으로 삽입."""
        gb = getattr(self.ui, "groupBox", None)
        lay = getattr(self.ui, "verticalLayout_quick_excel", None)
        if gb is None or lay is None:
            return
        self._apply_quick_excel_grid_layout(gb)
        # 입력 묶음 오른쪽에 「오늘 만든 단건 송장」 목록을 둔다(계획 D-2).
        row = QWidget()
        row.setObjectName("widget_quick_invoice_row")
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(12)
        row_layout.addWidget(gb)
        row_layout.addWidget(self._build_quick_today_box(), 1)
        # 하단 세로 스페이서 바로 앞(마지막 위젯 자리)에 넣는다.
        lay.insertWidget(max(lay.count() - 1, 0), row)
        self._refresh_quick_today_list()

    QUICK_TODAY_LIMIT = 50

    def _build_quick_today_box(self):
        box = QGroupBox("오늘 만든 단건 송장")
        box.setObjectName("groupBox_quick_today")
        box.setMinimumWidth(180)
        box_layout = QVBoxLayout(box)
        self.listWidget_quick_today = QListWidget()
        self.listWidget_quick_today.setObjectName("listWidget_quick_today")
        self.listWidget_quick_today.setToolTip("두 번 누르면 엑셀을 엽니다.")
        self.listWidget_quick_today.itemActivated.connect(self._on_quick_today_item_activated)
        # 높이는 왼쪽 입력 묶음에 맞춘다(목록의 기본 높이가 줄 높이를 늘리지 않게).
        self.listWidget_quick_today.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Ignored)
        self.listWidget_quick_today.setMinimumHeight(60)
        box_layout.addWidget(self.listWidget_quick_today, 1)
        folder_button = QPushButton("폴더 열기")
        folder_button.setObjectName("pushButton_quick_today_folder")
        folder_button.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._quick_output_dir().resolve()))))
        box_layout.addWidget(folder_button, 0, Qt.AlignmentFlag.AlignRight)
        return box

    @staticmethod
    def _quick_output_dir():
        return Path("output")  # save_invoice_excel과 같은 폴더(작업 폴더 기준)

    def _quick_excel_recipient(self, path):
        """퀵 엑셀 첫 행의 수취인명. 같은 파일은 다시 읽지 않는다."""
        cache = self.__dict__.setdefault("_quick_recipient_cache", {})
        try:
            key = (str(path), path.stat().st_mtime)
        except OSError:
            return ""
        if key not in cache:
            name = ""
            try:
                from openpyxl import load_workbook

                workbook = load_workbook(path, read_only=True)
                try:
                    rows = workbook.active.iter_rows(min_row=1, max_row=2, values_only=True)
                    header = next(rows, ())
                    first = next(rows, ())
                    if "수취인명" in header:
                        name = str(first[header.index("수취인명")] or "").strip()
                finally:
                    workbook.close()
            except Exception as error:
                print(f"! 단건 송장 목록: 수취인을 읽지 못했습니다 {path.name}: {error}")
            cache[key] = name
        return cache[key]

    def _refresh_quick_today_list(self):
        """오늘 output 폴더에 만든 퀵 엑셀(퀵_<스토어>_<YYYYMMDDHHMMSS>.xlsx)을 최근 것부터 보인다."""
        widget = getattr(self, "listWidget_quick_today", None)
        if widget is None:
            return
        folder = self._quick_output_dir()
        today = datetime.now().strftime("%Y%m%d")
        files = sorted(folder.glob(f"퀵_*_{today}??????.xlsx"), reverse=True) if folder.is_dir() else []
        widget.clear()
        for path in files[: self.QUICK_TODAY_LIMIT]:
            parts = path.stem.split("_")
            stamp = parts[-1]
            store = "_".join(parts[1:-1])
            recipient = self._quick_excel_recipient(path) or "(수취인 없음)"
            item = QListWidgetItem(f"{stamp[8:10]}:{stamp[10:12]}  {recipient}  · {store}")
            item.setToolTip(str(path.resolve()))
            item.setData(Qt.ItemDataRole.UserRole, str(path.resolve()))
            widget.addItem(item)
        if not files:
            item = QListWidgetItem("오늘 만든 단건 송장이 없습니다.")
            item.setFlags(Qt.ItemFlag.NoItemFlags)
            widget.addItem(item)

    def _on_quick_today_item_activated(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        if path:
            self._open_created_file(path)

    QUICK_REQUIRED_FIELDS = (
        ("lineEdit_quick_manual_name", "이름"),
        ("lineEdit_quick_manual_phone", "전화번호"),
        ("lineEdit_quick_manual_address", "기본주소"),
    )

    def _update_quick_excel_gen_enabled(self, *_args):
        """수동 모드에서는 이름·전화번호·기본주소를 다 넣어야 「생성」이 켜진다(빈 칸은 툴팁). 클립보드 모드는 항상 켠다."""
        button = getattr(self.ui, "pushButton_quick_excel_gen", None)
        combo = getattr(self.ui, "comboBox_store_select", None)
        if button is None:
            return
        manual = combo is not None and combo.currentText().strip() == "수동"
        missing = [label for name, label in self.QUICK_REQUIRED_FIELDS
                   if manual and hasattr(self.ui, name) and not getattr(self.ui, name).text().strip()]
        button.setEnabled(not missing)
        if missing:
            button.setToolTip("빈 칸을 채우면 누를 수 있습니다: " + ", ".join(missing))
        elif manual:
            button.setToolTip("입력한 정보로 송장 엑셀을 만듭니다 (Ctrl+Enter)")
        else:
            button.setToolTip("클립보드에 복사한 주문 정보로 송장 엑셀을 만듭니다 (Ctrl+Enter)")

    def _on_quick_excel_shortcut(self):
        button = getattr(self.ui, "pushButton_quick_excel_gen", None)
        if button is not None and button.isEnabled():
            button.click()

    def _apply_quick_excel_grid_layout(self, gb):
        """.ui의 절대좌표 배치를 격자 레이아웃으로 바꾼다.
        절대좌표에서는 버튼이 그룹 테두리 밖으로 나오고, 모드마다 그룹 크기를 손으로 맞춰야 했다."""
        if gb.layout() is not None:
            return
        grid = QGridLayout(gb)
        grid.setContentsMargins(12, 12, 12, 12)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        grid.setColumnStretch(1, 1)

        def w(name):
            return getattr(self.ui, name, None)

        def put(name, row, col, row_span=1, col_span=1):
            widget = w(name)
            if widget is not None:
                widget.setMinimumSize(0, 0)
                widget.setMaximumSize(16777215, 16777215)
                grid.addWidget(widget, row, col, row_span, col_span)

        # 1행: 스토어 선택. 「생성」·「초기화」는 입력을 다 한 뒤 누르도록 맨 아래 줄에 둔다(계획 D-2).
        put("label_store_select", 0, 0)
        put("comboBox_store_select", 0, 1)
        # 2행: 클립보드 안내(수동 모드가 아닐 때만 표시)
        put("label_gift_3", 1, 0, 1, 4)
        if w("label_gift_3") is not None:
            w("label_gift_3").setWordWrap(True)
        # 수동 입력란(수동 모드에서만 표시)
        put("label_quick_manual_name", 2, 0)
        put("lineEdit_quick_manual_name", 2, 1, 1, 3)
        put("label_quick_manual_phone", 3, 0)
        put("lineEdit_quick_manual_phone", 3, 1, 1, 3)
        put("label_quick_manual_postcode", 4, 0)
        put("lineEdit_quick_manual_postcode", 4, 1)
        if w("lineEdit_quick_manual_postcode") is not None:
            w("lineEdit_quick_manual_postcode").setMaximumWidth(120)
        put("pushButton_quick_manual_address_search", 4, 2, 1, 2)
        put("label_quick_manual_address", 5, 0)
        put("lineEdit_quick_manual_address", 5, 1)
        put("pushButton_quick_manual_postcode_lookup", 5, 2, 1, 2)
        put("label_quick_manual_detail", 6, 0)
        put("lineEdit_quick_manual_detail", 6, 1, 1, 3)
        put("label_quick_manual_product", 7, 0)
        put("lineEdit_quick_manual_product", 7, 1, 1, 3)
        # 마지막 줄: 초기화(왼쪽) · 생성(오른쪽, 이 화면의 다음 동작)
        put("pushButton_quick_excel_reset", 8, 0)
        put("pushButton_quick_excel_gen", 8, 2, 1, 2)
        if w("pushButton_quick_excel_gen") is not None:
            set_button_role(w("pushButton_quick_excel_gen"), "primary")
        gb.setMinimumWidth(380)
        gb.setMaximumWidth(520)

    def _reset_quick_excel_fields(self):
        """「초기화」 → 수동 주문 입력란을 모두 비웁니다(스토어 선택은 유지)."""
        for name in (
            "lineEdit_quick_manual_name",
            "lineEdit_quick_manual_phone",
            "lineEdit_quick_manual_postcode",
            "lineEdit_quick_manual_address",
            "lineEdit_quick_manual_detail",
            "lineEdit_quick_manual_product",
        ):
            if hasattr(self.ui, name):
                getattr(self.ui, name).clear()

    def _confirm_quick_excel_info(self, row):
        """클립보드에서 추출한 정보를 보여주고 생성을 확인합니다."""
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("퀵 엑셀 정보 확인")
        msg.setText(
            "아래 정보로 엑셀을 생성할까요?\n\n"
            f"수취인: {row.get('수취인명', '')}\n"
            f"연락처: {row.get('수취인 전화번호', '')}\n"
            f"우편번호: {row.get('우편번호') or '(없음)'}\n"
            f"주소: {row.get('수취인 주소', '')}\n"
            "상품명: 전자제품"
        )
        create_button = msg.addButton("생성", QMessageBox.AcceptRole)
        msg.addButton("취소", QMessageBox.RejectRole)
        msg.setDefaultButton(create_button)
        msg.exec()
        return msg.clickedButton() == create_button

    def generate_quick_excel(self):
        """클립보드 정보를 기반으로 단건 엑셀을 생성합니다. 수동 선택 시 입력란 값으로 생성합니다."""
        try:
            if hasattr(self.ui, "comboBox_store_select"):
                if self.ui.comboBox_store_select.currentText().strip() == "수동":
                    name = ""
                    phone = ""
                    address = ""
                    if hasattr(self.ui, "lineEdit_quick_manual_name"):
                        name = self.ui.lineEdit_quick_manual_name.text().strip()
                    if hasattr(self.ui, "lineEdit_quick_manual_phone"):
                        phone = self.ui.lineEdit_quick_manual_phone.text().strip()
                    if hasattr(self.ui, "lineEdit_quick_manual_address"):
                        address = self.ui.lineEdit_quick_manual_address.text().strip()
                    postcode = ""
                    detail = ""
                    if hasattr(self.ui, "lineEdit_quick_manual_postcode"):
                        postcode = self.ui.lineEdit_quick_manual_postcode.text().strip()
                    if hasattr(self.ui, "lineEdit_quick_manual_detail"):
                        detail = self.ui.lineEdit_quick_manual_detail.text().strip()
                    product_name = ""
                    if hasattr(self.ui, "lineEdit_quick_manual_product"):
                        product_name = self.ui.lineEdit_quick_manual_product.text().strip()
                    if not product_name:
                        product_name = "전자제품"
                    full_address = address
                    if detail:
                        full_address = f"{address} {detail}".strip()
                    missing = []
                    if not name:
                        missing.append("이름")
                    if not phone:
                        missing.append("전화번호")
                    if not address:
                        missing.append("기본주소")
                    if missing:
                        QMessageBox.warning(
                            self,
                            "입력 필요",
                            "수동 모드에서는 다음 항목을 모두 입력해 주세요.\n\n"
                            f"누락: {', '.join(missing)}",
                        )
                        return

                    invoice_data = [{
                        '주문번호': '',
                        '고객주문처명': '',
                        '수취인명': name,
                        '우편번호': postcode,
                        '수취인 주소': full_address,
                        '수취인 전화번호': phone,
                        '수취인 이동통신': phone,
                        '상품명': product_name,
                        '상품모델': '',
                        '배송메세지': '',
                        '비고': ''
                    }]
                    output_file = self.save_invoice_excel(invoice_data, "퀵_수동")
                    return

            clipboard = QApplication.clipboard()
            clipboard_text = clipboard.text()
            if not clipboard_text or not clipboard_text.strip():
                QMessageBox.warning(self, "경고", "클립보드에 복사된 내용이 없습니다.")
                return

            store_type = self.get_quick_excel_store_type(clipboard_text)
            if store_type == "coupang":
                quick_info = self.parse_coupang_quick_clipboard(clipboard_text)
                missing_fields = []
                if not quick_info.get("수취인명"):
                    missing_fields.append("수취인명")
                if not quick_info.get("연락처(안심번호)"):
                    missing_fields.append("연락처(안심번호)")
                if not quick_info.get("배송주소"):
                    missing_fields.append("배송주소")

                if missing_fields:
                    QMessageBox.warning(
                        self,
                        "오류",
                        "쿠팡 양식에서 필수 항목을 찾을 수 없습니다.\n\n"
                        f"누락 항목: {', '.join(missing_fields)}"
                    )
                    return

                invoice_data = [{
                    '주문번호': '',
                    '고객주문처명': '',
                    '수취인명': quick_info.get("수취인명", ""),
                    '우편번호': quick_info.get("우편번호", ""),
                    '수취인 주소': quick_info.get("배송주소", ""),
                    '수취인 전화번호': quick_info.get("연락처(안심번호)", ""),
                    '수취인 이동통신': quick_info.get("연락처(안심번호)", ""),
                    '상품명': '',
                    '상품모델': '',
                    '배송메세지': quick_info.get("배송메모", ""),
                    '비고': ''
                }]

                if not self._confirm_quick_excel_info(invoice_data[0]):
                    return
                output_file = self.save_invoice_excel(invoice_data, "퀵_쿠팡")
                return

            if store_type == "naver":
                quick_info = self.parse_naver_quick_clipboard(clipboard_text)
                missing_fields = []
                if not quick_info.get("수취인명"):
                    missing_fields.append("수취인명")
                if not quick_info.get("연락처1"):
                    missing_fields.append("연락처1")
                if not quick_info.get("배송지"):
                    missing_fields.append("배송지")

                if missing_fields:
                    QMessageBox.warning(
                        self,
                        "오류",
                        "네이버 양식에서 필수 항목을 찾을 수 없습니다.\n\n"
                        f"누락 항목: {', '.join(missing_fields)}"
                    )
                    return

                invoice_data = [{
                    '주문번호': '',
                    '고객주문처명': '',
                    '수취인명': quick_info.get("수취인명", ""),
                    '우편번호': '',
                    '수취인 주소': quick_info.get("배송지", ""),
                    '수취인 전화번호': quick_info.get("연락처1", ""),
                    '수취인 이동통신': quick_info.get("연락처2", ""),
                    '상품명': '',
                    '상품모델': '',
                    '배송메세지': quick_info.get("배송메모", ""),
                    '비고': ''
                }]

                if not self._confirm_quick_excel_info(invoice_data[0]):
                    return
                output_file = self.save_invoice_excel(invoice_data, "퀵_네이버")
                return

            if store_type == "gmarket":
                quick_info = self.parse_gmarket_quick_clipboard(clipboard_text)
                missing_fields = []
                if not quick_info.get("상품수령인"):
                    missing_fields.append("상품수령인")
                if not quick_info.get("연락처1"):
                    missing_fields.append("연락처1")
                if not quick_info.get("배송지주소"):
                    missing_fields.append("배송지주소")

                if missing_fields:
                    QMessageBox.warning(
                        self,
                        "오류",
                        "지마켓 양식에서 필수 항목을 찾을 수 없습니다.\n\n"
                        f"누락 항목: {', '.join(missing_fields)}"
                    )
                    return

                invoice_data = [{
                    '주문번호': '',
                    '고객주문처명': '',
                    '수취인명': quick_info.get("상품수령인", ""),
                    '우편번호': quick_info.get("우편번호", ""),
                    '수취인 주소': quick_info.get("배송지주소", ""),
                    '수취인 전화번호': quick_info.get("연락처1", ""),
                    '수취인 이동통신': quick_info.get("연락처2", ""),
                    '상품명': '',
                    '상품모델': '',
                    '배송메세지': quick_info.get("배송 요청사항", ""),
                    '비고': ''
                }]

                if not self._confirm_quick_excel_info(invoice_data[0]):
                    return
                output_file = self.save_invoice_excel(invoice_data, "퀵_지마켓")
                return

            if store_type == "generic":
                quick_info = _host(self)._parse_generic_quick_clipboard(clipboard_text)
                missing_fields = [
                    label for key, label in (
                        ("수취인명", "수령인/수취인"),
                        ("연락처", "연락처/전화번호"),
                        ("주소", "주소"),
                    ) if not quick_info[key]
                ]
                if missing_fields:
                    QMessageBox.warning(
                        self,
                        "오류",
                        "일반 양식에서 필수 항목을 찾을 수 없습니다.\n\n"
                        f"누락 항목: {', '.join(missing_fields)}",
                    )
                    return

                invoice_data = [{
                    '주문번호': '',
                    '고객주문처명': '',
                    '수취인명': quick_info["수취인명"],
                    '우편번호': quick_info["우편번호"],
                    '수취인 주소': quick_info["주소"],
                    '수취인 전화번호': quick_info["연락처"],
                    '수취인 이동통신': quick_info["연락처"],
                    '상품명': '전자제품',
                    '상품모델': '전자제품',
                    '배송메세지': '',
                    '비고': '',
                }]
                if not self._confirm_quick_excel_info(invoice_data[0]):
                    return
                output_file = self.save_invoice_excel(invoice_data, "퀵_일반")
                return

            QMessageBox.information(
                self,
                "안내",
                "클립보드에서 수령인(또는 수취인), 연락처, 주소를 찾지 못했습니다.\n"
                "각 항목을 '항목: 내용' 형식으로 복사했는지 확인해주세요."
            )
            return
        except Exception as e:
            error_msg = str(e)
            print(f"❌ 퀵 엑셀 생성 중 오류 발생: {error_msg}")
            QMessageBox.critical(
                self,
                "오류",
                "퀵 엑셀 생성 중 오류가 발생했습니다.\n\n"
                f"{error_msg}"
            )

    def get_quick_excel_store_type(self, clipboard_text):
        """콤보박스 또는 자동 판별로 스토어 타입을 결정합니다."""
        selected_text = ""
        if hasattr(self.ui, 'comboBox_store_select'):
            selected_text = self.ui.comboBox_store_select.currentText().strip()

        if selected_text == "수동":
            return "manual"
        if selected_text == "쿠팡":
            return "coupang"
        if selected_text == "네이버":
            return "naver"
        if selected_text == "지마켓":
            return "gmarket"

        return self.detect_store_from_clipboard(clipboard_text)

    def detect_store_from_clipboard(self, clipboard_text):
        """클립보드 텍스트로 스토어를 자동 판별합니다."""
        return detect_quick_store(clipboard_text)

    def parse_coupang_quick_clipboard(self, clipboard_text):
        """쿠팡 클립보드 텍스트에서 필수 정보를 추출합니다."""
        return parse_coupang_quick_clipboard(clipboard_text)

    def parse_naver_quick_clipboard(self, clipboard_text):
        """네이버 클립보드 텍스트에서 필수 정보를 추출합니다."""
        return parse_naver_quick_clipboard(clipboard_text)

    def parse_gmarket_quick_clipboard(self, clipboard_text):
        """지마켓 클립보드 텍스트에서 필수 정보를 추출합니다."""
        return parse_gmarket_quick_clipboard(clipboard_text)

    def extract_zip_code(self, address):
        """주소 문자열에서 우편번호를 추출합니다."""
        return extract_zip_code(address)

    def save_invoice_excel(self, invoice_data, filename_prefix):
        """송장 데이터로 엑셀 파일을 저장하고 경로를 반환합니다."""
        output_file = save_invoice_excel(invoice_data, filename_prefix)
        print(f"✓ 퀵 엑셀 저장 완료: {output_file}")
        self._refresh_quick_today_list()
        self.show_excel_created_message(output_file, "송장 엑셀 파일이 생성되었습니다.")
        return output_file
