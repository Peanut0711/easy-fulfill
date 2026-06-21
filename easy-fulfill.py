#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import re
import math
import time
from datetime import datetime, date
from pathlib import Path
import pandas as pd
import numpy as np
import subprocess
import tempfile
import msoffcrypto
import shutil
from PySide6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QMessageBox, 
                              QInputDialog, QLineEdit, QTableWidgetItem, QLabel, 
                              QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QWidget,
                              QProgressBar, QFrame, QGraphicsOpacityEffect, QListWidget,
                              QAbstractItemView)
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import (
    QFile,
    QIODevice,
    Qt,
    QSize,
    QUrl,
    QTimer,
    QThread,
    Signal,
    QPropertyAnimation,
    QEasingCurve,
    QRectF,
)
from PySide6.QtGui import (
    QPixmap,
    QImage,
    QIcon,
    QAction,
    QDesktopServices,
    QFont,
    QPainter,
    QColor,
    QPen,
)
import requests
from io import BytesIO
import warnings
import logging
import json
import traceback
import xml.etree.ElementTree as ET

try:
    import gspread
except ImportError:  # gspread는 스프레드시트 매핑에만 필요
    gspread = None
# 구글 스프레드시트(로컬 DB 대체) 설정
SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"
# 네이버 마크다운 상품 링크 (주문 처리 전용 고정 스토어)
NAVER_SMARTSTORE_PRODUCT_URL_PREFIX = "https://smartstore.naver.com/higenis/products/"
# 쿠팡 VP 상품 페이지 (스프레드시트 D열 상품번호 기준)
COUPANG_VP_PRODUCT_URL_PREFIX = "https://www.coupang.com/vp/products/"
# OAuth 경로·토큰: google_sheets_oauth.py (database-sync와 공유, google-oauth/)
# 일별 주문 인덱스(네이버·쿠팡·지마켓) 공유용 — 스프레드시트에서 세 번째 탭(gspread 워크시트 인덱스 2)
ORDER_INDEX_WORKSHEET_INDEX = 2
ORDER_INDEX_SHEET_TITLE = "일별 주문번호"
ORDER_INDEX_SHEET_HEADERS = ["날짜", "네이버", "쿠팡", "지마켓"]
ORDER_INDEX_SHEET_POLL_MS = 150_000  # 2.5분 (2~3분 간격)
ORDER_INDEX_SHEET_PUSH_DEBOUNCE_MS = 800
# 시작 시 DB동기화 로딩 바: 약 2초+여유 안에 99%까지 선형 증가, 완료 시 즉시 100%
STARTUP_SYNC_PROGRESS_CAP_MS = 2200
STARTUP_SYNC_PROGRESS_TICK_MS = 40
# 100% 도달 후 오버레이를 닫기까지 대기(너무 짧으면 급하게 사라진 느낌, 너무 길면 지루)
STARTUP_SYNC_PROGRESS_HIDE_DELAY_MS = 500
# 우정사업본부 KpostPortal 우편번호 통합검색 (target=postNew, UTF-8)
KPOST_OPENAPI2_URL = "http://biz.epost.go.kr/KpostPortal/openapi2"
# _show_kpost_api_error: 서버가 아닌 앱 내부용(인증키 미설정)
KPOST_ERROR_NO_REGKEY = "NO_REGKEY"


def _xml_local_lower(tag: str) -> str:
    if not tag:
        return ""
    return tag.split("}", 1)[-1].lower()


def _parse_order_index_int_cell(value, default=1):
    if value is None:
        return default
    s = str(value).strip()
    if not s:
        return default
    try:
        n = int(s.replace(",", ""))
        return n if n >= 1 else default
    except ValueError:
        return default


def _normalize_key_for_mapping_value(value):
    """상품번호/옵션ID 비교용 문자열로 정규화 (UI·백그라운드 매핑 로더 공통)."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _order_index_sheet_row_looks_like_header(row):
    if not row:
        return False
    cell = (row[0] or "").strip()
    return cell in ("날짜", "date", "Date")


def _ensure_order_index_worksheet_header_row(ws):
    """레거시 시트(1행이 바로 날짜 데이터)면 맨 위에 헤더 행을 삽입합니다."""
    values = ws.get_all_values()
    if not values or _order_index_sheet_row_looks_like_header(values[0]):
        return
    ws.insert_row(ORDER_INDEX_SHEET_HEADERS, index=1)
    print(
        f"✓ 「{ORDER_INDEX_SHEET_TITLE}」 시트 1행에 "
        "날짜·네이버·쿠팡·지마켓 열 제목을 추가했습니다. (기존 데이터는 한 행 아래로 밀렸습니다.)"
    )


def _standalone_open_order_index_ws(gc):
    """백그라운드 스레드용. worksheets() 메타데이터 한 번으로 워크시트를 찾고,
    없으면 생성합니다. (기존의 get_worksheet 중복 호출로 인한 추가 API 왕복 제거.
    빈 시트 초기화·레거시 헤더 처리는 값을 한 번 읽은 뒤
    _standalone_normalize_order_index_values 에서 재읽기 없이 수행합니다.)"""
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    worksheets = spreadsheet.worksheets()
    if len(worksheets) < ORDER_INDEX_WORKSHEET_INDEX:
        raise ValueError(
            f"스프레드시트에「{ORDER_INDEX_SHEET_TITLE}」({ORDER_INDEX_WORKSHEET_INDEX + 1}번째 탭)을 "
            f"둘 공간이 없습니다. 앞쪽 탭이 {ORDER_INDEX_WORKSHEET_INDEX}개(예: 네이버·쿠팡 DB) "
            f"있어야 합니다. (현재 탭 {len(worksheets)}개)"
        )
    if len(worksheets) == ORDER_INDEX_WORKSHEET_INDEX:
        # add_worksheet 는 생성된 워크시트를 반환하므로 별도 get_worksheet 불필요.
        ws = spreadsheet.add_worksheet(
            title=ORDER_INDEX_SHEET_TITLE,
            rows=100,
            cols=10,
            index=ORDER_INDEX_WORKSHEET_INDEX,
        )
        print(f"✓ 스프레드시트에 「{ORDER_INDEX_SHEET_TITLE}」 탭을 만들었습니다.")
        return ws
    return worksheets[ORDER_INDEX_WORKSHEET_INDEX]


def _standalone_normalize_order_index_values(ws, values):
    """단 한 번 읽은 get_all_values 결과로 빈 시트 초기화·레거시 헤더 삽입을
    처리하고, 시트를 다시 읽지 않고 최신 values 를 반환합니다."""
    if not values:
        today = date.today().strftime("%Y-%m-%d")
        ws.update(
            [ORDER_INDEX_SHEET_HEADERS, [today, 1, 1, 1]],
            range_name="A1:D2",
        )
        print(
            f"✓ 「{ORDER_INDEX_SHEET_TITLE}」 시트에 헤더와 오늘({today}) 인덱스 1·1·1을 초기 입력했습니다."
        )
        return [list(ORDER_INDEX_SHEET_HEADERS), [today, "1", "1", "1"]]
    if not _order_index_sheet_row_looks_like_header(values[0]):
        ws.insert_row(ORDER_INDEX_SHEET_HEADERS, index=1)
        print(
            f"✓ 「{ORDER_INDEX_SHEET_TITLE}」 시트 1행에 "
            "날짜·네이버·쿠팡·지마켓 열 제목을 추가했습니다. (기존 데이터는 한 행 아래로 밀렸습니다.)"
        )
        return [list(ORDER_INDEX_SHEET_HEADERS)] + values
    return values


def _standalone_read_today_order_indices_from_values(values):
    if not values:
        return None
    today = date.today().strftime("%Y-%m-%d")
    header_like = _order_index_sheet_row_looks_like_header(values[0])
    data_rows = values[1:] if header_like else values
    for row in data_rows:
        if not row:
            continue
        if (row[0] or "").strip() != today:
            continue
        naver = _parse_order_index_int_cell(row[1] if len(row) > 1 else None)
        coupang = _parse_order_index_int_cell(row[2] if len(row) > 2 else None)
        gmarket = _parse_order_index_int_cell(row[3] if len(row) > 3 else None)
        return {"naver": naver, "coupang": coupang, "gmarket": gmarket}
    return None


def _standalone_write_order_indices_ws(ws, n, c, g, values=None):
    today = date.today().strftime("%Y-%m-%d")
    if values is None:
        values = ws.get_all_values()
    if not values:
        ws.update(
            [ORDER_INDEX_SHEET_HEADERS, [today, n, c, g]],
            range_name="A1:D2",
        )
        return
    header_like = _order_index_sheet_row_looks_like_header(values[0])
    if header_like:
        search_ranges = list(enumerate(values[1:], start=2))
    else:
        search_ranges = list(enumerate(values, start=1))
    row_1based = None
    for ridx, row in search_ranges:
        if row and (row[0] or "").strip() == today:
            row_1based = ridx
            break
    if row_1based is not None:
        ws.update([[n, c, g]], range_name=f"B{row_1based}:D{row_1based}")
    else:
        ws.append_row([today, n, c, g])


def run_startup_order_index_sheet_sync_worker():
    """
    UI 스레드가 아닌 곳에서 호출. Google API만 수행.
    반환 dict: ok, kind(data|created), row — 또는 ok False, error
    """
    if gspread is None:
        return {"ok": False, "error": "gspread 패키지가 필요합니다. (pip install gspread)"}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as e:
        return {"ok": False, "error": str(e)}
    try:
        gc = get_authorized_gspread_client()
        ws = _standalone_open_order_index_ws(gc)
        values = ws.get_all_values()  # 시트 전체 읽기는 시작 동기화에서 단 한 번만 수행
        values = _standalone_normalize_order_index_values(ws, values)
        row = _standalone_read_today_order_indices_from_values(values)
        if row is not None:
            return {"ok": True, "kind": "data", "row": row}
        _standalone_write_order_indices_ws(ws, 1, 1, 1, values=values)
        return {
            "ok": True,
            "kind": "created",
            "row": {"naver": 1, "coupang": 1, "gmarket": 1},
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


class CircularBusySpinner(QWidget):
    """Google 시트·처리 대기용 무한 회전 링 스피너."""

    def __init__(self, parent=None, size: int = 58, line_width: int = 4, color: str = "#21838a"):
        super().__init__(parent)
        self._line_width = line_width
        self._color = QColor(color)
        self._rotation = 0
        self.setFixedSize(size, size)
        self._timer = QTimer(self)
        self._timer.setInterval(40)
        self._timer.timeout.connect(self._on_tick)

    def _on_tick(self):
        self._rotation = (self._rotation + 24) % 360
        self.update()

    def showEvent(self, event):
        super().showEvent(event)
        self._timer.start()

    def hideEvent(self, event):
        self._timer.stop()
        super().hideEvent(event)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        s = min(self.width(), self.height())
        m = self._line_width + 1
        rf = (s - 2 * m) / 2.0
        p.translate(self.width() / 2.0, self.height() / 2.0)
        p.rotate(self._rotation)
        pen = QPen(self._color)
        pen.setWidth(self._line_width)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(pen)
        rect = QRectF(-rf, -rf, 2 * rf, 2 * rf)
        p.drawArc(rect, 40 * 16, 290 * 16)


def run_product_code_map_load_worker(store_type: str):
    """
    백그라운드에서 상품코드 매핑만 로드 (_load_product_code_map_from_spreadsheet 와 동일 데이터).
    MainWindow._gspread_client 와 별도 클라이언트 사용.
    """
    if gspread is None:
        return {
            "ok": False,
            "error": "gspread 패키지가 필요합니다. (pip install gspread)",
            "store_type": store_type,
        }
    if store_type == "naver":
        sheet_index = 0
        header_row_num = 1
    elif store_type == "coupang":
        sheet_index = 1
        header_row_num = 2
    else:
        return {"ok": False, "error": f"지원되지 않는 store_type: {store_type}", "store_type": store_type}
    try:
        from google_sheets_oauth import get_authorized_gspread_client
    except ImportError as e:
        return {"ok": False, "error": str(e), "store_type": store_type}
    try:
        gc = get_authorized_gspread_client()
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.get_worksheet(sheet_index)
        values = worksheet.get_all_values()
        data_start_idx = header_row_num

        mapping = {}
        coupang_vp = {}
        for row in values[data_start_idx:]:
            product_code = row[0].strip() if len(row) > 0 else ""
            key_value = row[4] if len(row) > 4 else ""

            key_norm = _normalize_key_for_mapping_value(key_value)
            if not key_norm:
                continue

            product_code_norm = str(product_code).strip()
            if product_code_norm.lower() == "nan":
                product_code_norm = ""

            mapping[key_norm] = product_code_norm
            if store_type == "coupang":
                vp_raw = row[3] if len(row) > 3 else ""
                vp_norm = _normalize_key_for_mapping_value(vp_raw)
                coupang_vp[key_norm] = vp_norm

        print(f"✓ 스프레드시트 매핑 로드 완료: {store_type} - {len(mapping)}개")
        out = {"ok": True, "store_type": store_type, "mapping": mapping}
        if store_type == "coupang":
            out["coupang_vp"] = coupang_vp
        return out
    except Exception as e:
        return {"ok": False, "error": str(e), "store_type": store_type}


class ProductMappingLoadThread(QThread):
    """주문 엑셀 열 때 상품 매핑 시트 읽기만 백그라운드에서 수행."""

    result_ready = Signal(dict)

    def __init__(self, store_type: str):
        super().__init__()
        self._store_type = store_type

    def run(self):
        self.result_ready.emit(run_product_code_map_load_worker(self._store_type))


class OrderIndexStartupSyncThread(QThread):
    """앱 시작 시 스프레드시트 인덱스만 백그라운드에서 맞춥니다."""

    result_ready = Signal(dict)

    def run(self):
        self.result_ready.emit(run_startup_order_index_sheet_sync_worker())


class DbSheetSyncThread(QThread):
    """database 폴더 내보내기 → 스프레드시트 신규 행 append (백그라운드)."""

    result_ready = Signal(dict)

    def __init__(self, params: dict):
        super().__init__()
        self._params = params

    def run(self):
        p = self._params
        try:
            from db_sheet_sync import run_db_sheet_sync_job

            self.result_ready.emit(
                run_db_sheet_sync_job(
                    p["spreadsheet_id"],
                    do_naver=p["do_naver"],
                    do_coupang=p["do_coupang"],
                    test_mode=p["test_mode"],
                    test_count=p["test_count"],
                    naver_path=p.get("naver_path"),
                    coupang_path=p.get("coupang_path"),
                    verbose_log=p.get("verbose_log", True),
                )
            )
        except Exception as e:
            self.result_ready.emit(
                {
                    "ok": False,
                    "logs": [f"❌ DB동기화 작업 중 예외: {e}"],
                    "test_mode": p.get("test_mode"),
                    "naver": None,
                    "coupang": None,
                    "error": str(e),
                }
            )


class ImageDialog(QDialog):
    def __init__(self, image_url, product_name, parent=None, all_images=None, current_index=None, table_widget=None):
        super().__init__(parent)
        self.current_product_name = product_name  # 현재 상품명 저장
        self.current_subcategory = ""  # 현재 소분류 저장
        
        # 이미지 목록 및 현재 인덱스 저장
        self.all_images = all_images or []
        self.current_index = current_index or 0
        self.table_widget = table_widget  # 테이블 위젯 참조 저장
        
        # 레이아웃 설정
        main_layout = QVBoxLayout()
        
        # 이미지 레이블
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumSize(800, 600)  # 최소 크기 설정
        self.image_label.mousePressEvent = self.image_clicked  # 이미지 클릭 이벤트 추가
        main_layout.addWidget(self.image_label)
        
        # 소분류 입력 필드 추가 (중앙 배치)
        subcategory_layout = QHBoxLayout()
        # 왼쪽 스페이서 추가
        subcategory_layout.addStretch(1)
        
        subcategory_label = QLabel("소분류:")
        self.subcategory_input = QLineEdit()
        self.subcategory_input.setPlaceholderText("소분류 입력")
        self.subcategory_input.setMinimumWidth(200)  # 최소 너비 설정
        self.subcategory_input.returnPressed.connect(self.handle_subcategory_enter)
        self.subcategory_input.textChanged.connect(self.update_title_with_subcategory)
        
        subcategory_layout.addWidget(subcategory_label)
        subcategory_layout.addWidget(self.subcategory_input)
        
        # 오른쪽 스페이서 추가
        subcategory_layout.addStretch(1)
        
        main_layout.addLayout(subcategory_layout)
        
        # 버튼 레이아웃 (이전/다음/닫기) - 중앙 배치
        button_layout = QHBoxLayout()
        
        # 왼쪽 스페이서 추가
        button_layout.addStretch(1)
        
        # 이전 버튼
        self.prev_button = QPushButton("이전")
        self.prev_button.setMaximumWidth(100)
        self.prev_button.clicked.connect(self.show_previous_image)
        button_layout.addWidget(self.prev_button)
        
        # 닫기 버튼
        close_button = QPushButton("닫기")
        close_button.setMaximumWidth(100)
        close_button.clicked.connect(self.close)
        button_layout.addWidget(close_button)
        
        # 다음 버튼
        self.next_button = QPushButton("다음 (Ctrl+Space)")
        self.next_button.setMaximumWidth(150)
        self.next_button.clicked.connect(self.show_next_image)
        button_layout.addWidget(self.next_button)
        
        # 오른쪽 스페이서 추가
        button_layout.addStretch(1)
        
        main_layout.addLayout(button_layout)
        
        self.setLayout(main_layout)
        
        # 이미지 로드
        self.load_image(image_url)
        
        # 창 크기 설정
        self.resize(1024, 768)  # 더 큰 초기 크기로 설정
        
        # 이전/다음 버튼 상태 업데이트
        self.update_navigation_buttons()
        
        # 키보드 이벤트 처리
        self.setFocusPolicy(Qt.StrongFocus)
        
        # 마우스 휠 이벤트 활성화
        self.setMouseTracking(True)
        
        # 현재 상품의 소분류 값 가져오기
        self.load_subcategory_from_table()
        
        # 초기 타이틀 설정
        self.update_title_with_subcategory()
    
    def update_title_with_subcategory(self):
        """소분류 정보를 포함하여 타이틀 업데이트"""
        subcategory_text = self.subcategory_input.text().strip()
        if subcategory_text:
            self.setWindowTitle(f"상품이미지 - '{self.current_product_name}' [ 소분류: {subcategory_text} ]")
        else:
            self.setWindowTitle(f"상품이미지 - '{self.current_product_name}' [ 소분류: - ]")
    
    def handle_subcategory_enter(self):
        """소분류 입력 필드에서 엔터키를 눌렀을 때 처리"""
        # 현재 소분류 값 업데이트
        self.update_subcategory()
        
        # 다음 이미지로 이동 (다음 버튼이 활성화된 경우에만)
        if self.next_button.isEnabled():
            self.show_next_image()
    
    def load_subcategory_from_table(self):
        """테이블에서 현재 상품의 소분류 값을 가져와 입력 필드에 설정"""
        if self.table_widget:
            for row in range(self.table_widget.rowCount()):
                product_item = self.table_widget.item(row, 3)  # 상품명 열
                if product_item and product_item.text() == self.current_product_name:
                    subcategory_item = self.table_widget.item(row, 2)  # 소분류 열
                    if subcategory_item:
                        self.subcategory_input.setText(subcategory_item.text())
                    break
    
    def update_subcategory(self):
        """소분류 입력 필드의 값을 테이블에 업데이트"""
        if self.table_widget:
            for row in range(self.table_widget.rowCount()):
                product_item = self.table_widget.item(row, 3)  # 상품명 열
                if product_item and product_item.text() == self.current_product_name:
                    # 소분류 열에 값 설정
                    subcategory_item = QTableWidgetItem(self.subcategory_input.text())
                    self.table_widget.setItem(row, 2, subcategory_item)
                    break
    
    def image_clicked(self, event):
        """이미지 클릭 이벤트 처리"""
        self.close()
    
    def wheelEvent(self, event):
        """마우스 휠 이벤트 처리"""
        # 휠 방향에 따라 이전/다음 이미지 표시
        if event.angleDelta().y() > 0:  # 위로 스크롤
            self.show_previous_image()
        else:  # 아래로 스크롤
            self.show_next_image()
        event.accept()
    
    def keyPressEvent(self, event):
        """키보드 이벤트 처리"""
        if event.key() == Qt.Key_Left:
            self.show_previous_image()
        elif event.key() == Qt.Key_Right:
            self.show_next_image()
        elif event.key() == Qt.Key_Escape:
            self.close()
        elif event.key() == Qt.Key_Space and event.modifiers() == Qt.ControlModifier:
            # Ctrl+Space 단축키로 다음 이미지로 이동
            if self.next_button.isEnabled():
                self.show_next_image()
        else:
            super().keyPressEvent(event)
    
    def show_previous_image(self):
        """이전 이미지 표시"""
        if self.all_images and self.current_index > 0:
            self.current_index -= 1
            self.show_current_image()
    
    def show_next_image(self):
        """다음 이미지 표시"""
        if self.all_images and self.current_index < len(self.all_images) - 1:
            self.current_index += 1
            self.show_current_image()
    
    def show_current_image(self):
        """현재 인덱스의 이미지 표시"""
        if self.all_images and 0 <= self.current_index < len(self.all_images):
            image_info = self.all_images[self.current_index]
            self.current_product_name = image_info['product_name']  # 현재 상품명 업데이트
            self.load_image(image_info['image_url'])
            self.update_navigation_buttons()
            self.load_subcategory_from_table()  # 소분류 값 로드
            self.update_title_with_subcategory()  # 타이틀 업데이트
    
    def update_navigation_buttons(self):
        """이전/다음 버튼 상태 업데이트"""
        if not self.all_images:
            self.prev_button.setEnabled(False)
            self.next_button.setEnabled(False)
            return
            
        self.prev_button.setEnabled(self.current_index > 0)
        self.next_button.setEnabled(self.current_index < len(self.all_images) - 1)
    
    def closeEvent(self, event):
        """창이 닫힐 때 호출되는 이벤트"""
        # 소분류 값 업데이트
        self.update_subcategory()
        
        # 현재 표시 중인 상품의 이름을 사용하여 테이블에서 해당 행 찾기
        if self.table_widget:
            print(f"창 닫힘: 현재 상품명 '{self.current_product_name}'로 테이블 검색")
            for row in range(self.table_widget.rowCount()):
                product_item = self.table_widget.item(row, 3)  # 상품명 열
                if product_item and product_item.text() == self.current_product_name:
                    print(f"상품 '{self.current_product_name}'를 테이블의 {row}번 행에서 찾음")
                    # 해당 행 선택
                    self.table_widget.selectRow(row)
                    # 해당 행으로 스크롤
                    self.table_widget.scrollToItem(product_item)
                    break
        
        super().closeEvent(event)
    
    def load_image(self, image_url):
        try:
            # 이미지 다운로드
            response = requests.get(image_url)
            image_data = BytesIO(response.content)
            
            # QPixmap으로 변환
            pixmap = QPixmap()
            pixmap.loadFromData(image_data.getvalue())
            
            # 이미지 크기 조정 (최대 크기 설정)
            scaled_pixmap = pixmap.scaled(
                800, 600,  # 고정된 최대 크기
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            
            self.image_label.setPixmap(scaled_pixmap)
            
        except Exception as e:
            self.image_label.setText(f"이미지를 불러올 수 없습니다.\n{str(e)}")


def _is_likely_google_sheets_oauth_error(exc: BaseException) -> bool:
    """스프레드시트 OAuth·토큰 문제로 복구(재인증) 가능한 오류인지 휴리스틱 판별."""
    text = str(exc).lower()
    for needle in (
        "invalid_scope",
        "invalid_grant",
        "token has been expired",
        "token expired",
        "could not refresh",
        "bad request",
        "unauthorized",
        "access denied",
        "refresh",
    ):
        if needle in text:
            return True
    mod = getattr(type(exc), "__module__", "") or ""
    if "google.auth" in mod or "google_auth" in mod:
        return True
    return False


class KpostAddressSearchDialog(QDialog):
    """우체국 KpostPortal(postNew)으로 주소를 검색·선택합니다."""

    def __init__(self, host: "MainWindow", parent=None):
        super().__init__(parent or host)
        self.setWindowTitle("우체국 주소 검색")
        self.resize(540, 440)
        self._host = host
        self._items = []
        self.selected_postcd = ""
        self.selected_address = ""
        self.detail_address = ""

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel("검색어 (도로명+번지, 읍면동+지번, 건물명 등 · 2자 이상)")
        )
        self._query = QLineEdit()
        layout.addWidget(self._query)
        row = QHBoxLayout()
        btn_search = QPushButton("검색")
        btn_search.clicked.connect(self._on_search_clicked)
        row.addWidget(btn_search)
        row.addStretch()
        layout.addLayout(row)
        layout.addWidget(QLabel("검색 결과에서 한 건을 선택하세요."))
        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self._list, 1)
        layout.addWidget(QLabel("상세주소 (동·호수 등, 선택)"))
        self._detail = QLineEdit()
        layout.addWidget(self._detail)
        btn_row = QHBoxLayout()
        btn_ok = QPushButton("확인")
        btn_cancel = QPushButton("취소")
        btn_ok.clicked.connect(self._on_accept)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

    def _on_search_clicked(self):
        q = self._query.text().strip()
        if len(q) < 2:
            QMessageBox.warning(self, "입력", "검색어는 2자 이상 입력해 주세요.")
            return
        items, err = self._host._kpost_postnew_lookup_items(q)
        if err:
            self._host._show_kpost_api_error(self, err)
            return
        if not items:
            QMessageBox.information(
                self,
                "결과 없음",
                "검색 결과가 없습니다.\n검색어를 더 구체적으로 바꿔 보세요.",
            )
            return
        self._items = items
        self._list.clear()
        for it in items:
            self._list.addItem(f"[{it['postcd']}] {it['address']}")

    def _on_accept(self):
        row = self._list.currentRow()
        if row < 0 or row >= len(self._items):
            QMessageBox.warning(
                self, "선택 필요", "목록에서 주소 한 건을 선택한 뒤 확인을 눌러 주세요."
            )
            return
        it = self._items[row]
        self.selected_postcd = it.get("postcd", "")
        self.selected_address = it.get("address", "")
        self.detail_address = self._detail.text().strip()
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        print("초기화 시작")
        self.selected_file_path = None
        self.store_type = None
        self.orders = {}  # orders 변수를 인스턴스 변수로 초기화
        self.is_order_file_valid = False  # 주문서 파일 유효성 플래그
        self.is_invoice_file_valid = False  # 송장 파일 유효성 플래그
        
        # 스프레드시트 매핑 캐시 (네이버/쿠팡)
        self._spreadsheet_product_code_maps = {}
        # 쿠팡: 옵션ID(E) 정규화 키 -> 스프레드시트 D열 상품번호(VP URL용)
        self._coupang_option_to_vp_product_no = {}
        self._gspread_client = None
        
        # 인덱스 파일 경로 설정
        self.index_file_path = Path("database") / "order_index.json"
        self.app_settings_path = Path("database") / "app_settings.json"
        
        # 인덱스 값 초기화
        self.current_idx_naver = 1
        self.current_idx_coupang = 1
        self.current_idx_gmarket = 1

        self._index_sheet_push_pending = False
        self._index_sheet_last_sync_display = None
        self._index_sheet_push_timer = QTimer(self)
        self._index_sheet_push_timer.setSingleShot(True)
        self._index_sheet_push_timer.timeout.connect(self._flush_order_indices_to_sheet)
        self._index_sheet_poll_timer = QTimer(self)
        self._index_sheet_poll_timer.timeout.connect(self._on_order_index_sheet_poll)
        self._last_refresh_created_today_row = False
        self._startup_order_index_sync_started = False
        self._startup_order_index_sync_done = False
        self._startup_sync_thread = None
        self._product_mapping_thread = None
        self._db_sheet_sync_thread = None
        self._db_sync_naver_path_override = None
        self._db_sync_coupang_path_override = None

        self.load_ui()
        self.setup_connections()
        self.setup_status_bar()
        self._setup_startup_loading_overlay()
        self._setup_busy_processing_overlay()
        
        # 인덱스 값 로드
        self.load_index_values()
        self.load_app_settings()
        self._refresh_google_auth_status_ui()
        self._refresh_db_sheet_sync_path_labels()

        print("초기화 완료")

    def _normalize_key_for_mapping(self, value):
        """상품번호/옵션ID 비교용 문자열로 정규화합니다."""
        return _normalize_key_for_mapping_value(value)

    def _format_product_name_markdown_link(self, product_name, id_normalized, url_prefix):
        """id가 숫자만일 때만 상품명을 마크다운 링크로 감쌉니다. 그 외는 원문 유지."""
        if not product_name or not product_name.strip():
            return product_name or ""
        if not id_normalized or not id_normalized.isdigit():
            return product_name
        if "]" in product_name:
            return product_name
        return f"[{product_name}]({url_prefix}{id_normalized})"

    def _format_naver_product_name_markdown(self, product_name, product_no_normalized):
        return self._format_product_name_markdown_link(
            product_name, product_no_normalized, NAVER_SMARTSTORE_PRODUCT_URL_PREFIX
        )

    def _load_product_code_map_from_spreadsheet(self, store_type):
        """
        스프레드시트에서 상품코드 매핑을 로드합니다.

        - 네이버(1번 시트): A열=상품코드, E열=상품번호(스마트스토어)
        - 쿠팡(2번 시트): A열=상품코드, D열=상품번호(VP URL), E열=옵션 ID
        """
        if store_type in self._spreadsheet_product_code_maps:
            return self._spreadsheet_product_code_maps[store_type]

        if gspread is None:
            raise ImportError("gspread 패키지가 필요합니다. (pip install gspread)")

        if store_type == "naver":
            sheet_index = 0
            header_row_num = 1  # 1행 헤더
        elif store_type == "coupang":
            sheet_index = 1
            header_row_num = 2  # 2행 헤더
        else:
            raise ValueError(f"지원되지 않는 store_type: {store_type}")

        gc = self._get_gspread_client()
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.get_worksheet(sheet_index)

        values = worksheet.get_all_values()
        data_start_idx = header_row_num  # 0-based index에서 header 다음 행

        mapping = {}
        if store_type == "coupang":
            self._coupang_option_to_vp_product_no = {}
        for row in values[data_start_idx:]:
            # A(0)=상품코드, D(3)=쿠팡 상품번호(VP), E(4)=키(상품번호/옵션ID)
            product_code = row[0].strip() if len(row) > 0 else ""
            key_value = row[4] if len(row) > 4 else ""

            key_norm = _normalize_key_for_mapping_value(key_value)
            if not key_norm:
                continue

            product_code_norm = str(product_code).strip()
            if product_code_norm.lower() == "nan":
                product_code_norm = ""

            mapping[key_norm] = product_code_norm
            if store_type == "coupang":
                vp_raw = row[3] if len(row) > 3 else ""
                vp_norm = self._normalize_key_for_mapping(vp_raw)
                self._coupang_option_to_vp_product_no[key_norm] = vp_norm

        self._spreadsheet_product_code_maps[store_type] = mapping
        print(f"✓ 스프레드시트 매핑 로드 완료: {store_type} - {len(mapping)}개")
        return mapping

    def _get_gspread_client(self):
        """OAuth(사용자별 로그인)로 gspread 클라이언트를 생성/재사용합니다."""
        if self._gspread_client is not None:
            return self._gspread_client

        if gspread is None:
            raise ImportError("gspread 패키지가 필요합니다. (pip install gspread)")

        try:
            from google_sheets_oauth import get_authorized_gspread_client
        except ImportError as e:
            raise ImportError(
                "google-auth-oauthlib 패키지가 필요합니다. "
                "(pip install google-auth-oauthlib)"
            ) from e

        self._gspread_client = get_authorized_gspread_client()
        return self._gspread_client

    def load_app_settings(self):
        """앱 설정(체크박스 등)을 로드합니다."""
        key = "auto_generate_after_invoice_load"
        default_checked = False
        if not hasattr(self.ui, "checkBox_invoice_load_auto_generate"):
            return
        cb = self.ui.checkBox_invoice_load_auto_generate
        checked = default_checked
        try:
            if self.app_settings_path.exists() and self.app_settings_path.stat().st_size > 0:
                with open(self.app_settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and key in data:
                    checked = bool(data[key])
        except Exception as e:
            print(f"! 앱 설정 로드 중 오류: {e}")
        cb.blockSignals(True)
        cb.setChecked(checked)
        cb.blockSignals(False)

    def save_app_settings(self):
        """앱 설정을 database/app_settings.json 에 저장합니다."""
        if not hasattr(self.ui, "checkBox_invoice_load_auto_generate"):
            return
        try:
            self.app_settings_path.parent.mkdir(parents=True, exist_ok=True)
            data = {}
            if self.app_settings_path.exists() and self.app_settings_path.stat().st_size > 0:
                try:
                    with open(self.app_settings_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except json.JSONDecodeError:
                    data = {}
            if not isinstance(data, dict):
                data = {}
            data["auto_generate_after_invoice_load"] = (
                self.ui.checkBox_invoice_load_auto_generate.isChecked()
            )
            with open(self.app_settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print("✓ 앱 설정 저장 완료")
        except Exception as e:
            print(f"! 앱 설정 저장 중 오류: {e}")

    def load_index_values(self):
        """저장된 인덱스 값을 로드합니다."""
        self._block_index_line_edit_signals(True)
        try:
            if self.index_file_path.exists():
                with open(self.index_file_path, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)
                
                # 오늘 날짜의 인덱스 값 확인
                today = date.today().strftime('%Y-%m-%d')
                
                if today in index_data:
                    # 네이버 인덱스
                    if 'naver' in index_data[today]:
                        self.current_idx_naver = index_data[today]['naver']
                        if hasattr(self.ui, 'lineEdit_idx_naver'):
                            self.ui.lineEdit_idx_naver.setText(str(self.current_idx_naver))
                    # 지마켓 인덱스
                    if 'gmarket' in index_data[today]:
                        self.current_idx_gmarket = index_data[today]['gmarket']
                        if hasattr(self.ui, 'lineEdit_idx_gmarket'):
                            self.ui.lineEdit_idx_gmarket.setText(str(self.current_idx_gmarket))
                    # 쿠팡 인덱스
                    if 'coupang' in index_data[today]:
                        self.current_idx_coupang = index_data[today]['coupang']
                        if hasattr(self.ui, 'lineEdit_idx_coupang'):
                            self.ui.lineEdit_idx_coupang.setText(str(self.current_idx_coupang))
                
                print(f"✓ 인덱스 값 로드 완료 (날짜: {today})")
                print(f"  - 네이버: {self.current_idx_naver}")
                print(f"  - 지마켓: {self.current_idx_gmarket}")
                print(f"  - 쿠팡: {self.current_idx_coupang}")
            else:
                print("! 인덱스 파일이 없습니다. 기본값(1)을 사용합니다.")                        
                if hasattr(self.ui, 'lineEdit_idx_naver'):
                    self.ui.lineEdit_idx_naver.setText('1')
                if hasattr(self.ui, 'lineEdit_idx_gmarket'):
                    self.ui.lineEdit_idx_gmarket.setText('1')
                if hasattr(self.ui, 'lineEdit_idx_coupang'):
                    self.ui.lineEdit_idx_coupang.setText('1')
        except Exception as e:
            print(f"! 인덱스 값 로드 중 오류 발생: {str(e)}")
            print("! 기본값(1)을 사용합니다.")                
            if hasattr(self.ui, 'lineEdit_idx_naver'):
                self.ui.lineEdit_idx_naver.setText('1')
            if hasattr(self.ui, 'lineEdit_idx_gmarket'):
                self.ui.lineEdit_idx_gmarket.setText('1')
            if hasattr(self.ui, 'lineEdit_idx_coupang'):
                self.ui.lineEdit_idx_coupang.setText('1')
        finally:
            self._block_index_line_edit_signals(False)

    def _persist_index_values_to_json(self):
        """order_index.json 에만 저장합니다(스프레드시트 푸시 없음)."""
        self.index_file_path.parent.mkdir(exist_ok=True)

        index_data = {}
        if self.index_file_path.exists() and self.index_file_path.stat().st_size > 0:
            try:
                with open(self.index_file_path, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)
            except json.JSONDecodeError:
                print("! 인덱스 파일이 손상되었습니다. 새로 생성합니다.")
                index_data = {}

        today = date.today().strftime('%Y-%m-%d')
        prev = index_data.get(today, {})
        if (
            prev.get("naver") == self.current_idx_naver
            and prev.get("gmarket") == self.current_idx_gmarket
            and prev.get("coupang") == self.current_idx_coupang
        ):
            return

        if today not in index_data:
            index_data[today] = {}

        index_data[today]['naver'] = self.current_idx_naver
        index_data[today]['gmarket'] = self.current_idx_gmarket
        index_data[today]['coupang'] = self.current_idx_coupang

        with open(self.index_file_path, 'w', encoding='utf-8') as f:
            json.dump(index_data, f, ensure_ascii=False, indent=2)

        print(f"✓ 인덱스 값 저장 완료 (날짜: {today})")
        print(f"  - 네이버: {self.current_idx_naver}")
        print(f"  - 지마켓: {self.current_idx_gmarket}")
        print(f"  - 쿠팡: {self.current_idx_coupang}")

    def save_index_values(self, push_sheet=True):
        """현재 인덱스 값을 저장합니다. 로컬 JSON 후 스프레드시트 반영은 디바운스됩니다."""
        try:
            self._persist_index_values_to_json()
        except Exception as e:
            print(f"! 인덱스 값 저장 중 오류 발생: {str(e)}")
            if self.index_file_path.exists():
                backup_path = self.index_file_path.with_suffix('.json.bak')
                try:
                    shutil.copy2(self.index_file_path, backup_path)
                    print(f"! 기존 인덱스 파일을 백업했습니다: {backup_path}")
                except Exception as backup_error:
                    print(f"! 백업 파일 생성 실패: {str(backup_error)}")
            return

        self._update_index_sheet_sync_label()
        if push_sheet:
            self._schedule_order_index_sheet_push()

    @staticmethod
    def _parse_order_index_cell(value, default=1):
        """스프레드시트·입력 칸에서 정수 인덱스로 변환합니다."""
        if value is None:
            return default
        s = str(value).strip()
        if not s:
            return default
        try:
            n = int(s.replace(",", ""))
            return n if n >= 1 else default
        except ValueError:
            return default

    def _block_index_line_edit_signals(self, block):
        for name in ("lineEdit_idx_naver", "lineEdit_idx_coupang", "lineEdit_idx_gmarket"):
            w = getattr(self.ui, name, None)
            if w:
                w.blockSignals(block)

    def _init_order_index_worksheet_if_blank(self, ws):
        """인덱스 시트가 비어 있으면 헤더·오늘 날짜·1,1,1 행을 씁니다."""
        if ws.get_all_values():
            return
        today = date.today().strftime("%Y-%m-%d")
        ws.update(
            [ORDER_INDEX_SHEET_HEADERS, [today, 1, 1, 1]],
            range_name="A1:D2",
        )
        print(
            f"✓ 「{ORDER_INDEX_SHEET_TITLE}」 시트에 헤더와 오늘({today}) 인덱스 1·1·1을 초기 입력했습니다."
        )

    def _get_order_index_worksheet(self):
        if gspread is None:
            raise ImportError("gspread 패키지가 필요합니다. (pip install gspread)")
        gc = self._get_gspread_client()
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        worksheets = spreadsheet.worksheets()

        # 3번째 탭(index 2)까지 앞에 시트가 부족하면 자동 생성 불가(네이버·쿠팡 탭 먼저 필요)
        if len(worksheets) < ORDER_INDEX_WORKSHEET_INDEX:
            raise ValueError(
                f"스프레드시트에「{ORDER_INDEX_SHEET_TITLE}」({ORDER_INDEX_WORKSHEET_INDEX + 1}번째 탭)을 "
                f"둘 공간이 없습니다. 앞쪽 탭이 {ORDER_INDEX_WORKSHEET_INDEX}개(예: 네이버·쿠팡 DB) "
                f"있어야 합니다. (현재 탭 {len(worksheets)}개)"
            )

        if len(worksheets) == ORDER_INDEX_WORKSHEET_INDEX:
            spreadsheet.add_worksheet(
                title=ORDER_INDEX_SHEET_TITLE,
                rows=100,
                cols=10,
                index=ORDER_INDEX_WORKSHEET_INDEX,
            )
            print(f"✓ 스프레드시트에 「{ORDER_INDEX_SHEET_TITLE}」 탭을 만들었습니다.")

        ws = spreadsheet.get_worksheet(ORDER_INDEX_WORKSHEET_INDEX)
        self._init_order_index_worksheet_if_blank(ws)
        _ensure_order_index_worksheet_header_row(ws)
        return ws

    def _sheet_row_looks_like_order_index_header(self, row):
        if not row:
            return False
        cell = (row[0] or "").strip()
        return cell in ("날짜", "date", "Date")

    def _read_today_order_indices_from_sheet(self):
        """OAuth·시트 OK이고 오늘 행이 있을 때만 dict, 없으면 None."""
        ws = self._get_order_index_worksheet()
        values = ws.get_all_values()
        if not values:
            return None
        today = date.today().strftime("%Y-%m-%d")
        header_like = self._sheet_row_looks_like_order_index_header(values[0])
        data_rows = values[1:] if header_like else values
        for row in data_rows:
            if not row:
                continue
            if (row[0] or "").strip() != today:
                continue
            naver = self._parse_order_index_cell(row[1] if len(row) > 1 else None)
            coupang = self._parse_order_index_cell(row[2] if len(row) > 2 else None)
            gmarket = self._parse_order_index_cell(row[3] if len(row) > 3 else None)
            return {"naver": naver, "coupang": coupang, "gmarket": gmarket}
        return None

    def _write_order_indices_to_sheet(self):
        ws = self._get_order_index_worksheet()
        today = date.today().strftime("%Y-%m-%d")
        n, c, g = self.current_idx_naver, self.current_idx_coupang, self.current_idx_gmarket
        values = ws.get_all_values()
        if not values:
            ws.update(
                [ORDER_INDEX_SHEET_HEADERS, [today, n, c, g]],
                range_name="A1:D2",
            )
            return

        header_like = self._sheet_row_looks_like_order_index_header(values[0])
        if header_like:
            search_ranges = list(enumerate(values[1:], start=2))
        else:
            search_ranges = list(enumerate(values, start=1))

        row_1based = None
        for ridx, row in search_ranges:
            if row and (row[0] or "").strip() == today:
                row_1based = ridx
                break

        if row_1based is not None:
            ws.update([[n, c, g]], range_name=f"B{row_1based}:D{row_1based}")
        else:
            ws.append_row([today, n, c, g])

    def _apply_order_indices_to_ui(self, naver, coupang, gmarket):
        self.current_idx_naver = naver
        self.current_idx_coupang = coupang
        self.current_idx_gmarket = gmarket
        self._block_index_line_edit_signals(True)
        try:
            if hasattr(self.ui, "lineEdit_idx_naver"):
                self.ui.lineEdit_idx_naver.setText(str(naver))
            if hasattr(self.ui, "lineEdit_idx_coupang"):
                self.ui.lineEdit_idx_coupang.setText(str(coupang))
            if hasattr(self.ui, "lineEdit_idx_gmarket"):
                self.ui.lineEdit_idx_gmarket.setText(str(gmarket))
        finally:
            self._block_index_line_edit_signals(False)

    def _schedule_order_index_sheet_push(self):
        self._index_sheet_push_timer.start(ORDER_INDEX_SHEET_PUSH_DEBOUNCE_MS)
        self._update_index_sheet_sync_label()

    def _flush_order_indices_to_sheet(self):
        if gspread is None:
            self._index_sheet_push_pending = True
            self._update_index_sheet_sync_label()
            return
        try:
            self._write_order_indices_to_sheet()
            self._index_sheet_push_pending = False
            self._index_sheet_last_sync_display = datetime.now()
        except Exception as e:
            self._index_sheet_push_pending = True
            print(f"! 스프레드시트 인덱스 반영 실패: {e}")
        self._update_index_sheet_sync_label()

    def _update_index_sheet_sync_label(self):
        if not hasattr(self.ui, "label_index_sheet_sync"):
            return
        if self._index_sheet_push_timer.isActive():
            self.ui.label_index_sheet_sync.setText(
                "저장 대기: 스프레드시트에 곧 반영됩니다…"
            )
            return
        if self._index_sheet_push_pending:
            self.ui.label_index_sheet_sync.setText(
                "저장 대기: 스프레드시트에 반영하지 못했습니다. "
                "「Google Sheets 연동」 재인증 또는 네트워크를 확인해 주세요."
            )
            return
        if self._index_sheet_last_sync_display:
            t = self._index_sheet_last_sync_display.strftime("%H:%M:%S")
            self.ui.label_index_sheet_sync.setText(
                f"스프레드시트와 동기화됨 (마지막 반영 {t})"
            )
        else:
            self.ui.label_index_sheet_sync.setText(
                "스프레드시트와 인덱스를 맞추려면 Google 로그인 후 자동으로 반영됩니다."
            )

    def _create_today_order_index_row_with_ones(self, interactive=False):
        """시트에 오늘 행이 없을 때 네이버·쿠팡·지마켓을 각 1로 새 행을 만듭니다."""
        self._last_refresh_created_today_row = True
        self._apply_order_indices_to_ui(1, 1, 1)
        try:
            self._persist_index_values_to_json()
        except Exception as e:
            print(f"! 인덱스 JSON 저장 실패(시트 신규 행): {e}")

        if gspread is None:
            self._index_sheet_push_pending = True
            self._update_index_sheet_sync_label()
            if interactive:
                QMessageBox.warning(
                    self,
                    "인덱스 동기화",
                    "gspread 패키지가 없어 스프레드시트에 행을 만들 수 없습니다.",
                )
            return False

        try:
            self._write_order_indices_to_sheet()
            self._index_sheet_push_pending = False
            self._index_sheet_last_sync_display = datetime.now()
        except Exception as e:
            self._index_sheet_push_pending = True
            print(f"! 스프레드시트에 오늘 인덱스 행 생성 실패: {e}")
            if interactive:
                hint = self._oauth_error_dialog_hint()
                QMessageBox.warning(
                    self,
                    "인덱스 동기화",
                    f"오늘 날짜 행을 스프레드시트에 만들지 못했습니다.\n\n{e}\n\n{hint}",
                )
            self._update_index_sheet_sync_label()
            return False

        self._update_index_sheet_sync_label()
        return True

    def refresh_order_indices_from_sheet(self, interactive=False):
        """스프레드시트에서 오늘 날짜 행을 읽어 UI·JSON에 반영합니다."""
        self._index_sheet_push_timer.stop()
        self._last_refresh_created_today_row = False
        try:
            row = self._read_today_order_indices_from_sheet()
        except Exception as e:
            if interactive:
                hint = self._oauth_error_dialog_hint()
                QMessageBox.warning(
                    self,
                    "인덱스 동기화",
                    f"스프레드시트에서 인덱스를 읽지 못했습니다.\n\n{e}\n\n{hint}",
                )
            print(f"! 시트에서 인덱스 읽기 실패: {e}")
            self._update_index_sheet_sync_label()
            return False

        if row is None:
            return self._create_today_order_index_row_with_ones(interactive=interactive)

        na, co, gm = row["naver"], row["coupang"], row["gmarket"]
        if (
            self.current_idx_naver == na
            and self.current_idx_coupang == co
            and self.current_idx_gmarket == gm
        ):
            self._index_sheet_push_pending = False
            self._index_sheet_last_sync_display = datetime.now()
            self._update_index_sheet_sync_label()
            return True

        self._apply_order_indices_to_ui(na, co, gm)
        try:
            self._persist_index_values_to_json()
        except Exception as e:
            print(f"! 인덱스 JSON 저장 실패(시트에서 읽은 뒤): {e}")
        self._index_sheet_push_pending = False
        self._index_sheet_last_sync_display = datetime.now()
        self._update_index_sheet_sync_label()
        return True

    def _on_order_index_sheet_poll(self):
        if not self._startup_order_index_sync_done:
            return
        if self._index_sheet_push_timer.isActive():
            return
        for name in ("lineEdit_idx_naver", "lineEdit_idx_coupang", "lineEdit_idx_gmarket"):
            w = getattr(self.ui, name, None)
            if w is not None and w.hasFocus():
                return
        self.refresh_order_indices_from_sheet(interactive=False)

    def _on_push_button_index_sheet_refresh_clicked(self):
        applied = self.refresh_order_indices_from_sheet(interactive=True)
        if not applied:
            return
        if self._last_refresh_created_today_row:
            QMessageBox.information(
                self,
                "인덱스 동기화",
                "오늘 날짜 행이 없어 네이버·쿠팡·지마켓 인덱스를 각 1로 "
                "스프레드시트에 새로 만들었습니다.",
            )
        else:
            QMessageBox.information(
                self,
                "인덱스 동기화",
                "스프레드시트에서 인덱스를 불러왔습니다.",
            )

    def update_naver_index(self):
        """네이버 인덱스 값을 업데이트하고 저장합니다."""
        self.current_idx_naver += 1
        if hasattr(self.ui, 'lineEdit_idx_naver'):
            w = self.ui.lineEdit_idx_naver
            w.blockSignals(True)
            try:
                w.setText(str(self.current_idx_naver))
            finally:
                w.blockSignals(False)
        self.save_index_values()

    def update_coupang_index(self):
        """쿠팡 인덱스 값을 업데이트하고 저장합니다."""
        self.current_idx_coupang += 1
        if hasattr(self.ui, 'lineEdit_idx_coupang'):
            w = self.ui.lineEdit_idx_coupang
            w.blockSignals(True)
            try:
                w.setText(str(self.current_idx_coupang))
            finally:
                w.blockSignals(False)
        self.save_index_values()

    def update_gmarket_index(self):
        """지마켓 인덱스 값을 업데이트하고 저장합니다."""
        self.current_idx_gmarket += 1
        if hasattr(self.ui, 'lineEdit_idx_gmarket'):
            w = self.ui.lineEdit_idx_gmarket
            w.blockSignals(True)
            try:
                w.setText(str(self.current_idx_gmarket))
            finally:
                w.blockSignals(False)
        self.save_index_values()

    def load_ui(self):
        """UI 파일을 로드합니다."""
        ui_file_path = Path(__file__).parent / "ui" / "main_window.ui"
        print(f"UI 파일 경로: {ui_file_path}")
        ui_file = QFile(str(ui_file_path))
        
        if not ui_file.open(QIODevice.ReadOnly):
            print(f"UI 파일을 열 수 없음: {ui_file_path}")
            QMessageBox.critical(self, "오류", f"UI 파일을 열 수 없습니다: {ui_file_path}")
            sys.exit(1)
            
        loader = QUiLoader()
        window = loader.load(ui_file)
        ui_file.close()
        
        if not window:
            print(f"UI 로드 실패: {loader.errorString()}")
            QMessageBox.critical(self, "오류", f"UI 파일을 로드할 수 없습니다: {loader.errorString()}")
            sys.exit(1)
            
        print("UI 로드 완료")
        
        # UI 객체를 인스턴스 변수로 저장
        self.ui = window
        
        # 메인 윈도우 설정
        self.setCentralWidget(window.centralwidget)
        # self.setMenuBar(window.menubar)
        
        # 툴바 설정
        self.setup_toolbar()
        
        self.setStatusBar(window.statusbar)
        self.setWindowTitle(window.windowTitle())
        self.resize(window.size())
        self._order_ship_splitter_initialized = False

    def showEvent(self, event):
        super().showEvent(event)
        if not self._order_ship_splitter_initialized:
            QTimer.singleShot(0, self._on_first_show_splitter)
        if not self._startup_order_index_sync_started:
            self._startup_order_index_sync_started = True
            QTimer.singleShot(0, self._kickoff_startup_order_index_sync)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        ov = getattr(self, "_startup_overlay", None)
        if ov is not None and ov.isVisible():
            self._layout_startup_overlay_full()
        bov = getattr(self, "_busy_overlay", None)
        if bov is not None and bov.isVisible():
            self._layout_busy_overlay_full()

    def _layout_startup_overlay_full(self):
        if getattr(self, "_startup_overlay", None) is None:
            return
        self._startup_overlay.setGeometry(
            0, 0, max(1, self.width()), max(1, self.height())
        )

    def _setup_startup_loading_overlay(self):
        accent = "#21838a"
        accent_soft = "#e8f4f5"

        self._startup_overlay = QFrame(self)
        self._startup_overlay.setObjectName("startupLoadingOverlay")
        self._startup_overlay.hide()
        outer = QVBoxLayout(self._startup_overlay)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addStretch(1)

        row = QHBoxLayout()
        row.addStretch(1)

        card = QFrame()
        card.setObjectName("startupLoadingCard")
        card.setMaximumWidth(480)
        card.setMinimumWidth(320)

        self._startup_card_opacity_effect = QGraphicsOpacityEffect(card)
        self._startup_card_opacity_effect.setOpacity(1.0)
        card.setGraphicsEffect(self._startup_card_opacity_effect)

        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(32, 28, 32, 28)
        card_lay.setSpacing(12)

        badge = QLabel("Easy Fulfill")
        badge.setAlignment(Qt.AlignCenter)
        bf = QFont()
        bf.setFamilies(
            ["Bahnschrift", "맑은 고딕", "Malgun Gothic", "Segoe UI Semibold", "sans-serif"]
        )
        bf.setPixelSize(16)
        bf.setWeight(QFont.Weight.DemiBold)
        bf.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 2.2)
        badge.setFont(bf)
        badge.setStyleSheet(f"color: {accent}; margin-bottom: 4px;")

        title = QLabel("초기 데이터를 불러오는 중…")
        title.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setFamilies(
            ["Bahnschrift", "맑은 고딕", "Malgun Gothic", "Segoe UI Variable Display", "sans-serif"]
        )
        title_font.setPixelSize(24)
        title_font.setWeight(QFont.Weight.DemiBold)
        title_font.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.0)
        title.setFont(title_font)
        title.setStyleSheet("color: #1a202c;")

        sub = QLabel("Google 시트와 주문 인덱스를 동기화합니다.")
        sub.setAlignment(Qt.AlignCenter)
        sub_font = QFont(title_font)
        sub_font.setPixelSize(14)
        sub_font.setWeight(QFont.Weight.Normal)
        sub_font.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 0.25)
        sub.setFont(sub_font)
        sub.setStyleSheet("color: #5c6b7a;")
        sub.setWordWrap(True)

        self._startup_progress_bar = QProgressBar()
        self._startup_progress_bar.setRange(0, 100)
        self._startup_progress_bar.setValue(0)
        self._startup_progress_bar.setTextVisible(True)
        self._startup_progress_bar.setFormat("%p%")
        self._startup_progress_bar.setFixedSize(300, 28)
        self._startup_progress_bar.setStyleSheet(
            f"""
            QProgressBar {{
                border: 1px solid #c5d0d8;
                border-radius: 12px;
                background-color: {accent_soft};
                text-align: center;
                color: #2c3e50;
                font-size: 11px;
                font-weight: 600;
            }}
            QProgressBar::chunk {{
                background-color: {accent};
                border-radius: 11px;
            }}
            """
        )
        self._startup_progress_timer = QTimer(self)
        self._startup_progress_timer.setInterval(STARTUP_SYNC_PROGRESS_TICK_MS)
        self._startup_progress_timer.timeout.connect(self._on_startup_progress_tick)
        self._startup_sync_progress_done = False

        bar_row = QHBoxLayout()
        bar_row.addStretch(1)
        bar_row.addWidget(self._startup_progress_bar)
        bar_row.addStretch(1)

        hint = QLabel("환경에 따라 수 초 걸릴 수 있습니다.")
        hint.setAlignment(Qt.AlignCenter)
        hf = QFont(sub_font)
        hf.setPixelSize(12)
        hint.setFont(hf)
        hint.setStyleSheet("color: #8896a3; margin-top: 4px;")
        hint.setWordWrap(True)

        card_lay.addWidget(badge)
        card_lay.addWidget(title)
        card_lay.addWidget(sub)
        card_lay.addSpacing(8)
        card_lay.addLayout(bar_row)
        card_lay.addWidget(hint)

        row.addWidget(card, alignment=Qt.AlignmentFlag.AlignHCenter)
        row.addStretch(1)
        outer.addLayout(row)
        outer.addStretch(1)

        self._startup_overlay.setStyleSheet(
            """
            QFrame#startupLoadingOverlay {
                background-color: rgba(28, 34, 42, 0.52);
                border: none;
            }
            QFrame#startupLoadingCard {
                background-color: #fbfbfc;
                border: 1px solid rgba(200, 210, 220, 0.95);
                border-radius: 18px;
            }
            """
        )
        self._startup_card_entrance_anim = None

    def _setup_busy_processing_overlay(self):
        accent = "#21838a"
        self._busy_overlay = QFrame(self)
        self._busy_overlay.setObjectName("busyProcessingOverlay")
        self._busy_overlay.hide()
        outer = QVBoxLayout(self._busy_overlay)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addStretch(1)
        row = QHBoxLayout()
        row.addStretch(1)

        card = QFrame()
        card.setObjectName("busyProcessingCard")
        card.setMaximumWidth(440)
        card.setMinimumWidth(280)

        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(28, 24, 28, 24)
        card_lay.setSpacing(14)

        spin_row = QHBoxLayout()
        spin_row.addStretch(1)
        self._busy_spinner = CircularBusySpinner(card, size=56, line_width=4, color=accent)
        spin_row.addWidget(self._busy_spinner)
        spin_row.addStretch(1)

        self._busy_overlay_title = QLabel("처리 중…")
        self._busy_overlay_title.setAlignment(Qt.AlignCenter)
        tf = QFont()
        tf.setFamilies(
            ["Bahnschrift", "맑은 고딕", "Malgun Gothic", "Segoe UI Variable Display", "sans-serif"]
        )
        tf.setPixelSize(18)
        tf.setWeight(QFont.Weight.DemiBold)
        self._busy_overlay_title.setFont(tf)
        self._busy_overlay_title.setStyleSheet("color: #1a202c;")

        self._busy_overlay_sub = QLabel("")
        self._busy_overlay_sub.setAlignment(Qt.AlignCenter)
        sf = QFont(tf)
        sf.setPixelSize(13)
        sf.setWeight(QFont.Weight.Normal)
        self._busy_overlay_sub.setFont(sf)
        self._busy_overlay_sub.setStyleSheet("color: #5c6b7a;")
        self._busy_overlay_sub.setWordWrap(True)

        card_lay.addLayout(spin_row)
        card_lay.addWidget(self._busy_overlay_title)
        card_lay.addWidget(self._busy_overlay_sub)

        row.addWidget(card, alignment=Qt.AlignmentFlag.AlignHCenter)
        row.addStretch(1)
        outer.addLayout(row)
        outer.addStretch(1)

        self._busy_overlay.setStyleSheet(
            """
            QFrame#busyProcessingOverlay {
                background-color: rgba(28, 34, 42, 0.48);
                border: none;
            }
            QFrame#busyProcessingCard {
                background-color: #fbfbfc;
                border: 1px solid rgba(200, 210, 220, 0.95);
                border-radius: 16px;
            }
            """
        )

    def _layout_busy_overlay_full(self):
        if getattr(self, "_busy_overlay", None) is None:
            return
        self._busy_overlay.setGeometry(
            0, 0, max(1, self.width()), max(1, self.height())
        )

    def _show_busy_processing_overlay(self, title: str, subtitle: str = ""):
        if getattr(self, "_busy_overlay", None) is None:
            return
        self._busy_overlay_title.setText(title)
        self._busy_overlay_sub.setText(subtitle)
        self._busy_overlay_sub.setVisible(bool(subtitle and subtitle.strip()))
        self._layout_busy_overlay_full()
        self._busy_overlay.raise_()
        if (
            getattr(self, "_startup_overlay", None) is not None
            and self._startup_overlay.isVisible()
        ):
            self._startup_overlay.raise_()
        self._busy_overlay.show()

    def _hide_busy_processing_overlay(self):
        if getattr(self, "_busy_overlay", None) is not None:
            self._busy_overlay.hide()

    def _run_order_file_processing_with_async_mapping(self, store_type: str):
        self._product_mapping_thread = ProductMappingLoadThread(store_type)
        self._product_mapping_thread.result_ready.connect(
            self._on_product_mapping_for_order_file_ready
        )
        self._product_mapping_thread.finished.connect(self._cleanup_product_mapping_thread)
        self._product_mapping_thread.start()

    def _cleanup_product_mapping_thread(self):
        self._product_mapping_thread = None

    def _on_product_mapping_for_order_file_ready(self, payload: dict):
        try:
            if not payload.get("ok"):
                err = payload.get("error", "")
                print(f"! 스프레드시트 매핑 로드 실패: {err}")
                exc = RuntimeError(str(err))
                if _is_likely_google_sheets_oauth_error(exc):
                    QMessageBox.critical(
                        self,
                        "오류",
                        "Google 스프레드시트 연동 중 오류가 발생했습니다.\n\n"
                        f"{err}\n\n"
                        + self._oauth_error_dialog_hint(),
                    )
                else:
                    QMessageBox.warning(
                        self,
                        "경고",
                        f"스프레드시트 매핑을 불러오지 못했습니다.\n\n{err}",
                    )
                self.is_order_file_valid = False
                return

            st = payload["store_type"]
            self._spreadsheet_product_code_maps[st] = payload["mapping"]
            if st == "coupang":
                self._coupang_option_to_vp_product_no = payload.get("coupang_vp") or {}

            try:
                if st == "naver":
                    self.process_naver_excel_file()
                elif st == "coupang":
                    self.process_coupang_excel_file()
                self.is_order_file_valid = True
            except Exception as e:
                self.is_order_file_valid = False
                print(f"❌ 파일 처리 중 오류 발생: {str(e)}")
                if _is_likely_google_sheets_oauth_error(e):
                    QMessageBox.critical(
                        self,
                        "오류",
                        "Google 스프레드시트 연동 중 오류가 발생했습니다.\n\n"
                        f"{str(e)}\n\n"
                        + self._oauth_error_dialog_hint(),
                    )
                else:
                    QMessageBox.warning(
                        self,
                        "오류",
                        f"파일 처리 중 오류가 발생했습니다: {str(e)}",
                    )
        finally:
            self._hide_busy_processing_overlay()

    def _on_startup_progress_tick(self):
        if getattr(self, "_startup_sync_progress_done", False):
            return
        t0 = getattr(self, "_startup_sync_debug_t0", None)
        if t0 is None:
            return
        elapsed_ms = (time.perf_counter() - t0) * 1000
        if elapsed_ms <= 0:
            return
        pct = min(99, int(elapsed_ms * 100 / STARTUP_SYNC_PROGRESS_CAP_MS))
        pb = getattr(self, "_startup_progress_bar", None)
        if pb is not None and pb.value() < pct:
            pb.setValue(pct)

    def _hide_startup_loading_overlay_deferred(self):
        if getattr(self, "_startup_overlay", None) is not None:
            self._startup_overlay.hide()

    def _start_startup_overlay_entrance_anim(self):
        self._startup_card_opacity_effect.setOpacity(0.0)
        anim = QPropertyAnimation(self._startup_card_opacity_effect, b"opacity", self)
        anim.setDuration(280)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._startup_card_entrance_anim = anim
        anim.start()

    def _set_startup_syncing_indicator(self):
        """시작 동기화 진행 중임을 막지 않는 가벼운 라벨 표시로만 알립니다."""
        if hasattr(self.ui, "label_index_sheet_sync"):
            self.ui.label_index_sheet_sync.setText("Google 시트와 인덱스 동기화 중…")

    def _kickoff_startup_order_index_sync(self):
        # 논블로킹 시작: 더 이상 전체 화면 오버레이로 사용자를 막지 않습니다.
        # 사용자는 즉시 작업할 수 있고, 주문 인덱스 동기화는 백그라운드에서 진행되어
        # 완료되면 _on_startup_order_index_sync_finished 에서 UI에 반영됩니다.
        # (오버레이/진행바 위젯은 _setup_startup_loading_overlay 에서 빌드되지만
        #  여기서 표시하지 않습니다 — 되돌릴 때 show 만 다시 켜면 됩니다.)
        self._startup_sync_progress_done = False
        self._set_startup_syncing_indicator()
        self._startup_sync_thread = OrderIndexStartupSyncThread(self)
        self._startup_sync_thread.result_ready.connect(self._on_startup_order_index_sync_finished)
        self._startup_sync_thread.finished.connect(self._cleanup_startup_sync_thread)
        self._startup_sync_thread.start()

    def _cleanup_startup_sync_thread(self):
        self._startup_sync_thread = None

    def _on_startup_order_index_sync_finished(self, payload: dict):
        # 논블로킹 시작: 결과는 백그라운드 완료 후 라벨·UI에만 조용히 반영합니다.
        self._startup_sync_progress_done = True
        self._startup_order_index_sync_done = True
        if not self._index_sheet_poll_timer.isActive():
            self._index_sheet_poll_timer.start(ORDER_INDEX_SHEET_POLL_MS)

        if not payload.get("ok"):
            err = payload.get("error", "")
            print(f"! 시작 시 DB동기화 실패: {err}")
            if _is_likely_google_sheets_oauth_error(RuntimeError(str(err))):
                self._prompt_google_reauth_for_oauth_error(
                    title="Google 로그인 만료",
                    err_text=str(err),
                )
            self._update_index_sheet_sync_label()
            return

        kind = payload.get("kind")
        if kind == "created":
            self._last_refresh_created_today_row = True
            self._apply_order_indices_to_ui(1, 1, 1)
            try:
                self._persist_index_values_to_json()
            except Exception as e:
                print(f"! 인덱스 JSON 저장 실패(시트 신규 행): {e}")
            self._index_sheet_push_pending = False
            self._index_sheet_last_sync_display = datetime.now()
            self._update_index_sheet_sync_label()
            return

        if kind != "data":
            self._update_index_sheet_sync_label()
            return

        row = payload["row"]
        na, co, gm = row["naver"], row["coupang"], row["gmarket"]
        self._last_refresh_created_today_row = False
        if (
            self.current_idx_naver == na
            and self.current_idx_coupang == co
            and self.current_idx_gmarket == gm
        ):
            self._index_sheet_push_pending = False
            self._index_sheet_last_sync_display = datetime.now()
            self._update_index_sheet_sync_label()
            return
        self._apply_order_indices_to_ui(na, co, gm)
        try:
            self._persist_index_values_to_json()
        except Exception as e:
            print(f"! 인덱스 JSON 저장 실패(시트에서 읽은 뒤): {e}")
        self._index_sheet_push_pending = False
        self._index_sheet_last_sync_display = datetime.now()
        self._update_index_sheet_sync_label()

    def _on_first_show_splitter(self):
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

    def setup_toolbar(self):
        """툴바를 설정합니다."""
        # 툴바 생성
        toolbar = self.addToolBar('툴바')
        toolbar.setIconSize(QSize(32, 32))
        
        # 툴바를 세로로 설정하고 좌측에 배치
        # toolbar.setOrientation(Qt.Vertical)
        # self.addToolBar(Qt.LeftToolBarArea, toolbar)
        
        # 툴바 스타일 설정
        toolbar.setStyleSheet("""
            QToolBar {
                spacing: 5px;
                padding: 5px;
                background-color: #f5f5f5;
                border: none;
            }
            QToolBar QToolButton {
                min-width: 40px;
                min-height: 40px;
                padding: 5px;
                margin: 2px;
                border: 1.5px solid #dcdcdc;
                border-radius: 8px;
                background-color: white;
            }
            QToolBar QToolButton:hover {
                background-color: #e8e8e8;
                border-color: #c0c0c0;
            }
            QToolBar QToolButton:pressed {
                background-color: #d0d0d0;
                border-color: #a0a0a0;
            }
        """)
        
        # 불러오기 액션
        loadAction = QAction(QIcon('image/open-file-icon.png'), '파일 불러오기', self)
        loadAction.setShortcut('Ctrl+O')
        loadAction.setStatusTip('파일 불러오기 (Ctrl+O)')
        loadAction.triggered.connect(self.select_excel_file)
        toolbar.addAction(loadAction)

        # 복사 액션
        copyAction = QAction(QIcon('image/copy-icon.png'), '클립보드에 복사', self)
        copyAction.setShortcut('Ctrl+C')
        copyAction.setStatusTip('클립보드에 복사 (Ctrl+C)')
        copyAction.triggered.connect(self.copy_to_clipboard)
        toolbar.addAction(copyAction)        

        # 엑셀 송장 출력 액션
        exportAction = QAction(QIcon('image/microsoft-excel-icon.png'), '엑셀 송장 출력', self)
        exportAction.setShortcut('Ctrl+E')
        exportAction.setStatusTip('엑셀 송장 출력 (Ctrl+E)')
        exportAction.triggered.connect(self.export_invoice_excel)
        toolbar.addAction(exportAction)

        # 폴더 열기 액션
        openAction = QAction(QIcon('image/folder-icon.png'), '출력 폴더 열기', self)        
        openAction.setShortcut('Ctrl+F')
        openAction.setStatusTip('출력 폴더 열기 (Ctrl+F)')
        openAction.triggered.connect(self.open_output_folder)
        toolbar.addAction(openAction)                                
        
        # 구분자 추가
        toolbar.addSeparator()
        
        # # 노션 홈페이지 액션
        # notionAction = QAction(QIcon('image/notion-icon.png'), '노션 홈페이지로 이동', self)        
        # notionAction.setStatusTip('노션 홈페이지로 이동')
        # notionAction.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://www.notion.so")))
        # toolbar.addAction(notionAction)
        
        # 데이터베이스 액션
        databaseAction = QAction(QIcon('image/database-icon.png'), '데이터베이스 확인', self)
        databaseAction.setShortcut('Ctrl+D')
        databaseAction.setStatusTip('데이터베이스 확인 (Ctrl+D)')
        databaseAction.triggered.connect(lambda: QDesktopServices.
                                         openUrl(QUrl("https://docs.google.com/spreadsheets/d/1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw/edit?gid=195401368#gid=195401368")))
        toolbar.addAction(databaseAction)  
        
        # 우체국 홈페이지 액션
        notionAction = QAction(QIcon('image/korea-post-icon.png'), '우체국 홈페이지로 이동', self)
        notionAction.setStatusTip('우체국 홈페이지로 이동')
        notionAction.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://biz.epost.go.kr/ui/index.jsp")))
        toolbar.addAction(notionAction)
        
        # 네이버 스마트스토어 액션
        naverStoreAction = QAction(QIcon('image/smart-store-icon.png'), '네이버 스마트스토어 페이지로 이동', self)        
        naverStoreAction.setStatusTip('네이버 스마트스토어 페이지로 이동')
        naverStoreAction.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://sell.smartstore.naver.com/#/home/about")))
        toolbar.addAction(naverStoreAction)
        
        # 쿠팡 스토어 액션
        coupangStoreAction = QAction(QIcon('image/coupang-wing-icon.png'), '쿠팡 스토어 페이지로 이동', self)        
        coupangStoreAction.setStatusTip('쿠팡 스토어 페이지로 이동')
        coupangStoreAction.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://wing.coupang.com/")))
        toolbar.addAction(coupangStoreAction)
        
         # 구분자 추가
        toolbar.addSeparator()
        
        # 초기화 액션
        clearAction = QAction(QIcon('image/reset-icon.png'), '주문 정보 초기화', self)
        clearAction.setShortcut('Ctrl+R')
        clearAction.setStatusTip('주문 정보 초기화 (Ctrl+R)')
        clearAction.triggered.connect(self.clear_list)
        toolbar.addAction(clearAction)  
        
        # 종료 액션
        exitAction = QAction(QIcon('image/exit-icon.png'), '프로그램 종료', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('프로그램 종료 (Ctrl+Q)')
        exitAction.triggered.connect(QApplication.quit)
        toolbar.addAction(exitAction)
        
        

    def copy_to_clipboard(self):
        """plainTextEdit의 내용을 클립보드에 복사합니다."""
        try:
            if hasattr(self.ui, 'plainTextEdit'):
                text = self.ui.plainTextEdit.toPlainText()
                clipboard = QApplication.clipboard()
                clipboard.setText(text)
                self.statusBar().showMessage("텍스트가 클립보드에 복사되었습니다.", 2000)
            else:
                QMessageBox.warning(self, "오류", "plainTextEdit 위젯이 존재하지 않습니다.")
        except Exception as e:
            QMessageBox.warning(self, "오류", f"복사 중 오류가 발생했습니다: {str(e)}")
        
    def setup_status_bar(self):
        """상태바를 초기화하고 기본 메시지를 설정합니다."""
        statusbar = self.statusBar()
        
        # 기본 레이블 생성 및 설정
        self.status_label = QLabel()
        self.status_label.setAlignment(Qt.AlignCenter)  # 중앙 정렬
        self.status_label.setMinimumHeight(30)  # 최소 높이 설정
        self.status_label.setStyleSheet("""
            QLabel {
                color: #333333;
                font-size: 12pt;
                font-weight: bold;
                font-family: Arial;
                padding: 0 10px;
                background-color: #f5f5f5;
                width: 100%;
            }
        """)
        
        # 상태바에 레이블을 추가하고 너비를 최대로 설정
        statusbar.removeWidget(self.status_label)  # 기존 위젯 제거
        statusbar.addWidget(self.status_label, 1)  # stretch factor 1로 설정하여 최대 너비 사용
        
        # 상태바 스타일 설정
        statusbar.setStyleSheet("""
            QStatusBar {
                background-color: #f5f5f5;
                min-height: 35px;
                border-top: 1px solid #dcdcdc;
            }
            QStatusBar::item {
                border: none;
                width: 100%;
            }
        """)
        
        def update_status_label(message):
            """상태바 메시지 업데이트 시 레이블도 함께 업데이트"""
            if message:  # 메시지가 있을 때만 업데이트
                self.status_label.setText(message)
                statusbar.clearMessage()  # 기본 메시지 클리어
        
        # 상태바 메시지 변경 시그널 연결
        statusbar.messageChanged.connect(update_status_label)
        statusbar.showMessage("준비")  # 초기 메시지 설정
        
    def setup_connections(self):
        """버튼과 메뉴 동작을 연결합니다."""        
        # 주문·발송 탭 버튼 연결
        self.ui.pushButton_load_order.clicked.connect(self.select_excel_file)
        self.ui.pushButton_load_invoice.clicked.connect(self.load_invoice_file)
        self.ui.pushButton_generate_invoice.clicked.connect(self.generate_invoice_file)
        
        # 환경설정 탭 버튼 연결
        # 사은품 UI(groupBox_2): main_window.ui에서 제거함. 복구 시 위젯·로직 재연결.
        # 신규 DB 반영 UI·로직 제거(2026-04). (pushButton_database_* / load·apply_database_* 삭제)
        if hasattr(self.ui, 'pushButton_quick_excel_gen'):
            self.ui.pushButton_quick_excel_gen.clicked.connect(self.generate_quick_excel)
        if hasattr(self.ui, 'comboBox_store_select'):
            self.ui.comboBox_store_select.currentIndexChanged.connect(
                self._refresh_quick_excel_manual_ui
            )
            self._refresh_quick_excel_manual_ui()
        if hasattr(self.ui, "pushButton_quick_manual_postcode_lookup"):
            self.ui.pushButton_quick_manual_postcode_lookup.clicked.connect(
                self._on_quick_manual_postcode_lookup_clicked
            )
        if hasattr(self.ui, "pushButton_quick_manual_address_search"):
            self.ui.pushButton_quick_manual_address_search.clicked.connect(
                self._on_quick_manual_address_search_clicked
            )

        # 인덱스 입력 필드 연결
        if hasattr(self.ui, 'lineEdit_idx_naver'):
            self.ui.lineEdit_idx_naver.textChanged.connect(self.on_naver_index_changed)
        if hasattr(self.ui, 'lineEdit_idx_coupang'):
            self.ui.lineEdit_idx_coupang.textChanged.connect(self.on_coupang_index_changed)
        if hasattr(self.ui, 'lineEdit_idx_gmarket'):
            self.ui.lineEdit_idx_gmarket.textChanged.connect(self.on_gmarket_index_changed)
        
        if hasattr(self.ui, 'checkBox_invoice_load_auto_generate'):
            self.ui.checkBox_invoice_load_auto_generate.stateChanged.connect(
                self.save_app_settings
            )
        
        # 메뉴 동작 연결
        self.ui.actionOpenExcel.triggered.connect(self.select_excel_file)
        self.ui.actionExit.triggered.connect(self.close)
        self.ui.actionAbout.triggered.connect(self.show_about)

        if hasattr(self.ui, "pushButton_google_reauth"):
            self.ui.pushButton_google_reauth.clicked.connect(self._on_google_reauth_clicked)
        if hasattr(self.ui, "pushButton_google_disconnect"):
            self.ui.pushButton_google_disconnect.clicked.connect(self._on_google_disconnect_clicked)
        if hasattr(self.ui, "pushButton_google_open_oauth_folder"):
            self.ui.pushButton_google_open_oauth_folder.clicked.connect(
                self._on_google_open_oauth_folder_clicked
            )
        if hasattr(self.ui, "tabWidget"):
            self.ui.tabWidget.currentChanged.connect(self._on_main_tab_changed)
        if hasattr(self.ui, "pushButton_index_sheet_refresh"):
            self.ui.pushButton_index_sheet_refresh.clicked.connect(
                self._on_push_button_index_sheet_refresh_clicked
            )
        if hasattr(self.ui, "pushButton_db_sync_refresh_paths"):
            self.ui.pushButton_db_sync_refresh_paths.clicked.connect(
                self._on_db_sync_refresh_paths_clicked
            )
        if hasattr(self.ui, "pushButton_db_sync_naver_browse"):
            self.ui.pushButton_db_sync_naver_browse.clicked.connect(
                self._on_db_sync_naver_browse_clicked
            )
        if hasattr(self.ui, "pushButton_db_sync_coupang_browse"):
            self.ui.pushButton_db_sync_coupang_browse.clicked.connect(
                self._on_db_sync_coupang_browse_clicked
            )
        if hasattr(self.ui, "pushButton_db_sync_run"):
            self.ui.pushButton_db_sync_run.clicked.connect(self._on_db_sheet_sync_run_clicked)

        print("버튼과 메뉴 연결 완료")

    def _invalidate_google_sheets_client_and_caches(self):
        self._gspread_client = None
        self._spreadsheet_product_code_maps = {}
        self._coupang_option_to_vp_product_no = {}

    def _refresh_google_auth_status_ui(self):
        if not hasattr(self.ui, "label_google_auth_status"):
            return
        try:
            from google_sheets_oauth import get_oauth_status_description

            self.ui.label_google_auth_status.setText(get_oauth_status_description())
        except Exception as e:
            self.ui.label_google_auth_status.setText(f"상태를 표시할 수 없습니다.\n{e}")

    def _on_main_tab_changed(self, index):
        if index == 1:
            self._refresh_google_auth_status_ui()
        if index == 2:
            self._refresh_db_sheet_sync_path_labels()

    def _on_db_sync_refresh_paths_clicked(self):
        self._refresh_db_sheet_sync_path_labels(reset_manual_overrides=True)

    def _db_sync_start_dir(self, config_db_dir: str) -> str:
        p = Path(config_db_dir).resolve()
        return str(p) if p.is_dir() else str(Path.cwd())

    def _on_db_sync_naver_browse_clicked(self):
        if not hasattr(self.ui, "label_db_sync_naver_file"):
            return
        try:
            from db_sheet_sync import NAVER_CONFIG
        except ImportError:
            QMessageBox.warning(self, "DB동기화", "db_sheet_sync 모듈을 불러올 수 없습니다.")
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "네이버 상품 DB (CSV) 선택",
            self._db_sync_start_dir(NAVER_CONFIG["db_dir"]),
            "CSV (*.csv);;모든 파일 (*.*)",
        )
        if not path:
            return
        self._db_sync_naver_path_override = Path(path)
        self._refresh_db_sheet_sync_path_labels()

    def _on_db_sync_coupang_browse_clicked(self):
        if not hasattr(self.ui, "label_db_sync_coupang_file"):
            return
        try:
            from db_sheet_sync import COUPANG_CONFIG
        except ImportError:
            QMessageBox.warning(self, "DB동기화", "db_sheet_sync 모듈을 불러올 수 없습니다.")
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "쿠팡 가격·재고 DB (Excel) 선택",
            self._db_sync_start_dir(COUPANG_CONFIG["db_dir"]),
            "Excel (*.xlsx);;모든 파일 (*.*)",
        )
        if not path:
            return
        self._db_sync_coupang_path_override = Path(path)
        self._refresh_db_sheet_sync_path_labels()

    def _set_db_sync_path_controls_busy(self, busy: bool):
        w = self.ui
        for name in (
            "pushButton_db_sync_refresh_paths",
            "pushButton_db_sync_run",
            "pushButton_db_sync_naver_browse",
            "pushButton_db_sync_coupang_browse",
        ):
            b = getattr(w, name, None)
            if b is not None:
                b.setEnabled(not busy)

    def _refresh_db_sheet_sync_path_labels(self, reset_manual_overrides=False):
        if not hasattr(self.ui, "label_db_sync_naver_file"):
            return
        try:
            from db_sheet_sync import (
                COUPANG_CONFIG,
                NAVER_CONFIG,
                format_db_sync_label_line,
                get_latest_file_from_pattern,
                get_latest_file_from_patterns,
            )
        except ImportError as e:
            self.ui.label_db_sync_naver_file.setText(f"네이버 DB [ — ] : (모듈 오류: {e})")
            self.ui.label_db_sync_coupang_file.setText("쿠팡 DB [ — ] : —")
            return
        if reset_manual_overrides:
            self._db_sync_naver_path_override = None
            self._db_sync_coupang_path_override = None
        n = self._db_sync_naver_path_override
        if n is None:
            n = get_latest_file_from_patterns(NAVER_CONFIG["db_dir"], NAVER_CONFIG["file_patterns"])
        c = self._db_sync_coupang_path_override
        if c is None:
            c = get_latest_file_from_pattern(COUPANG_CONFIG["db_dir"], COUPANG_CONFIG["file_pattern"])
        self.ui.label_db_sync_naver_file.setText(format_db_sync_label_line("네이버 DB", n))
        self.ui.label_db_sync_coupang_file.setText(format_db_sync_label_line("쿠팡 DB", c))

    def _on_db_sheet_sync_run_clicked(self):
        if not hasattr(self.ui, "plainTextEdit_db_sync_log"):
            return
        if self._db_sheet_sync_thread is not None and self._db_sheet_sync_thread.isRunning():
            QMessageBox.information(self, "DB동기화", "이미 실행 중입니다.")
            return
        if gspread is None:
            QMessageBox.warning(self, "DB동기화", "gspread 패키지가 필요합니다. (pip install gspread)")
            return
        do_naver = self.ui.checkBox_db_sync_naver.isChecked()
        do_coupang = self.ui.checkBox_db_sync_coupang.isChecked()
        if not do_naver and not do_coupang:
            QMessageBox.warning(self, "DB동기화", "네이버 또는 쿠팡 중 하나 이상 선택하세요.")
            return
        test_mode = self.ui.checkBox_db_sync_preview.isChecked()
        test_count = 0
        verbose_log = (
            hasattr(self.ui, "checkBox_db_sync_verbose_log")
            and self.ui.checkBox_db_sync_verbose_log.isChecked()
        )
        if not test_mode:
            reply = QMessageBox.question(
                self,
                "DB동기화",
                "테스트 모드가 꺼져 있습니다. 스프레드시트에 신규 행을 실제로 추가합니다. 계속할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        self._show_busy_processing_overlay(
            "DB 동기화 작업 중…",
            "시트와 로컬 DB 파일을 처리하고 있습니다. 잠시만 기다려 주세요.",
        )
        QApplication.processEvents()
        self._set_db_sync_path_controls_busy(True)
        self.ui.plainTextEdit_db_sync_log.setPlainText("실행 중… 잠시만 기다려 주세요.\n")
        params = {
            "spreadsheet_id": SPREADSHEET_ID,
            "do_naver": do_naver,
            "do_coupang": do_coupang,
            "test_mode": test_mode,
            "test_count": test_count,
            "naver_path": (
                str(self._db_sync_naver_path_override)
                if self._db_sync_naver_path_override is not None
                else None
            ),
            "coupang_path": (
                str(self._db_sync_coupang_path_override)
                if self._db_sync_coupang_path_override is not None
                else None
            ),
            "verbose_log": verbose_log,
        }
        self._db_sheet_sync_thread = DbSheetSyncThread(params)
        self._db_sheet_sync_thread.result_ready.connect(self._on_db_sheet_sync_finished)
        self._db_sheet_sync_thread.start()

    def _on_db_sheet_sync_finished(self, result: dict):
        self._hide_busy_processing_overlay()
        self._set_db_sync_path_controls_busy(False)
        self._db_sheet_sync_thread = None
        lines = result.get("logs") or []
        self.ui.plainTextEdit_db_sync_log.setPlainText("\n".join(lines))
        if result.get("error"):
            err_text = str(result["error"])
            if _is_likely_google_sheets_oauth_error(RuntimeError(err_text)):
                handled = self._prompt_google_reauth_for_oauth_error(
                    title="DB동기화",
                    err_text=err_text,
                )
                if not handled:
                    QMessageBox.warning(self, "DB동기화", err_text)
            else:
                QMessageBox.warning(self, "DB동기화", err_text)
        wrote = False
        for key in ("naver", "coupang"):
            ch = result.get(key)
            if ch and ch.get("appended"):
                wrote = True
        if not result.get("test_mode") and wrote and result.get("ok"):
            self._invalidate_google_sheets_client_and_caches()

        done_ch = []
        skipped_ch = []
        for key, label in (("naver", "네이버"), ("coupang", "쿠팡")):
            ch = result.get(key)
            if not ch:
                continue
            if ch.get("skipped_missing_file"):
                skipped_ch.append(label)
            elif ch.get("ok"):
                done_ch.append(label)

        if result.get("error"):
            self.statusBar().showMessage("DB동기화 중 오류가 있습니다.", 8000)
        elif result.get("ok"):
            if skipped_ch and done_ch:
                self.statusBar().showMessage(
                    "DB동기화 완료 · 일부 채널은 로컬 파일 없음으로 건너뜀", 7000
                )
            elif skipped_ch and not done_ch:
                self.statusBar().showMessage(
                    "DB동기화: 선택한 채널에 해당하는 로컬 파일이 없습니다", 7000
                )
            else:
                self.statusBar().showMessage("DB동기화 완료", 5000)
        else:
            self.statusBar().showMessage("DB동기화 중 오류가 있습니다.", 8000)

    def _on_google_disconnect_clicked(self):
        reply = QMessageBox.question(
            self,
            "연결 해제",
            "이 PC에 저장된 Google 로그인 정보(token.json)를 삭제할까요?\n"
            "다음에 스프레드시트 연동 시 다시 브라우저 로그인이 필요합니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        from google_sheets_oauth import delete_oauth_token_file

        deleted = delete_oauth_token_file()
        self._invalidate_google_sheets_client_and_caches()
        self._refresh_google_auth_status_ui()
        if deleted:
            QMessageBox.information(self, "연결 해제", "저장된 로그인 정보를 삭제했습니다.")
        else:
            QMessageBox.information(self, "연결 해제", "삭제할 token.json이 없었습니다.")

    def _on_google_reauth_clicked(self, ask_confirm: bool = True):
        if ask_confirm:
            reply = QMessageBox.question(
                self,
                "재인증",
                "브라우저가 열리면 Google 계정으로 로그인·권한 승인을 완료해 주세요.\n"
                "진행 시 기존 token.json은 삭제됩니다.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return False

        if gspread is None:
            QMessageBox.warning(self, "재인증", "gspread 패키지가 필요합니다. (pip install gspread)")
            return False
        try:
            from google_sheets_oauth import delete_oauth_token_file, get_authorized_gspread_client
        except ImportError:
            QMessageBox.warning(self, "재인증", "google-auth-oauthlib 패키지가 필요합니다.")
            return False

        delete_oauth_token_file()
        self._invalidate_google_sheets_client_and_caches()
        try:
            self._gspread_client = get_authorized_gspread_client()
        except Exception as e:
            self._gspread_client = None
            QMessageBox.critical(
                self,
                "재인증 실패",
                f"Google 로그인에 실패했습니다.\n\n{e}",
            )
            self._refresh_google_auth_status_ui()
            return False
        self._refresh_google_auth_status_ui()
        QMessageBox.information(self, "재인증", "Google Sheets 연동이 완료되었습니다.")
        return True

    def _on_google_open_oauth_folder_clicked(self):
        from google_sheets_oauth import GOOGLE_AUTH_DIR

        GOOGLE_AUTH_DIR.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(GOOGLE_AUTH_DIR.resolve())))

    @staticmethod
    def _oauth_error_dialog_hint() -> str:
        return (
            "「환경설정」탭 → 「Google Sheets 연동」에서 "
            "「연결 해제」 후 「재인증」을 눌러 보세요."
        )

    def _prompt_google_reauth_for_oauth_error(self, *, title: str, err_text: str) -> bool:
        reply = QMessageBox.question(
            self,
            title,
            "Google Sheets 로그인 토큰이 만료되었거나 철회되었습니다.\n\n"
            f"{err_text}\n\n"
            "지금 저장된 토큰을 삭제하고 재인증하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return False
        return bool(self._on_google_reauth_clicked(ask_confirm=False))

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
        """우체국 KpostPortal 우편번호 조회 regkey."""
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

        if code == KPOST_ERROR_NO_REGKEY:
            QMessageBox.warning(
                parent,
                "우체국 API 인증키가 없습니다",
                "「조회」나「주소 검색」은 우체국 우편번호 오픈API를 사용합니다.\n\n"
                "이 PC에는 아직 API 인증키(regkey)가 등록되어 있지 않습니다.\n\n"
                "등록 방법(택 1):\n"
                "· 환경 변수 KPOST_REGKEY 또는 EPOST_REGKEY 에 발급받은 키 설정\n"
                '· database/app_settings.json 에 "kpost_regkey": "발급키" 추가\n\n'
                "인증키 없이도 이름·기본주소·상세주소·우편번호를 직접 입력한 뒤 "
                "「생성」을 누르면 엑셀은 정상적으로 만들 수 있습니다.",
            )
            return

        QMessageBox.warning(
            parent,
            "우편번호 조회 오류",
            f"오류 코드: {code}\n\n원인:\n{msg}",
        )

    def _kpost_parse_postnew_response(self, text: str):
        """
        Returns:
            (items, None) — items: [{postcd, address, addrjibun}, ...]
            (None, (error_code, message)) — API 오류 XML
        """
        try:
            root = ET.fromstring(text)
        except ET.ParseError:
            return None, ("XML 오류", "응답을 XML로 해석할 수 없습니다.")

        for el in root.iter():
            if _xml_local_lower(el.tag) != "error":
                continue
            code, msg = "", ""
            for ch in el:
                ln = _xml_local_lower(ch.tag)
                if ln == "error_code":
                    code = (ch.text or "").strip()
                elif ln == "message":
                    msg = (ch.text or "").strip()
            return None, (code or "오류", msg or "알 수 없는 오류입니다.")

        items = []
        for el in root.iter():
            if _xml_local_lower(el.tag) != "item":
                continue
            row = {}
            for ch in el:
                ln = _xml_local_lower(ch.tag)
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
        regkey = self._get_kpost_regkey()
        if not regkey:
            return None, (KPOST_ERROR_NO_REGKEY, "")

        q = (query or "").strip()
        if len(q) < 2:
            return None, (
                "ERR-121",
                "검색어는 2자 이상 입력해야 합니다.",
            )

        params = {
            "regkey": regkey,
            "target": "postNew",
            "query": q,
            "countPerPage": 20,
            "currentPage": 1,
        }
        try:
            resp = requests.get(KPOST_OPENAPI2_URL, params=params, timeout=15)
            resp.raise_for_status()
        except requests.RequestException as e:
            return None, ("HTTP 오류", str(e))

        text = resp.content.decode("utf-8", errors="replace")
        return self._kpost_parse_postnew_response(text)

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
        dlg = KpostAddressSearchDialog(self, self)
        if dlg.exec() != QDialog.Accepted:
            return
        if hasattr(self.ui, "lineEdit_quick_manual_address"):
            self.ui.lineEdit_quick_manual_address.setText(dlg.selected_address)
        if hasattr(self.ui, "lineEdit_quick_manual_postcode"):
            self.ui.lineEdit_quick_manual_postcode.setText(dlg.selected_postcd)
        if hasattr(self.ui, "lineEdit_quick_manual_detail"):
            self.ui.lineEdit_quick_manual_detail.setText(dlg.detail_address)

    def _refresh_quick_excel_manual_ui(self):
        """수동 모드일 때만 이름·전화·주소 입력란을 표시하고 그룹 높이를 조정합니다."""
        if not hasattr(self.ui, "comboBox_store_select"):
            return
        manual = self.ui.comboBox_store_select.currentText().strip() == "수동"
        for w in self._iter_quick_excel_manual_widgets():
            w.setVisible(manual)
        if hasattr(self.ui, "label_gift_3"):
            self.ui.label_gift_3.setVisible(not manual)

        gb = getattr(self.ui, "groupBox", None)
        if gb is None:
            return
        quick_h = 278 if manual else 102
        grect = gb.geometry()
        gb.setGeometry(grect.x(), grect.y(), grect.width(), quick_h)

        gap_after_quick = 13
        gs = getattr(self.ui, "groupBox_google_sheets", None)
        bb = getattr(self.ui, "groupBox_batch_ship", None)

        # 우측 컬럼으로 이동된 경우(퀵 엑셀과 x축이 충분히 떨어진 경우),
        # 순서를 Google Sheets -> 일괄 발송으로 고정한다.
        right_column_x = None
        if gs is not None:
            right_column_x = gs.geometry().x()
        elif bb is not None:
            right_column_x = bb.geometry().x()
        is_right_column = (
            right_column_x is not None
            and right_column_x >= (grect.x() + grect.width() + 8)
        )

        if is_right_column:
            top_anchor_y = grect.y()
            order_index_gb = getattr(self.ui, "groupBox_order_index", None)
            if order_index_gb is not None:
                top_anchor_y = order_index_gb.geometry().y()

            gap_between_right = 9
            if gs is not None:
                gsrect = gs.geometry()
                gs.setGeometry(gsrect.x(), top_anchor_y, gsrect.width(), gsrect.height())
            if bb is not None:
                brect = bb.geometry()
                if gs is not None:
                    gsrect_after = gs.geometry()
                    batch_y = gsrect_after.y() + gsrect_after.height() + gap_between_right
                else:
                    batch_y = top_anchor_y
                bb.setGeometry(brect.x(), batch_y, brect.width(), brect.height())
        else:
            if bb is not None:
                brect = bb.geometry()
                batch_y = grect.y() + quick_h + gap_after_quick
                bb.setGeometry(brect.x(), batch_y, brect.width(), brect.height())

            if gs is not None:
                gsrect = gs.geometry()
                gap_after_batch = 9
                if bb is not None:
                    bbr = bb.geometry()
                    google_y = bbr.y() + bbr.height() + gap_after_batch
                else:
                    google_y = grect.y() + quick_h + gap_after_quick + 61 + gap_after_batch
                gs.setGeometry(gsrect.x(), google_y, gsrect.width(), gsrect.height())

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
                    self.statusBar().showMessage(
                        f"퀵 엑셀 생성 완료: {output_file.name}", 3000
                    )
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

                output_file = self.save_invoice_excel(invoice_data, "퀵_쿠팡")
                self.statusBar().showMessage(f"퀵 엑셀 생성 완료: {output_file.name}", 3000)
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

                output_file = self.save_invoice_excel(invoice_data, "퀵_네이버")
                self.statusBar().showMessage(f"퀵 엑셀 생성 완료: {output_file.name}", 3000)
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

                output_file = self.save_invoice_excel(invoice_data, "퀵_지마켓")
                self.statusBar().showMessage(f"퀵 엑셀 생성 완료: {output_file.name}", 3000)
                return

            QMessageBox.information(
                self,
                "안내",
                "현재 퀵 엑셀은 쿠팡, 네이버, 지마켓만 지원합니다.\n"
                "지원되는 양식으로 복사했는지 확인해주세요."
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
        text = clipboard_text.replace("\r", "")
        tokens = []
        for line in text.split("\n"):
            if not line.strip():
                continue
            for token in line.split("\t"):
                cleaned = token.strip()
                if cleaned:
                    tokens.append(cleaned)

        def get_value_after(key):
            try:
                idx = tokens.index(key)
            except ValueError:
                return ""
            if idx + 1 >= len(tokens):
                return ""
            return tokens[idx + 1].strip()

        if "연락처(안심번호)" in tokens or "배송주소" in tokens:
            return "coupang"

        gmarket_address = ""
        if "배송지주소" in tokens:
            gmarket_address = get_value_after("배송지주소")
        if "상품수령인" in tokens or "배송 요청사항" in tokens:
            return "gmarket"
        if gmarket_address and re.match(r"^\d{5}\b", gmarket_address):
            return "gmarket"

        if "연락처1" in tokens or "배송지" in tokens:
            return "naver"

        return None

    def parse_coupang_quick_clipboard(self, clipboard_text):
        """쿠팡 클립보드 텍스트에서 필수 정보를 추출합니다."""
        key_map = {
            "수취인명": "",
            "연락처(안심번호)": "",
            "배송주소": "",
            "배송메모": "",
            "우편번호": ""
        }

        tokens = []
        for line in clipboard_text.replace("\r", "").split("\n"):
            if not line.strip():
                continue
            for token in line.split("\t"):
                cleaned = token.strip()
                if cleaned:
                    tokens.append(cleaned)

        key_set = set(key_map.keys())
        for idx, token in enumerate(tokens):
            if token in key_set and idx + 1 < len(tokens):
                key_map[token] = tokens[idx + 1].strip()

        address = key_map.get("배송주소", "")
        zip_code = key_map.get("우편번호", "")
        if not zip_code:
            zip_code = self.extract_zip_code(address)

        key_map["우편번호"] = zip_code
        return key_map

    def parse_naver_quick_clipboard(self, clipboard_text):
        """네이버 클립보드 텍스트에서 필수 정보를 추출합니다."""
        key_map = {
            "수취인명": "",
            "연락처1": "",
            "연락처2": "",
            "배송지": "",
            "배송메모": ""
        }

        tokens = []
        for line in clipboard_text.replace("\r", "").split("\n"):
            if not line.strip():
                continue
            for token in line.split("\t"):
                cleaned = token.strip()
                if cleaned:
                    tokens.append(cleaned)

        key_set = set(key_map.keys())
        current_key = None
        for token in tokens:
            if token in key_set:
                current_key = token
                continue
            if not current_key:
                continue
            if key_map[current_key]:
                key_map[current_key] = f"{key_map[current_key]} {token}".strip()
            else:
                key_map[current_key] = token

        return key_map

    def parse_gmarket_quick_clipboard(self, clipboard_text):
        """지마켓 클립보드 텍스트에서 필수 정보를 추출합니다."""
        key_map = {
            "상품수령인": "",
            "연락처1": "",
            "연락처2": "",
            "배송지주소": "",
            "배송 요청사항": "",
            "우편번호": ""
        }

        tokens = []
        for line in clipboard_text.replace("\r", "").split("\n"):
            if not line.strip():
                continue
            for token in line.split("\t"):
                cleaned = token.strip()
                if cleaned:
                    tokens.append(cleaned)

        key_set = set(key_map.keys())
        current_key = None
        for token in tokens:
            if token in key_set:
                current_key = token
                continue
            if not current_key:
                continue
            if key_map[current_key]:
                key_map[current_key] = f"{key_map[current_key]} {token}".strip()
            else:
                key_map[current_key] = token

        address = key_map.get("배송지주소", "")
        zip_code = self.extract_zip_code(address)
        key_map["우편번호"] = zip_code
        if zip_code:
            cleaned_address = re.sub(rf"^\s*{re.escape(zip_code)}\s*", "", address).strip()
            key_map["배송지주소"] = cleaned_address

        return key_map

    def extract_zip_code(self, address):
        """주소 문자열에서 우편번호를 추출합니다."""
        if not address:
            return ""
        match = re.search(r"\((\d{5})\)", address)
        if match:
            return match.group(1)
        match = re.search(r"(\d{5})", address)
        if match:
            return match.group(1)
        return ""

    def save_invoice_excel(self, invoice_data, filename_prefix):
        """송장 데이터로 엑셀 파일을 저장하고 경로를 반환합니다."""
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)

        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        output_file = (output_dir / f"{filename_prefix}_{current_time}.xlsx").resolve()

        df_invoice = pd.DataFrame(invoice_data)
        if '배송메세지' in df_invoice.columns:
            df_invoice['배송메세지'] = df_invoice['배송메세지'].fillna('')
        if '상품명' in df_invoice.columns:
            df_invoice['상품명'] = '전자제품'
        if '상품모델' in df_invoice.columns:
            df_invoice['상품모델'] = '전자제품'

        columns = [
            '주문번호',
            '고객주문처명',
            '수취인명',
            '우편번호',
            '수취인 주소',
            '수취인 전화번호',
            '수취인 이동통신',
            '상품명',
            '상품모델',
            '배송메세지',
            '비고'
        ]
        df_invoice = df_invoice[columns]

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_invoice.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            workbook = writer.book

            center_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter'
            })
            header_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'bold': True
            })

            for idx, col in enumerate(df_invoice.columns):
                max_length = max(
                    df_invoice[col].astype(str).apply(len).max(),
                    len(str(col))
                )
                adjusted_width = max_length * 2 if any('\u3131' <= c <= '\u318E' or '\uAC00' <= c <= '\uD7A3' for c in str(col)) else max_length
                worksheet.set_column(idx, idx, adjusted_width + 2, center_format)

            for col_num, value in enumerate(df_invoice.columns.values):
                worksheet.write(0, col_num, value, header_format)

        print(f"✓ 퀵 엑셀 저장 완료: {output_file}")
        self.show_excel_created_message(output_file, "송장 엑셀 파일이 생성되었습니다.")
        return output_file

    def show_excel_created_message(self, output_file, message_text):
        """엑셀 생성 완료 메시지를 표시합니다."""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("완료")
        msg.setText(message_text)

        open_location_button = msg.addButton("폴더 열기", QMessageBox.ActionRole)
        open_file_button = msg.addButton("엑셀 열기", QMessageBox.ActionRole)
        close_button = msg.addButton("닫기", QMessageBox.RejectRole)

        msg.setDefaultButton(close_button)
        msg.exec()

        if msg.clickedButton() == open_location_button:
            if not self.open_file_location(output_file):
                QMessageBox.warning(self, "오류", "파일 위치를 열 수 없습니다.")
        elif msg.clickedButton() == open_file_button:
            if not self.open_file_with_default_app(output_file):
                QMessageBox.warning(self, "오류", "파일을 열 수 없습니다.\n엑셀이 설치되어 있는지 확인해주세요.")

    def on_naver_index_changed(self, text):
        """네이버 인덱스가 수동으로 변경되었을 때 호출됩니다."""
        try:
            if text.strip() and text.strip().isdigit():
                self.current_idx_naver = int(text.strip())
                self.save_index_values()
                print(f"✓ 네이버 인덱스 수동 변경: {self.current_idx_naver}")
        except Exception as e:
            print(f"! 네이버 인덱스 변경 중 오류 발생: {str(e)}")

    def on_coupang_index_changed(self, text):
        """쿠팡 인덱스가 수동으로 변경되었을 때 호출됩니다."""
        try:
            if text.strip() and text.strip().isdigit():
                self.current_idx_coupang = int(text.strip())
                self.save_index_values()
                print(f"✓ 쿠팡 인덱스 수동 변경: {self.current_idx_coupang}")
        except Exception as e:
            print(f"! 쿠팡 인덱스 변경 중 오류 발생: {str(e)}")


    def on_gmarket_index_changed(self, text):
        """지마켓 인덱스가 수동으로 변경되었을 때 호출됩니다."""
        try:
            if text.strip() and text.strip().isdigit():
                self.current_idx_gmarket = int(text.strip())
                self.save_index_values()
                print(f"✓ 지마켓 인덱스 수동 변경: {self.current_idx_gmarket}")
        except Exception as e:
            print(f"! 지마켓 인덱스 변경 중 오류 발생: {str(e)}")

    def is_valid_filename(self, filename):
        """파일명이 올바른 형식인지 검사합니다."""
        print(f"\n[파일명 검증 시작] 파일명: {filename}")
        
        # 네이버 스토어 파일 형식 검사
        naver_pattern = r'^스마트스토어_(전체|선택)주문발주발송관리_(\d{8}).*\.xlsx$'
        naver_match = re.match(naver_pattern, filename)
        
        print("\n[네이버 스토어 패턴 검사]")
        print(f"패턴: {naver_pattern}")
        print(f"매칭 결과: {naver_match}")
        if naver_match:
            print(f"매칭된 그룹: {naver_match.groups()}")
        
        # 쿠팡 스토어 파일 형식 검사
        coupang_pattern = r'^DeliveryList\((\d{4}-\d{2}-\d{2})\).*\.xlsx$'
        coupang_match = re.match(coupang_pattern, filename)
        
        print("\n[쿠팡 스토어 패턴 검사]")
        print(f"패턴: {coupang_pattern}")
        print(f"매칭 결과: {coupang_match}")
        if coupang_match:
            print(f"매칭된 그룹: {coupang_match.groups()}")
        
        # 지마켓 스토어 파일 형식 검사
        gmarket_pattern = r'^발송관리(?: \((\d+)\))?\.xlsx$'
        gmarket_match = re.match(gmarket_pattern, filename)
        
        print("\n[지마켓 스토어 패턴 검사]")
        print(f"패턴: {gmarket_pattern}")
        print(f"매칭 결과: {gmarket_match}")
        if gmarket_match:
            print(f"매칭된 그룹: {gmarket_match.groups()}")
        
        if naver_match:
            # 날짜 유효성 검사
            date_str = naver_match.group(2)
            
            try:
                # 날짜 형식 검증 (YYYYMMDD)
                date = datetime.strptime(date_str, '%Y%m%d')
                print(f"✓ 날짜 형식 검증 완료: {date.strftime('%Y년 %m월 %d일')}")
                print("✓ 네이버 스토어 파일명 검증 성공")
                return True, "naver"
                
            except ValueError:
                print("❌ 날짜 형식이 올바르지 않습니다.")
                return False, None
        elif coupang_match:
            # 날짜 유효성 검사
            date_str = coupang_match.group(1)
            
            try:
                # 날짜 형식 검증 (YYYY-MM-DD)
                date = datetime.strptime(date_str, '%Y-%m-%d')
                print(f"✓ 날짜 형식 검증 완료: {date.strftime('%Y년 %m월 %d일')}")
                print("✓ 쿠팡 스토어 파일명 검증 성공")
                return True, "coupang"
                
            except ValueError:
                print("❌ 날짜 형식이 올바르지 않습니다.")
                return False, None
        elif gmarket_match:
            print("✓ 지마켓 스토어 파일명 검증 성공")
            return True, "gmarket"
        else:
            print("❌ 파일명 형식이 올바르지 않습니다.")
            print("   - 네이버 스토어 형식: 스마트스토어_전체주문발주발송관리_YYYYMMDD로 시작하는 .xlsx 파일")
            print("   - 쿠팡 스토어 형식: DeliveryList(YYYY-MM-DD)로 시작하는 .xlsx 파일")
            print("   - 지마켓 스토어 형식: 발송관리.xlsx 또는 발송관리 (숫자).xlsx")
            return False, None
            
    def open_file_with_default_app(self, file_path):
        """기본 애플리케이션으로 파일을 엽니다."""
        try:
            if sys.platform == 'win32':
                os.startfile(file_path)
            elif sys.platform == 'darwin':  # macOS
                subprocess.call(('open', file_path))
            else:  # linux
                subprocess.call(('xdg-open', file_path))
            return True
        except Exception as e:
            print(f"! 파일 열기 실패: {str(e)}")
            return False
            
    def select_excel_file(self):
        """주문서 엑셀 파일을 선택하는 다이얼로그를 표시합니다."""
        # 다운로드 폴더 경로 설정
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        
        file_path, _ = QFileDialog.getOpenFileName(
        self,
        "엑셀 파일 선택",
        downloads_path,
        "Excel Files (*.xlsx *.xls)"
        )                
        if file_path:
            print(f"\n[파일 선택됨] 경로: {file_path}")
            filename = os.path.basename(file_path)
            
            is_valid, self.store_type = self.is_valid_filename(filename)
            
            if is_valid:
                self.selected_file_path = file_path
                self.ui.filePathLabel.setText(filename)
                self.statusBar().showMessage(f"파일 선택됨: {filename}")
                print("✓ 파일이 성공적으로 선택되었습니다.")

                self._show_busy_processing_overlay(
                    "주문 정보를 불러오는 중…",
                    "스프레드시트와 엑셀을 처리하고 있습니다. 잠시만 기다려 주세요.",
                )
                QApplication.processEvents()

                order_processing_async = False
                try:
                    self.refresh_order_indices_from_sheet(interactive=False)

                    # 스토어 타입에 따라 로고 표시
                    if self.store_type == "naver":
                        logo_path = "image/naver-logo.png"
                        logo_size = QSize(120, 40)  # 네이버 로고 크기
                    elif self.store_type == "coupang":
                        logo_path = "image/coupang-logo.png"
                        logo_size = QSize(120, 40)  # 쿠팡 로고 크기
                    elif self.store_type == "gmarket":
                        logo_path = "image/gmarket-logo.png"
                        logo_size = QSize(120, 40)  # 지마켓 로고 크기

                    # 로고 이미지 로드 및 표시
                    if os.path.exists(logo_path):
                        pixmap = QPixmap(logo_path)
                        scaled_pixmap = pixmap.scaled(
                            logo_size, Qt.KeepAspectRatio, Qt.SmoothTransformation
                        )
                        self.ui.label_logo.setPixmap(scaled_pixmap)
                        self.ui.label_logo.setAlignment(Qt.AlignCenter)
                    else:
                        print(f"! 로고 파일을 찾을 수 없습니다: {logo_path}")

                    # 스토어 타입에 따라 다른 처리 메서드 호출
                    try:
                        if self.store_type == "naver":
                            print("✓ 네이버 스토어 파일 처리 시작")
                            order_processing_async = True
                            self._run_order_file_processing_with_async_mapping("naver")
                        elif self.store_type == "coupang":
                            print("✓ 쿠팡 스토어 파일 처리 시작")
                            order_processing_async = True
                            self._run_order_file_processing_with_async_mapping("coupang")
                        elif self.store_type == "gmarket":
                            print("✓ 지마켓 스토어 파일 처리 시작")
                            self.process_gmarket_excel_file()
                            self.is_order_file_valid = True
                    except Exception as e:
                        self.is_order_file_valid = False
                        print(f"❌ 파일 처리 중 오류 발생: {str(e)}")
                        QMessageBox.warning(
                            self, "오류", f"파일 처리 중 오류가 발생했습니다: {str(e)}"
                        )
                finally:
                    if not order_processing_async:
                        self._hide_busy_processing_overlay()
            else:
                self.is_order_file_valid = False
                QMessageBox.warning(
                    self,
                    "잘못된 파일명",
                    "올바른 파일명 형식이 아닙니다.\n\n"
                    "네이버 스토어: 스마트스토어_전체주문발주발송관리_YYYYMMDD_HHMM.xlsx\n"
                    "쿠팡 스토어: DeliveryList(YYYY-MM-DD)_(0).xlsx"
                )
                self.selected_file_path = None
                self.ui.filePathLabel.setText("주문 정보가 없습니다. ( Ctrl + O )")
                self.statusBar().showMessage("잘못된 파일명")
                print("❌ 파일 선택이 취소되었습니다.")
        else:
            self.is_order_file_valid = False
            print("\n[알림] 파일 선택이 취소되었습니다.")     
    
    def process_naver_excel_file(self):
        """네이버 스토어 엑셀 파일에서 주문번호를 처리합니다."""
        try:
            print(f"\n[네이버 스토어 엑셀 파일 처리 시작] 파일: {self.selected_file_path}")

            # 스프레드시트에서 상품번호(E) → 상품코드(A) 매핑 로드
            print("\n[스프레드시트 상품번호-상품코드 매핑 로드]")
            product_mapping = self._load_product_code_map_from_spreadsheet("naver")

            
            # print(f"\n총 {len(product_mapping)}개의 상품 매핑이 생성되었습니다.")
            # print("\n[생성된 매핑 목록]")
            # for num, code in product_mapping.items():
            #     print(f"{num} -> {code}")
            
            # 파일 정보 출력
            file_size = os.path.getsize(self.selected_file_path)
            print(f"파일 크기: {file_size:,} 바이트")
            
            # 파일 헤더 확인
            try:
                with open(self.selected_file_path, 'rb') as f:
                    header = f.read(8)
                    print(f"파일 헤더 (16진수): {header.hex()}")
                    if header.startswith(b'PK\x03\x04'):
                        print("✓ 파일이 ZIP 형식(.xlsx)으로 확인됨")
                    elif header.startswith(b'\xD0\xCF\x11\xE0'):
                        print("✓ 파일이 OLE2 형식(.xls)으로 확인됨")
                    else:
                        print("! 알 수 없는 파일 형식")
            except Exception as e:
                print(f"! 파일 헤더 확인 실패: {str(e)}")
            
            # 비밀번호 고정값 사용
            password = "1234"
            print(f"✓ 고정 비밀번호 사용: {password}")
            
            # 비밀번호 보호 파일 처리
            try:
                # 임시 파일 생성
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                temp_file.close()
                temp_path = temp_file.name
                
                print(f"✓ 임시 파일 생성: {temp_path}")
                
                # 파일 복호화
                with open(self.selected_file_path, 'rb') as file:
                    office_file = msoffcrypto.OfficeFile(file)
                    office_file.load_key(password=password)
                    
                    with open(temp_path, 'wb') as output_file:
                        office_file.decrypt(output_file)
                
                print("✓ 파일 복호화 완료")
                
                # 복호화된 파일 읽기
                df = pd.read_excel(temp_path, header=1)
                print(f"✓ 파일 읽기 성공")
                print(f"데이터프레임 정보:")
                print(f"- 행 수: {len(df)}")
                print(f"- 열 수: {len(df.columns)}")
                
                # 임시 파일 삭제
                os.unlink(temp_path)
                print("✓ 임시 파일 삭제 완료")
                
            except Exception as e:
                print(f"! 비밀번호 보호 파일 처리 실패: {str(e)}")
                
                # 기존 방식으로 시도
                print("\n기존 방식으로 파일 읽기 시도...")
                
                # 다양한 엔진으로 파일 읽기 시도
                engines_to_try = [
                    ('openpyxl', {'engine': 'openpyxl'}),
                    ('xlrd', {'engine': 'xlrd'}),
                    ('openpyxl', {'engine': 'openpyxl', 'data_only': True}),
                    ('pyxlsb', {'engine': 'pyxlsb'}),
                ]
                
                df = None
                last_error = None
                
                for engine_name, options in engines_to_try:
                    try:
                        print(f"\n✓ {engine_name} 엔진으로 시도 중... (옵션: {options})")
                        options['header'] = 1  # 2행을 헤더로 사용
                        
                        # 시트 정보 확인
                        if engine_name == 'openpyxl':
                            import openpyxl
                            wb = openpyxl.load_workbook(
                                self.selected_file_path, 
                                read_only=True, 
                                data_only=options.get('data_only', False)
                            )
                            print(f"시트 목록: {wb.sheetnames}")
                            print(f"활성 시트: {wb.active.title}")
                            wb.close()
                        
                        df = pd.read_excel(self.selected_file_path, **options)
                        print(f"✓ {engine_name} 엔진으로 파일 읽기 성공")
                        print(f"데이터프레임 정보:")
                        print(f"- 행 수: {len(df)}")
                        print(f"- 열 수: {len(df.columns)}")
                        
                        # 임시 파일 삭제
                        os.unlink(temp_path)
                        print("✓ 임시 파일 삭제 완료")
                        
                        break
                    except Exception as e:
                        print(f"! {engine_name} 엔진 실패: {str(e)}")
                        last_error = e
                
                if df is None:
                    error_msg = "모든 엔진으로 파일 읽기 실패"
                    if last_error:
                        error_msg += f"\n마지막 오류: {str(last_error)}"
                    print(f"\n파일 읽기 시도 결과:")
                    print("- 파일이 손상되었거나 지원되지 않는 형식일 수 있습니다.")
                    print("- 파일을 다시 저장하거나 다른 형식으로 변환해보세요.")
                    
                    # 파일 열기 시도
                    msg = QMessageBox()
                    msg.setIcon(QMessageBox.Critical)
                    msg.setWindowTitle("파일 읽기 오류")
                    msg.setText("엑셀 파일을 읽을 수 없습니다.")
                    msg.setInformativeText(
                        "파일이 손상되었거나 지원되지 않는 형식일 수 있습니다.\n"
                        "파일을 직접 열어서 확인하시겠습니까?"
                    )
                    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                    msg.setDefaultButton(QMessageBox.Yes)
                    
                    if msg.exec() == QMessageBox.Yes:
                        if self.open_file_with_default_app(self.selected_file_path):
                            print("✓ 파일을 기본 애플리케이션으로 열었습니다.")
                        else:
                            print("! 파일을 열 수 없습니다.")
                    
                    raise Exception(error_msg)
            
            # 필요한 열 찾기
            required_columns = {
                '주문번호': None,
                '수취인명': None,
                '수취인연락처1': None,
                '통합배송지': None,
                '구매자연락처': None,
                '배송메세지': None,
                '상품명': None,
                '옵션정보': None,
                '수량': None,
                '우편번호': None,
                '상품번호': None,  # 상품번호 열 추가
                '배송방법(구매자 요청)': None,  # 배송 방법 열 추가
                '최종 상품별 총 주문금액': None  # AA열 금액 정보 추가
            }
            
            for col in df.columns:
                col_str = str(col).strip()
                for key in required_columns.keys():
                    # if col_str == key:  # 정확히 일치하는 경우에만 매칭
                    if col_str.replace(' ', '') == key.replace(' ', ''):  # 공백 제거 후 비교
                        required_columns[key] = col
                        print(f"✓ '{key}' 열을 찾았습니다: {col}")
            
            # AA열(최종 상품별 총 주문금액)을 찾지 못한 경우 인덱스로 직접 접근 시도
            if required_columns['최종 상품별 총 주문금액'] is None:
                # AA열은 27번째 열 (0-based로는 26, pandas는 0-based)
                if len(df.columns) > 26:
                    required_columns['최종 상품별 총 주문금액'] = df.columns[26]
                    print(f"✓ '최종 상품별 총 주문금액' 열을 인덱스로 찾았습니다: {df.columns[26]}")
                else:
                    print("⚠️ AA열(최종 상품별 총 주문금액)을 찾을 수 없습니다. 금액 정보가 없을 수 있습니다.")

            # 필수 열이 모두 있는지 확인 (금액 열은 선택사항으로 처리)
            missing_columns = [key for key, value in required_columns.items() if value is None and key != '최종 상품별 총 주문금액']
            if missing_columns:
                print(f"❌ 다음 열을 찾을 수 없습니다: {', '.join(missing_columns)}")
                QMessageBox.warning(self, "오류", f"다음 열을 찾을 수 없습니다:\n{', '.join(missing_columns)}")
                return
            
            # 주문 정보 정리
            print("\n[주문 정보 정리]")
            self.orders = {}  # 기존 orders를 self.orders로 변경
            
            # 주문번호 패턴 분석
            order_patterns = {}
            
            # 먼저 주문번호 패턴 분석
            for _, row in df.iterrows():
                order_number = str(row[required_columns['주문번호']])
                if pd.isna(order_number) or order_number.strip() == '':
                    continue
                
                # 주문번호 패턴 추출 (앞 13자리)
                pattern = order_number[:13] if len(order_number) >= 13 else order_number
                
                if pattern not in order_patterns:
                    order_patterns[pattern] = []
                
                order_patterns[pattern].append(order_number)
            
            # 패턴별 주문 정보 정리
            for pattern, order_numbers in order_patterns.items():
                # 패턴에 해당하는 첫 번째 주문 정보 가져오기
                first_order = None
                for _, row in df.iterrows():
                    order_number = str(row[required_columns['주문번호']])
                    if order_number in order_numbers:
                        first_order = row
                        break
                
                if first_order is None:
                    continue
                
                # 주문 정보 초기화
                self.orders[pattern] = {
                    '주문번호목록': order_numbers,
                    '수취인명': str(first_order[required_columns['수취인명']]),
                    '수취인연락처1': str(first_order[required_columns['수취인연락처1']]),
                    '통합배송지': str(first_order[required_columns['통합배송지']]),
                    '구매자연락처': str(first_order[required_columns['구매자연락처']]),
                    '배송메세지': str(first_order[required_columns['배송메세지']]),
                    '우편번호': str(first_order[required_columns['우편번호']]),
                    '배송방법': str(first_order[required_columns['배송방법(구매자 요청)']]) if not pd.isna(first_order[required_columns['배송방법(구매자 요청)']]) else '배송방법 오류',
                    '상품수': 0,
                    '상품목록': [],
                    '주문총액': 0  # 주문 총액 초기화
                }
                
                # 배송방법 확인을 위한 플래그 (하나라도 '택배,등기,소포'가 있으면 True)
                has_delivery_method = False
                
                # 패턴에 해당하는 모든 주문의 상품 정보 추가
                for _, row in df.iterrows():
                    order_number = str(row[required_columns['주문번호']])
                    if order_number in order_numbers:
                        # 상품 정보
                        product_name = str(row[required_columns['상품명']])
                        quantity = int(row[required_columns['수량']]) if not pd.isna(row[required_columns['수량']]) else 1
                        option = str(row[required_columns['옵션정보']]) if not pd.isna(row[required_columns['옵션정보']]) else "없음"
                        
                        # 배송방법 확인 (배송비 중복 결제 방지: 하나라도 '택배,등기,소포'가 있으면 전체를 '택배,등기,소포'로 처리)
                        delivery_method = str(row[required_columns['배송방법(구매자 요청)']]) if not pd.isna(row[required_columns['배송방법(구매자 요청)']]) else ''
                        if delivery_method == '택배,등기,소포':
                            has_delivery_method = True
                        
                        # 상품번호 가져오기
                        product_number = self._normalize_key_for_mapping(row[required_columns['상품번호']])
                        print(f"\n[상품번호 매칭]")
                        print(f"상품명: {product_name}")
                        print(f"원본 상품번호: {row[required_columns['상품번호']]}")
                        print(f"변환된 상품번호: {product_number}")
                        # print(f"매핑 딕셔너리 키 목록: {list(product_mapping.keys())}")
                        product_code = product_mapping.get(product_number, '') or '        '
                        print(f"매칭된 상품코드: {product_code}")
                        
                        # 금액 정보 가져오기 (AA열)
                        product_amount = 0
                        if required_columns['최종 상품별 총 주문금액'] is not None:
                            try:
                                amount_value = row[required_columns['최종 상품별 총 주문금액']]
                                if not pd.isna(amount_value):
                                    # 숫자로 변환 시도
                                    if isinstance(amount_value, str):
                                        # 쉼표 제거 후 숫자 변환
                                        amount_value = amount_value.replace(',', '').strip()
                                    product_amount = float(amount_value)
                            except (ValueError, TypeError) as e:
                                print(f"⚠️ 금액 변환 실패: {amount_value}, 오류: {e}")
                                product_amount = 0
                        
                        # 주문 총액에 추가
                        self.orders[pattern]['주문총액'] += product_amount
                        
                        # 상품 정보 추가
                        self.orders[pattern]['상품수'] += 1
                        self.orders[pattern]['상품목록'].append({
                            '상품명': product_name,
                            '수량': quantity,
                            '옵션': option,
                            '상품코드': product_code,
                            '금액': product_amount,
                            '상품번호': product_number,
                        })
                        
                        # 수취인 정보가 다른 경우 경고
                        if (self.orders[pattern]['수취인명'] != str(row[required_columns['수취인명']]) or
                            self.orders[pattern]['수취인연락처1'] != str(row[required_columns['수취인연락처1']]) or
                            self.orders[pattern]['통합배송지'] != str(row[required_columns['통합배송지']])):
                            print(f"! 주문번호 패턴 {pattern}의 수취인 정보가 다릅니다:")
                            print(f"  - 기존: {self.orders[pattern]['수취인명']} / {self.orders[pattern]['수취인연락처1']} / {self.orders[pattern]['통합배송지']}")
                            print(f"  - 새로운: {row[required_columns['수취인명']]} / {row[required_columns['수취인연락처1']]} / {row[required_columns['통합배송지']]}")
                
                # 배송방법 최종 결정: 하나라도 '택배,등기,소포'가 있으면 전체를 '택배,등기,소포'로 설정
                if has_delivery_method:
                    self.orders[pattern]['배송방법'] = '택배,등기,소포'
                    print(f"✓ 주문번호 패턴 {pattern}: 배송비 중복 결제 방지를 위해 '택배,등기,소포'로 설정됨")
            
            # 주문 정보 출력
            print("\n[주문 정보 출력]")
            print(f"총 {len(self.orders)}개의 주문이 있습니다.")
            
            # 마크다운 형식으로 주문 정보 생성
            markdown_text = ""
            
            # 인덱스 값 다시 로드
            if hasattr(self.ui, 'lineEdit_idx_naver'):
                saved_index = self.ui.lineEdit_idx_naver.text().strip()
                if saved_index and saved_index.isdigit():
                    self.current_idx_naver = int(saved_index)
                else:
                    self.current_idx_naver = 1
                    self.ui.lineEdit_idx_naver.setText(str(self.current_idx_naver))
            
            for pattern, info in self.orders.items():
                # 주문 총액 가져오기
                total_amount = info.get('주문총액', 0)
                
                # 만원 단위로 포맷팅 (0.1만 단위로 올림)
                # 예: 950 -> 0.1만, 61000 -> 6.1만, 63820 -> 6.4만
                amount_rounded = math.ceil(total_amount / 1000) / 10 if total_amount > 0 else 0
                formatted_amount = f"{amount_rounded}만"
                
                # 배송 방법이 '택배,등기,소포'인 경우에만 기존 형식으로 표시
                if info['배송방법'] == '택배,등기,소포':
                    markdown_text += f"[ ] {self.current_idx_naver}.{info['수취인명']} - {formatted_amount}\n"
                else:
                    # 그 외의 경우 배송 방법을 함께 표시
                    markdown_text += f"[ ] {self.current_idx_naver}.{info['수취인명']} - {formatted_amount} **({info['배송방법']})**\n"
                
                self.update_naver_index()
                if hasattr(self.ui, 'lineEdit_idx_naver'):
                    self.ui.lineEdit_idx_naver.setText(str(self.current_idx_naver))
                
                # 상품 목록 표시
                for product in info['상품목록']:
                    product_name = product['상품명']
                    quantity = product['수량']
                    option = product['옵션']
                    product_code = product['상품코드']
                    product_no = product.get('상품번호') or ''
                    name_for_md = self._format_naver_product_name_markdown(
                        product_name, product_no
                    )
                    markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {name_for_md} ( 옵션 : {option} )\n"
                
                markdown_text += "\n"  # 주문 간 구분을 위한 빈 줄
            
            # plainTextEdit에 마크다운 텍스트 표시
            self.ui.plainTextEdit.setPlainText(markdown_text)
            
            # 클립보드에 자동 복사
            clipboard = QApplication.clipboard()
            clipboard.setText(markdown_text)
            self.statusBar().showMessage("주문 정보가 클립보드에 복사되었습니다.", 2000)
            
        except Exception as e:
            error_msg = str(e)
            print(f"❌ 엑셀 파일 처리 중 오류 발생: {error_msg}")
            if _is_likely_google_sheets_oauth_error(e):
                QMessageBox.critical(
                    self,
                    "오류",
                    "Google 스프레드시트 연동 중 오류가 발생했습니다.\n\n"
                    f"{error_msg}\n\n"
                    + self._oauth_error_dialog_hint(),
                )
            else:
                QMessageBox.critical(
                    self,
                    "오류",
                    "엑셀 파일 처리 중 오류가 발생했습니다.\n\n"
                    f"{error_msg}\n\n"
                    "다음 사항을 확인해주세요:\n"
                    "1. 파일이 손상되지 않았는지\n"
                    "2. 다른 프로그램에서 파일을 열고 있지 않은지\n"
                    "3. 파일을 다시 저장하거나 다른 형식(.xlsx)으로 변환해보세요.",
                )

    def process_coupang_excel_file(self):
        """쿠팡 스토어 엑셀 파일에서 주문번호를 처리합니다."""
        try:
            print(f"\n[쿠팡 스토어 엑셀 파일 처리 시작] 파일: {self.selected_file_path}")
            
            # 파일 정보 출력
            file_size = os.path.getsize(self.selected_file_path)
            print(f"파일 크기: {file_size:,} 바이트")
            
            # 파일 헤더 확인
            try:
                with open(self.selected_file_path, 'rb') as f:
                    header = f.read(8)
                    print(f"파일 헤더 (16진수): {header.hex()}")
                    if header.startswith(b'PK\x03\x04'):
                        print("✓ 파일이 ZIP 형식(.xlsx)으로 확인됨")
                    elif header.startswith(b'\xD0\xCF\x11\xE0'):
                        print("✓ 파일이 OLE2 형식(.xls)으로 확인됨")
                    else:
                        print("! 알 수 없는 파일 형식")
            except Exception as e:
                print(f"! 파일 헤더 확인 실패: {str(e)}")
            
            # 엑셀 파일 읽기
            df = pd.read_excel(self.selected_file_path)
            print(f"\n[열 정보]")
            print(f"감지된 열 목록: {', '.join(str(col) for col in df.columns)}")

            # 스프레드시트에서 옵션ID(E) → 상품코드(A) 매핑 로드
            try:
                print("\n[스프레드시트 옵션ID-상품코드 매핑 로드]")
                product_code_map = self._load_product_code_map_from_spreadsheet("coupang")
            except Exception as e:
                print(f"! 스프레드시트 매핑 로드 중 오류 발생: {str(e)}")
                msg = f"스프레드시트 매핑 로드 중 오류가 발생했습니다.\n\n{str(e)}"
                if _is_likely_google_sheets_oauth_error(e):
                    msg += "\n\n" + self._oauth_error_dialog_hint()
                QMessageBox.warning(self, "경고", msg)
                return
            
            # 필요한 열 찾기
            required_columns = {
                '주문번호': None,
                '수취인이름': None,
                '수취인 주소': None,
                '수취인전화번호': None,
                '노출상품명(옵션명)': None,
                '등록옵션명': None,
                '구매수(수량)': None,
                '배송메세지': None,
                '우편번호': None,
                '옵션ID': None,  # 공백 제거
                '결제액': None  # S열 결제액 추가
            }
            
            for col in df.columns:
                col_str = str(col).strip()
                for key in required_columns.keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        required_columns[key] = col
                        print(f"✓ '{key}' 열을 찾았습니다: {col}")
            
            # S열(결제액)을 찾지 못한 경우 인덱스로 직접 접근 시도
            if required_columns['결제액'] is None:
                # S열은 19번째 열 (0-based로는 18, pandas는 0-based)
                if len(df.columns) > 18:
                    required_columns['결제액'] = df.columns[18]
                    print(f"✓ '결제액' 열을 인덱스로 찾았습니다: {df.columns[18]}")
                else:
                    print("⚠️ S열(결제액)을 찾을 수 없습니다. 금액 정보가 없을 수 있습니다.")
            
            # 필수 열이 모두 있는지 확인 (결제액 열은 선택사항으로 처리)
            missing_columns = [key for key, value in required_columns.items() if value is None and key != '결제액']
            if missing_columns:
                print(f"❌ 다음 열을 찾을 수 없습니다: {', '.join(missing_columns)}")
                QMessageBox.warning(self, "오류", f"다음 열을 찾을 수 없습니다:\n{', '.join(missing_columns)}")
                return
            
            # 주문 정보 정리
            self.orders = {}
            
            # 주문번호별로 주문 정보 정리
            for _, row in df.iterrows():
                order_number = str(row[required_columns['주문번호']])
                if pd.isna(order_number) or order_number.strip() == '':
                    continue
                
                if order_number not in self.orders:
                    phone_value = row[required_columns['수취인전화번호']]
                    
                    self.orders[order_number] = {
                        '수취인이름': str(row[required_columns['수취인이름']]),
                        '수취인주소': str(row[required_columns['수취인 주소']]),
                        '수취인전화번호': str(row[required_columns['수취인전화번호']]) if not pd.isna(row[required_columns['수취인전화번호']]) else '',
                        '배송메세지': str(row[required_columns['배송메세지']]) if not pd.isna(row[required_columns['배송메세지']]) else '',
                        '우편번호': str(row[required_columns['우편번호']]) if not pd.isna(row[required_columns['우편번호']]) else '',
                        '상품목록': [],
                        '결제액': 0  # 주문별 결제액 초기화 (각 행마다 합산)
                    }
                
                # 각 행마다 결제액 읽어서 합산 (같은 주문번호에 여러 행이 있을 수 있음)
                if required_columns['결제액'] is not None:
                    try:
                        payment_value = row[required_columns['결제액']]
                        if not pd.isna(payment_value):
                            # 숫자로 변환 시도
                            if isinstance(payment_value, str):
                                # 쉼표 제거 후 숫자 변환
                                payment_value = payment_value.replace(',', '').strip()
                            row_payment_amount = float(payment_value)
                            # 주문 총액에 추가
                            self.orders[order_number]['결제액'] += row_payment_amount
                    except (ValueError, TypeError) as e:
                        print(f"⚠️ 결제액 변환 실패: {payment_value}, 오류: {e}")
                
                # 상품 정보 추가
                product_name = str(row[required_columns['노출상품명(옵션명)']])
                option = str(row[required_columns['등록옵션명']])
                quantity = int(row[required_columns['구매수(수량)']]) if not pd.isna(row[required_columns['구매수(수량)']]) else 1
                
                # 옵션ID로 상품코드 찾기
                option_id_raw = row[required_columns['옵션ID']]
                
                # NaN 값 처리
                if pd.isna(option_id_raw):
                    option_id = ''
                else:
                    option_id = self._normalize_key_for_mapping(option_id_raw)
                
                product_code = product_code_map.get(option_id, '')  # 매칭되는 상품코드가 없으면 빈 문자열
                vp_product_no = self._coupang_option_to_vp_product_no.get(option_id, '') if option_id else ''
                
                self.orders[order_number]['상품목록'].append({
                    '상품명': product_name,
                    '옵션': option,
                    '수량': quantity,
                    '상품코드': product_code,
                    '쿠팡상품번호': vp_product_no,
                })
            
            # 마크다운 형식으로 주문 정보 생성
            markdown_text = ""
            
            # 인덱스 값 다시 로드
            if hasattr(self.ui, 'lineEdit_idx_coupang'):
                saved_index = self.ui.lineEdit_idx_coupang.text().strip()
                if saved_index and saved_index.isdigit():
                    self.current_idx_coupang = int(saved_index)
                else:
                    self.current_idx_coupang = 1
                    self.ui.lineEdit_idx_coupang.setText(str(self.current_idx_coupang))
            
            for order_number, info in self.orders.items():
                # 주문번호별 결제액 가져오기
                total_amount = info.get('결제액', 0)
                
                # 만원 단위로 포맷팅 (0.1만 단위로 올림)
                # 예: 950 -> 0.1만, 61000 -> 6.1만, 63820 -> 6.4만
                amount_rounded = math.ceil(total_amount / 1000) / 10 if total_amount > 0 else 0
                formatted_amount = f"{amount_rounded}만"
                
                markdown_text += f"[ ] {self.current_idx_coupang}.{info['수취인이름']} - {formatted_amount}\n"
                self.update_coupang_index()
                if hasattr(self.ui, 'lineEdit_idx_coupang'):
                    self.ui.lineEdit_idx_coupang.setText(str(self.current_idx_coupang))
                
                # 상품 목록 표시
                for product in info['상품목록']:
                    product_name = product['상품명']
                    quantity = product['수량']
                    option = product['옵션']
                    product_code = product['상품코드']
                    vp_no = product.get('쿠팡상품번호') or ''
                    name_for_md = self._format_product_name_markdown_link(
                        product_name, vp_no, COUPANG_VP_PRODUCT_URL_PREFIX
                    )
                    markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {name_for_md} ( 옵션 : {option} )\n"
                
                markdown_text += "\n"
            
            # plainTextEdit에 마크다운 텍스트 표시
            self.ui.plainTextEdit.setPlainText(markdown_text)
            
            # 클립보드에 자동 복사
            clipboard = QApplication.clipboard()
            clipboard.setText(markdown_text)
            self.statusBar().showMessage("주문 정보가 클립보드에 복사되었습니다.", 2000)
            
        except Exception as e:
            error_msg = str(e)
            print(f"❌ 엑셀 파일 처리 중 오류 발생: {error_msg}")
            if _is_likely_google_sheets_oauth_error(e):
                QMessageBox.critical(
                    self,
                    "오류",
                    "Google 스프레드시트 연동 중 오류가 발생했습니다.\n\n"
                    f"{error_msg}\n\n"
                    + self._oauth_error_dialog_hint(),
                )
            else:
                QMessageBox.critical(
                    self,
                    "오류",
                    f"엑셀 파일 처리 중 오류가 발생했습니다.\n\n{error_msg}",
                )

    def process_gmarket_excel_file(self):
        """지마켓 스토어의 주문 정보 엑셀 파일을 처리합니다."""
        try:
            print(f"\n[지마켓 스토어 엑셀 파일 처리 시작] 파일: {self.selected_file_path}")
            
            # 파일 정보 출력
            file_size = os.path.getsize(self.selected_file_path)
            print(f"파일 크기: {file_size:,} 바이트")
            
            # 파일 헤더 확인
            try:
                with open(self.selected_file_path, 'rb') as f:
                    header = f.read(8)
                    print(f"파일 헤더 (16진수): {header.hex()}")
                    if header.startswith(b'PK\x03\x04'):
                        print("✓ 파일이 ZIP 형식(.xlsx)으로 확인됨")
                    elif header.startswith(b'\xD0\xCF\x11\xE0'):
                        print("✓ 파일이 OLE2 형식(.xls)으로 확인됨")
                    else:
                        print("! 알 수 없는 파일 형식")
            except Exception as e:
                print(f"! 파일 헤더 확인 실패: {str(e)}")
            
            # 엑셀 파일 읽기
            df = pd.read_excel(self.selected_file_path)
            print(f"\n[열 정보]")
            print(f"감지된 열 목록: {', '.join(str(col) for col in df.columns)}")
            
            # store_database.xlsx 파일 읽기
            # try:
            #     db_path = Path("database") / "store_database.xlsx"
            #     if not db_path.exists():
            #         raise FileNotFoundError("store_database.xlsx 파일을 찾을 수 없습니다.")
                
            #     # 두 번째 시트 읽기 (시트 이름이 날짜로 변경될 수 있으므로 인덱스로 접근)
            #     db_df = pd.read_excel(db_path, sheet_name=1, header=1)
            #     print("\n[상품 데이터베이스 로드 완료]")
            #     print("데이터베이스 열 목록:")
            #     for col in db_df.columns:
            #         print(f"- {col}")
                
            #     # 상품코드와 옵션ID 매핑 생성
            #     product_code_map = {}
            #     for _, row in db_df.iterrows():
            #         option_id = str(row['옵션 ID']).strip()
            #         product_code = str(row['상품코드']).strip()
            #         if option_id and not pd.isna(option_id):
            #             product_code_map[option_id] = product_code
                
            #     print(f"✓ {len(product_code_map)}개의 상품 매핑 정보를 로드했습니다.")
                
            # except Exception as e:
            #     print(f"! 상품 데이터베이스 로드 중 오류 발생: {str(e)}")
            #     QMessageBox.warning(self, "경고", "상품 데이터베이스 로드 중 오류가 발생했습니다.")
            #     return
            
            # 필요한 열 찾기
            required_columns = {
                '주문번호': None,
                '수령인명': None,
                '주소': None,
                '수령인 전화번호': None,
                '수령인 휴대폰': None,
                '상품명': None,
                '옵션': None,
                '수량': None,
                '배송시 요구사항': None,
                '우편번호': None,
                '판매금액': None,  # Y열 판매금액 추가
                '추가구성': None,  # S열 추가구성 추가
                '배송비 금액': None  # AI열 배송비 금액 추가
            }
            
            for col in df.columns:
                col_str = str(col).strip()
                for key in required_columns.keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        required_columns[key] = col
                        print(f"✓ '{key}' 열을 찾았습니다: {col}")
            
            # Y열(판매금액)을 찾지 못한 경우 인덱스로 직접 접근 시도
            if required_columns['판매금액'] is None:
                # Y열은 25번째 열 (0-based로는 24, pandas는 0-based)
                if len(df.columns) > 24:
                    required_columns['판매금액'] = df.columns[24]
                    print(f"✓ '판매금액' 열을 인덱스로 찾았습니다: {df.columns[24]}")
                else:
                    print("⚠️ Y열(판매금액)을 찾을 수 없습니다. 금액 정보가 없을 수 있습니다.")
            
            # S열(추가구성)을 찾지 못한 경우 인덱스로 직접 접근 시도
            if required_columns['추가구성'] is None:
                # S열은 19번째 열 (0-based로는 18, pandas는 0-based)
                if len(df.columns) > 18:
                    required_columns['추가구성'] = df.columns[18]
                    print(f"✓ '추가구성' 열을 인덱스로 찾았습니다: {df.columns[18]}")
                else:
                    print("⚠️ S열(추가구성)을 찾을 수 없습니다. 추가구성 정보가 없을 수 있습니다.")
            
            # AI열(배송비 금액)을 찾지 못한 경우 인덱스로 직접 접근 시도
            if required_columns['배송비 금액'] is None:
                # AI열은 35번째 열 (0-based로는 34, pandas는 0-based)
                if len(df.columns) > 34:
                    required_columns['배송비 금액'] = df.columns[34]
                    print(f"✓ '배송비 금액' 열을 인덱스로 찾았습니다: {df.columns[34]}")
                else:
                    print("⚠️ AI열(배송비 금액)을 찾을 수 없습니다. 배송비 정보가 없을 수 있습니다.")
            
            # 필수 열이 모두 있는지 확인 (판매금액, 추가구성, 배송비 금액 열은 선택사항으로 처리)
            missing_columns = [key for key, value in required_columns.items() if value is None and key not in ['판매금액', '추가구성', '배송비 금액']]
            if missing_columns:
                print(f"❌ 다음 열을 찾을 수 없습니다: {', '.join(missing_columns)}")
                QMessageBox.warning(self, "오류", f"다음 열을 찾을 수 없습니다:\n{', '.join(missing_columns)}")
                return
            
            # 주문 정보 정리
            self.orders = {}
            
            # 주문번호별로 주문 정보 정리
            for _, row in df.iterrows():
                order_number = str(row[required_columns['주문번호']])
                if pd.isna(order_number) or order_number.strip() == '':
                    continue
                
                if order_number not in self.orders:
                    phone_value = row[required_columns['수령인 전화번호']]
                    
                    # 판매금액 가져오기 (Y열)
                    sale_amount = 0
                    if required_columns['판매금액'] is not None:
                        try:
                            sale_value = row[required_columns['판매금액']]
                            if not pd.isna(sale_value):
                                # 숫자로 변환 시도
                                if isinstance(sale_value, str):
                                    # 쉼표 제거 후 숫자 변환
                                    sale_value = sale_value.replace(',', '').strip()
                                sale_amount = float(sale_value)
                        except (ValueError, TypeError) as e:
                            print(f"⚠️ 판매금액 변환 실패: {sale_value}, 오류: {e}")
                            sale_amount = 0
                    
                    # 배송비 금액 가져오기 (AI열)
                    shipping_amount = 0
                    if required_columns['배송비 금액'] is not None:
                        try:
                            shipping_value = row[required_columns['배송비 금액']]
                            if not pd.isna(shipping_value):
                                # 숫자로 변환 시도
                                if isinstance(shipping_value, str):
                                    # 쉼표 제거 후 숫자 변환
                                    shipping_value = shipping_value.replace(',', '').strip()
                                shipping_amount = float(shipping_value)
                        except (ValueError, TypeError) as e:
                            print(f"⚠️ 배송비 금액 변환 실패: {shipping_value}, 오류: {e}")
                            shipping_amount = 0
                    
                    self.orders[order_number] = {
                        '수령인명': str(row[required_columns['수령인명']]),
                        '주소': str(row[required_columns['주소']]),
                        '수령인 전화번호': str(row[required_columns['수령인 전화번호']]) if not pd.isna(row[required_columns['수령인 전화번호']]) else '',
                        '수령인 휴대폰': str(row[required_columns['수령인 휴대폰']]) if not pd.isna(row[required_columns['수령인 휴대폰']]) else '',
                        '배송시 요구사항': str(row[required_columns['배송시 요구사항']]) if not pd.isna(row[required_columns['배송시 요구사항']]) else '',
                        '우편번호': str(row[required_columns['우편번호']]) if not pd.isna(row[required_columns['우편번호']]) else '',
                        '상품목록': [],
                        '판매금액': sale_amount,  # 주문별 판매금액 저장
                        '배송비 금액': shipping_amount  # 주문별 배송비 금액 저장
                    }
                
                # 상품 정보 추가
                product_name = str(row[required_columns['상품명']])
                raw_option = row[required_columns['옵션']]
                if pd.isna(raw_option):
                    option = '없음'
                else:
                    option = str(raw_option).strip()
                    if option.lower() in ('nan', ''):
                        option = '없음'
                quantity = int(row[required_columns['수량']]) if not pd.isna(row[required_columns['수량']]) else 1
                
                # 추가구성 정보 가져오기 (S열)
                additional_config = ''
                if required_columns['추가구성'] is not None:
                    raw_additional = row[required_columns['추가구성']]
                    if pd.notna(raw_additional):
                        additional_config = str(raw_additional).strip()
                        if additional_config.lower() in ('nan', ''):
                            additional_config = ''
                
                # 옵션ID로 상품코드 찾기
                # option_id = str(row[required_columns['옵션ID']]).strip()
                # product_code = product_code_map.get(option_id, '')  # 매칭되는 상품코드가 없으면 빈 문자열
                
                self.orders[order_number]['상품목록'].append({
                    '상품명': product_name,
                    '옵션': option,
                    '수량': quantity,
                    '추가구성': additional_config,
                    # '상품코드': product_code
                })
            
            # 같은 주문자(수령인명)로 주문 통합
            consolidated_orders = {}
            for order_number, info in self.orders.items():
                # 수령인명을 키로 사용하여 주문 통합
                customer_name = info['수령인명']
                
                if customer_name not in consolidated_orders:
                    consolidated_orders[customer_name] = {
                        '수령인명': customer_name,
                        '주문번호목록': [order_number],
                        '상품목록': info['상품목록'].copy(),
                        '총판매금액': info.get('판매금액', 0),
                        '총배송비금액': info.get('배송비 금액', 0)
                    }
                else:
                    # 기존 주문에 추가
                    consolidated_orders[customer_name]['주문번호목록'].append(order_number)
                    consolidated_orders[customer_name]['상품목록'].extend(info['상품목록'])
                    consolidated_orders[customer_name]['총판매금액'] += info.get('판매금액', 0)
                    consolidated_orders[customer_name]['총배송비금액'] += info.get('배송비 금액', 0)
            
            # 마크다운 형식으로 주문 정보 생성
            markdown_text = ""
            
            # 인덱스 값 다시 로드
            if hasattr(self.ui, 'lineEdit_idx_gmarket'):
                saved_index = self.ui.lineEdit_idx_gmarket.text().strip()
                if saved_index and saved_index.isdigit():
                    self.current_idx_gmarket = int(saved_index)
                else:
                    self.current_idx_gmarket = 1
                    self.ui.lineEdit_idx_gmarket.setText(str(self.current_idx_gmarket))
            
            for customer_name, info in consolidated_orders.items():
                # 총판매금액과 총배송비금액 합산
                total_sale_amount = info.get('총판매금액', 0)
                total_shipping_amount = info.get('총배송비금액', 0)
                total_amount = total_sale_amount + total_shipping_amount
                
                # 만원 단위로 포맷팅 (100원 단위 아래는 내림, 소수점 첫째 자리에서도 내림)
                # 예: 61000 -> 6.1만, 63820 -> 6.3만
                amount_in_100 = total_amount // 100  # 100원 단위로 내림
                amount_in_manwon = amount_in_100 / 100  # 만원 단위로 변환
                # 소수점 첫째 자리까지 표시 (둘째 자리에서 내림)
                amount_rounded = math.floor(amount_in_manwon * 10) / 10
                formatted_amount = f"{amount_rounded}만"
                
                markdown_text += f"[ ] {self.current_idx_gmarket}.{customer_name} - {formatted_amount}\n"
                self.update_gmarket_index()
                if hasattr(self.ui, 'lineEdit_idx_gmarket'):
                    self.ui.lineEdit_idx_gmarket.setText(str(self.current_idx_gmarket))
                
                # 상품 목록 표시
                for product in info['상품목록']:
                    product_name = product['상품명']
                    quantity = product['수량']
                    option = product['옵션']
                    additional_config = product.get('추가구성', '')
                    # product_code = product['상품코드'] 
                    product_code = 'PASS'
                   
                    # markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {product_name} ( 옵션 : {option} )\n"
                    markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {product_name} ( 옵션 : {option} )\n"
                    
                    # 추가구성이 있으면 표시
                    if additional_config and additional_config.strip():
                        markdown_text += f"  (추가구성 : {additional_config})\n"
                
                markdown_text += "\n"
            
            # plainTextEdit에 마크다운 텍스트 표시
            self.ui.plainTextEdit.setPlainText(markdown_text)
            
            # 클립보드에 자동 복사
            clipboard = QApplication.clipboard()
            clipboard.setText(markdown_text)
            self.statusBar().showMessage("주문 정보가 클립보드에 복사되었습니다.", 2000)
            
        except Exception as e:
            error_msg = str(e)
            print(f"❌ 엑셀 파일 처리 중 오류 발생: {error_msg}")
            if _is_likely_google_sheets_oauth_error(e):
                QMessageBox.critical(
                    self,
                    "오류",
                    "Google 스프레드시트 연동 중 오류가 발생했습니다.\n\n"
                    f"{error_msg}\n\n"
                    + self._oauth_error_dialog_hint(),
                )
            else:
                QMessageBox.critical(
                    self,
                    "오류",
                    f"엑셀 파일 처리 중 오류가 발생했습니다.\n\n{error_msg}",
                )

    def generate_work_order(self):
        """작업지시서 생성 버튼 클릭 시 실행되는 함수입니다."""
        if not self.selected_file_path:
            QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 선택해주세요.")
            return
            
        print("작업지시서 생성됨")
        self.statusBar().showMessage("작업지시서 생성 완료")
        # 여기에 실제 작업지시서 생성 로직이 추가될 예정입니다.

    def clear_list(self):
        """리스트 초기화 버튼 클릭 시 실행되는 함수입니다."""
        # 주문 정보 초기화
        self.orders = {}
        
        # 파일 관련 정보 초기화
        self.selected_file_path = None
        self.store_type = None
        
        # UI 요소 초기화
        self.ui.filePathLabel.setText("주문 정보가 없습니다. ( Ctrl + O )")
        self.ui.plainTextEdit.setPlainText("")
        
        # 로고 초기화
        self.ui.label_logo.clear()
        self.ui.label_logo.setText("image")
        
        # 송장 미리보기(주문·발송 탭) 텍스트 초기화
        self.ui.label_invoice.setText("송장 정보가 없습니다.")
        self.ui.label_generate_invoice.setText("생성 버튼을 누르세요.")
        self.ui.plainTextEdit_invoice.setPlainText("")
        
        # 상태바 메시지 업데이트
        self.statusBar().showMessage("모든 정보가 초기화되었습니다.")

    def show_about(self):
        """프로그램 정보를 보여주는 대화상자를 표시합니다."""
        QMessageBox.about(
            self,
            "프로그램 정보",
            "Easy Fulfill - 주문 처리 시스템\n\n"
            "Version 1.0.0\n"
            "Copyright © 2024\n\n"
            "오픈마켓 주문 처리를 위한 자동 생성 프로그램입니다."
        )
        self.statusBar().showMessage("프로그램 정보 표시됨")

    def open_file_location(self, file_path):
        """파일이 있는 폴더를 엽니다."""
        try:
            if sys.platform == 'win32':
                # Windows
                subprocess.run(['explorer', '/select,', str(file_path)])
            elif sys.platform == 'darwin':
                # macOS
                subprocess.run(['open', '-R', str(file_path)])
            else:
                # Linux
                subprocess.run(['xdg-open', str(Path(file_path).parent)])
            return True
        except Exception as e:
            print(f"! 폴더 열기 실패: {str(e)}")
            return False

    def export_invoice_excel(self):
        """송장 엑셀 파일을 생성합니다."""
        if not self.selected_file_path:
            QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 선택해주세요.")
            return

        try:
            # output 디렉토리 생성
            output_dir = Path("output")
            output_dir.mkdir(exist_ok=True)
            
            # 현재 시간을 파일명에 포함
            current_time = datetime.now().strftime("%Y%m%d%H%M%S")
            output_file = (output_dir / f"하이제니스 폼_{current_time}.xlsx").resolve()
            
            print(f"\n[송장 엑셀 생성 시작]")
            print(f"출력 파일: {output_file}")
            
            # 데이터프레임 생성을 위한 데이터 준비
            invoice_data = []
            
            if self.store_type == "naver":
                # 수취인 정보 기준으로 주문 통합 (시간 무관)
                consolidated_orders = {}
                
                # 주문 통합 처리
                for pattern, info in self.orders.items():
                    # 배송 방법이 '택배,등기,소포'가 아닌 경우 건너뛰기
                    if info['배송방법'] != '택배,등기,소포':
                        print(f"주문 {pattern} 건너뛰기: 배송방법 '{info['배송방법']}'")
                        continue
                        
                    # 수취인명, 연락처, 주문번호로 키 생성
                    key = (info['수취인명'], info['수취인연락처1'], pattern)
                    
                    if key not in consolidated_orders:
                        # 우편번호를 5자리로 고정 (앞에 0 채우기)
                        zipcode = str(info['우편번호']).strip()
                        if zipcode.isdigit():
                            zipcode = zipcode.zfill(5)
                        
                        consolidated_orders[key] = {
                            '주문번호': pattern,
                            '수취인명': info['수취인명'],
                            '수취인연락처1': info['수취인연락처1'],
                            '통합배송지': info['통합배송지'],
                            '배송메세지': info['배송메세지'],
                            '우편번호': zipcode,
                            '상품목록': info['상품목록'].copy()
                        }
                    else:
                        # 기존 주문에 상품 정보 추가
                        consolidated_orders[key]['상품목록'].extend(info['상품목록'])
                        # 배송메시지가 있는 경우에만 업데이트
                        if info['배송메세지'].strip():
                            consolidated_orders[key]['배송메세지'] = info['배송메세지']

                # 통합된 주문 정보로 송장 데이터 생성
                for info in consolidated_orders.values():
                    delivery_msg = info.get('배송메세지', '')
                    if pd.isna(delivery_msg) or str(delivery_msg).lower() == 'nan':
                        delivery_msg = ''
                    
                    # 상품 정보를 하나의 문자열로 결합
                    product_info = []
                    for product in info['상품목록']:
                        product_str = f"{product['상품명']} (옵션: {product['옵션']}) - {product['수량']}개"
                        product_info.append(product_str)
                    
                    # 모든 상품 정보를 줄바꿈으로 구분하여 하나의 문자열로 결합
                    combined_product_info = '\n'.join(product_info)
                        
                    invoice_data.append({
                        '주문번호': info['주문번호'],
                        '고객주문처명': '',
                        '수취인명': info['수취인명'],
                        '우편번호': info['우편번호'],
                        '수취인 주소': info['통합배송지'],
                        '수취인 전화번호': info['수취인연락처1'],
                        '수취인 이동통신': info['수취인연락처1'],
                        '상품명': combined_product_info,
                        '상품모델': '전자제품',
                        '배송메세지': delivery_msg,
                        '비고': ''
                    })
            elif self.store_type == "coupang":
                # 쿠팡 스토어 처리
                print("\n[쿠팡 스토어 데이터 구조 확인]")
                
                # 주문 통합 처리
                consolidated_orders = {}
                
                for order_number, info in self.orders.items():
                    print(f"[주문 처리 시작] 주문번호: {order_number}")
                    
                    # 수취인명, 연락처, 주문번호로 키 생성
                    key = (info['수취인이름'], info['수취인전화번호'], order_number)
                    
                    if key not in consolidated_orders:
                        # 우편번호 처리
                        zipcode = str(info.get('우편번호', '')).strip()
                        if zipcode.isdigit():
                            zipcode = zipcode.zfill(5)
                        
                        consolidated_orders[key] = {
                            '주문번호': order_number,
                            '수취인이름': info['수취인이름'],
                            '수취인전화번호': info['수취인전화번호'],
                            '수취인주소': info['수취인주소'],
                            '배송메세지': info.get('배송메세지', ''),
                            '우편번호': zipcode,
                            '상품목록': info['상품목록'].copy()
                        }
                    else:
                        # 기존 주문에 상품 정보 추가
                        consolidated_orders[key]['상품목록'].extend(info['상품목록'])
                        # 배송메시지가 있는 경우에만 업데이트
                        if info.get('배송메세지', '').strip():
                            consolidated_orders[key]['배송메세지'] = info['배송메세지']
                
                # 통합된 주문 정보로 송장 데이터 생성
                for info in consolidated_orders.values():
                    delivery_msg = info.get('배송메세지', '')
                    if pd.isna(delivery_msg) or str(delivery_msg).lower() == 'nan':
                        delivery_msg = ''
                    
                    # 상품 정보를 하나의 문자열로 결합
                    product_info = []
                    for product in info['상품목록']:
                        product_str = f"{product['상품명']} (옵션: {product['옵션']}) - {product['수량']}개"
                        product_info.append(product_str)
                    
                    # 모든 상품 정보를 줄바꿈으로 구분하여 하나의 문자열로 결합
                    combined_product_info = '\n'.join(product_info)
                    
                    invoice_data.append({
                        '주문번호': info['주문번호'],
                        '고객주문처명': '',
                        '수취인명': info['수취인이름'],
                        '우편번호': info['우편번호'],
                        '수취인 주소': info['수취인주소'],
                        '수취인 전화번호': info['수취인전화번호'],
                        '수취인 이동통신': info['수취인전화번호'],
                        '상품명': combined_product_info,
                        '상품모델': '전자제품',
                        '배송메세지': delivery_msg,
                        '비고': ''
                    })
            elif self.store_type == "gmarket":
                # 지마켓 스토어 처리
                print("\n[지마켓 스토어 데이터 구조 확인]")
                
                # 주문 통합 처리
                consolidated_orders = {}
                
                for order_number, info in self.orders.items():
                    print(f"[주문 처리 시작] 주문번호: {order_number}")
                    
                    # 수취인명, 연락처, 주문번호로 키 생성
                    key = (info['수령인명'], info['수령인 전화번호'], order_number)
                    
                    if key not in consolidated_orders:
                        # 우편번호 처리
                        zipcode = str(info.get('우편번호', '')).strip()
                        if zipcode.isdigit():
                            zipcode = zipcode.zfill(5)
                        
                        consolidated_orders[key] = {
                            '주문번호': order_number,
                            '수령인명': info['수령인명'],
                            '수취인 이동통신': info['수령인 휴대폰'],
                            '수취인 전화번호': info['수령인 전화번호'],
                            '주소': info['주소'],
                            '배송시 요구사항': info.get('배송시 요구사항', ''),
                            '우편번호': zipcode,
                            '상품목록': info['상품목록'].copy()
                        }
                    else:
                        # 기존 주문에 상품 정보 추가
                        consolidated_orders[key]['상품목록'].extend(info['상품목록'])
                        # 배송메시지가 있는 경우에만 업데이트
                        if info.get('배송시 요구사항', '').strip():
                            consolidated_orders[key]['배송시 요구사항'] = info['배송시 요구사항']
                
                # 통합된 주문 정보로 송장 데이터 생성
                for info in consolidated_orders.values():
                    delivery_msg = info.get('배송시 요구사항', '')
                    if pd.isna(delivery_msg) or str(delivery_msg).lower() == 'nan':
                        delivery_msg = ''
                    
                    # 상품 정보를 하나의 문자열로 결합
                    product_info = []
                    for product in info['상품목록']:
                        product_str = f"{product['상품명']} (옵션: {product['옵션']}) - {product['수량']}개"
                        product_info.append(product_str)
                    
                    # 모든 상품 정보를 줄바꿈으로 구분하여 하나의 문자열로 결합
                    combined_product_info = '\n'.join(product_info)
                    
                    invoice_data.append({
                        '주문번호': info['주문번호'],
                        '고객주문처명': '',
                        '수취인명': info['수령인명'],
                        '우편번호': info['우편번호'],
                        '수취인 주소': info['주소'],
                        '수취인 전화번호': info['수취인 전화번호'],
                        '수취인 이동통신': info['수취인 이동통신'],
                        '상품명': combined_product_info,
                        '상품모델': '전자제품',
                        '배송메세지': delivery_msg,
                        '비고': ''
                    })          
                
            # 데이터프레임 생성
            df_invoice = pd.DataFrame(invoice_data)
            
            # 'nan' 값을 빈 문자열로 변환
            df_invoice['배송메세지'] = df_invoice['배송메세지'].fillna('')
            
            # 열 순서 지정
            columns = [
                '주문번호',
                '고객주문처명',
                '수취인명',
                '우편번호',
                '수취인 주소',
                '수취인 전화번호',
                '수취인 이동통신',
                '상품명',
                '상품모델',
                '배송메세지',
                '비고'
            ]
            df_invoice = df_invoice[columns]
            
            # 엑셀 파일 저장 (with xlsxwriter)
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df_invoice.to_excel(writer, index=False, sheet_name='Sheet1')
                
                # 워크시트와 워크북 객체 가져오기
                worksheet = writer.sheets['Sheet1']
                workbook = writer.book
                
                # 가운데 정렬을 위한 셀 포맷 설정
                center_format = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                # 헤더 포맷 설정 (가운데 정렬 + 굵게)
                header_format = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter',
                    'bold': True
                })
                
                # 열 너비 자동 조정 및 포맷 적용
                for idx, col in enumerate(df_invoice.columns):
                    # 열 이름의 길이와 데이터의 최대 길이 계산
                    max_length = max(
                        df_invoice[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    # 한글은 2배의 너비가 필요하므로 조정
                    adjusted_width = max_length * 2 if any('\u3131' <= c <= '\u318E' or '\uAC00' <= c <= '\uD7A3' for c in str(col)) else max_length
                    worksheet.set_column(idx, idx, adjusted_width + 2, center_format)
                
                # 헤더에 포맷 적용
                for col_num, value in enumerate(df_invoice.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # 전체 행 높이 조정
                worksheet.set_default_row(20)

            print(f"✓ 송장 엑셀 파일이 생성되었습니다.")
            print(f"  - 파일 위치: {output_file}")
            print(f"  - 행 수: {len(df_invoice)}")

            self.statusBar().showMessage(f"송장 엑셀 파일 생성 완료: {output_file}")
            
            # 성공 메시지 표시 (커스텀 버튼 포함)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("완료")
            msg.setText("송장 엑셀 파일이 생성되었습니다.")
            
            # 버튼 추가
            open_location_button = msg.addButton("폴더 열기", QMessageBox.ActionRole)
            open_file_button = msg.addButton("엑셀 열기", QMessageBox.ActionRole)
            close_button = msg.addButton("닫기", QMessageBox.RejectRole)
            
            msg.setDefaultButton(close_button)
            
            # 메시지 박스 표시
            clicked_button = msg.exec()
            
            # 버튼 클릭 처리
            if msg.clickedButton() == open_location_button:
                if not self.open_file_location(output_file):
                    QMessageBox.warning(self, "오류", "파일 위치를 열 수 없습니다.")
            elif msg.clickedButton() == open_file_button:
                if not self.open_file_with_default_app(output_file):
                    QMessageBox.warning(self, "오류", "파일을 열 수 없습니다.\n엑셀이 설치되어 있는지 확인해주세요.")

        except Exception as e:
            error_msg = str(e)
            print(f"❌ 송장 엑셀 파일 생성 중 오류 발생: {error_msg}")
            QMessageBox.critical(
                self,
                "오류",
                f"송장 엑셀 파일 생성 중 오류가 발생했습니다.\n\n{error_msg}"
            )
            
    def load_invoice_file(self):
        """엑셀 파일을 선택하는 다이얼로그를 표시합니다."""
        # 다운로드 폴더 경로 설정
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        
        file_path, _ = QFileDialog.getOpenFileName(
        self,
        "엑셀 파일 선택",
        downloads_path,
        "Excel Files (*.xlsx *.xls)"
        )                
        if file_path:
            try:
                # plainTextEdit_invoice 초기화
                self.ui.plainTextEdit_invoice.setPlainText("")
                
                print(f"\n[파일 선택됨] 경로: {file_path}")
                filename = os.path.basename(file_path)
                
                # 파일명 검증
                if not filename.endswith('.xlsx'):
                    self.is_invoice_file_valid = False
                    QMessageBox.warning(self, "오류", "파일 확장자가 .xlsx가 아닙니다.")
                    return
                    
                # 파일명에서 확장자를 제외한 부분 추출
                base_name = filename[:-5]  # .xlsx를 제외
                
                # 13자리 숫자 검증
                if len(base_name) < 13 or not base_name[:13].isdigit():
                    self.is_invoice_file_valid = False
                    QMessageBox.warning(
                        self,
                        "오류",
                        "파일명이 올바른 형식이 아닙니다.\n"
                        "파일명은 13자리 숫자로 시작해야 합니다.\n"
                        "예: 1744778498617.xlsx 또는 1744171918081 (1).xlsx"
                    )
                    return
                
                self.invoice_file_path = file_path
                self.ui.label_invoice.setText(filename)
                
                # 엑셀 파일 읽기
                df = pd.read_excel(file_path, header=6)  # 7번째 행을 헤더로 사용
                
                # 필요한 열 찾기
                required_columns = {
                    '등기번호': None,
                    '수취인명': None,
                    '수취인 이동통신': None,
                    '수취인상세주소': None
                }
                
                # 열 이름 매칭
                for col in df.columns:
                    col_str = str(col).strip()
                    for key in required_columns.keys():
                        # if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        if col_str.replace(' ', '') == key.replace(' ', ''):  # 공백 제거 후 비교
                            required_columns[key] = col
                            print(f"✓ '{key}' 열을 찾았습니다: {col}")
                
                # 필수 열이 모두 있는지 확인
                missing_columns = [key for key, value in required_columns.items() if value is None]
                if missing_columns:
                    self.is_invoice_file_valid = False
                    print(f"❌ 다음 열을 찾을 수 없습니다: {', '.join(missing_columns)}")
                    QMessageBox.warning(self, "오류", f"다음 열을 찾을 수 없습니다:\n{', '.join(missing_columns)}")
                    return
                
                # 마크다운 형식으로 주문 정보 생성
                markdown_text = ""
                                
                # 각 행을 순회하며 주문 정보 생성
                for idx, (_, row) in enumerate(df.iterrows(), start=1):
                    # 빈 행 건너뛰기
                    if pd.isna(row[required_columns['등기번호']]):
                        continue

                    # 주문 정보 추가 (인덱스 포함)
                    markdown_text += f"{idx}. {row[required_columns['수취인명']]}\n"
                    markdown_text += f"  송장번호: {row[required_columns['등기번호']]}\n"
                    markdown_text += f"  이동통신: {row[required_columns['수취인 이동통신']]}\n"
                    markdown_text += f"  주소: {row[required_columns['수취인상세주소']]}\n\n"

                
                # plainTextEdit_invoice 마크다운 텍스트 표시
                self.ui.plainTextEdit_invoice.setPlainText(markdown_text)
                
                # 클립보드에 자동 복사
                clipboard = QApplication.clipboard()
                clipboard.setText(markdown_text)
                
                self.statusBar().showMessage(f"송장 파일 선택됨: {filename}")
                print("✓ 송장 파일이 성공적으로 선택되었습니다.")
                self.is_invoice_file_valid = True  # 파일 처리 성공 시 플래그 설정

                if hasattr(self.ui, "checkBox_invoice_load_auto_generate"):
                    if self.ui.checkBox_invoice_load_auto_generate.isChecked():
                        self.generate_invoice_file()
                
            except Exception as e:
                self.is_invoice_file_valid = False
                error_msg = str(e)
                print(f"❌ 엑셀 파일 처리 중 오류 발생: {error_msg}")
                # 사용자에게는 간단한 메시지만 표시
                QMessageBox.warning(
                    self,
                    "오류",
                    "올바르지 않은 파일 형식입니다."
                )
                
    def generate_invoice_file(self):
        """일괄 발송 파일 생성 메인 메소드"""
        if not self.is_order_file_valid or not self.is_invoice_file_valid:
            QMessageBox.warning(self, "경고", "먼저 유효한 주문서와 송장 파일을 선택해주세요.")
            return
            
        try:
            # 임시 디렉토리 생성
            temp_dir = Path("temp")
            temp_dir.mkdir(exist_ok=True)
            
            # 파일 복사
            temp_order_file = temp_dir / "temp_order.xlsx"
            temp_invoice_file = temp_dir / "temp_invoice.xlsx"
            shutil.copy2(self.selected_file_path, temp_order_file)
            shutil.copy2(self.invoice_file_path, temp_invoice_file)
            
            # 스토어 타입에 따른 처리 분기
            if self.store_type == "naver":
                self._process_naver_invoice(temp_order_file, temp_invoice_file)
            elif self.store_type == "coupang":
                self._process_coupang_invoice(temp_order_file, temp_invoice_file)
            
            # 임시 파일 정리
            temp_order_file.unlink()
            temp_invoice_file.unlink()
            temp_dir.rmdir()
            
        except Exception as e:
            error_msg = str(e)
            print(f"❌ 일괄 발송 파일 생성 중 오류 발생: {error_msg}")
            QMessageBox.critical(
                self,
                "오류",
                f"일괄 발송 파일 생성 중 오류가 발생했습니다.\n\n{error_msg}"
            )
            
    def _process_naver_invoice(self, temp_order_file, temp_invoice_file):
        """네이버 스토어 일괄 발송 파일 처리"""
        print("\n[네이버 스토어 일괄 발송 파일 처리 시작]")
        
        try:
            # 1. 주문서 파일 복호화
            password = "1234"
            decrypted_order_file = Path("temp") / "decrypted_order.xlsx"
            
            with open(temp_order_file, 'rb') as file:
                office_file = msoffcrypto.OfficeFile(file)
                office_file.load_key(password=password)
                with open(decrypted_order_file, 'wb') as output_file:
                    office_file.decrypt(output_file)
            
            # 2. 주문서 데이터프레임 생성
            order_df = pd.read_excel(decrypted_order_file, sheet_name='발주발송관리', header=None)
            order_df = order_df.drop(0).reset_index(drop=True)
            
            # 열 이름 설정
            new_columns = []
            for i in range(len(order_df.columns)):
                col_name = str(order_df.iloc[0, i]).strip()
                new_columns.append(col_name if col_name and not col_name.startswith('Unnamed') else f'Column_{i}')
            
            order_df.columns = new_columns
            order_df = order_df.drop(0).reset_index(drop=True)
            
            # 운송장번호 컬럼 추가 (B열)
            order_df['운송장번호'] = ''
            
            # 3. 배송방법 복사 처리
            print("\n[배송방법 복사 처리]")
            
            # 배송방법 관련 컬럼 찾기
            delivery_method_col = None  # E열: 배송방법
            delivery_request_col = None  # G열: 배송방법(구매자 요청)
            
            for col in order_df.columns:
                col_str = str(col).strip()
                if col_str == '배송방법':
                    delivery_method_col = col
                    print(f"✓ 배송방법 컬럼 찾음: {col}")
                elif col_str == '배송방법(구매자 요청)':
                    delivery_request_col = col
                    print(f"✓ 배송방법(구매자 요청) 컬럼 찾음: {col}")
            
            # 배송방법 복사 처리
            if delivery_method_col and delivery_request_col:
                copy_count = 0
                delete_count = 0
                rows_to_delete = []  # 삭제할 행 인덱스 저장
                
                for idx, row in order_df.iterrows():
                    requested_method = str(row[delivery_request_col]).strip()
                    
                    # NaN 값 처리
                    if pd.isna(row[delivery_request_col]) or requested_method == 'nan':
                        requested_method = ''
                                        
                    if requested_method:
                        # G열의 값이 '택배,등기,소포'인 경우 E열에 무조건 복사
                        if requested_method == '택배,등기,소포':                            
                            order_df.loc[idx, delivery_method_col] = requested_method
                            copy_count += 1
                        else:
                            # G열의 값이 '택배,등기,소포'가 아닌 경우 해당 행 삭제
                            print(f"행 {idx+1}: 배송방법이 '택배,등기,소포'가 아님 - 행 삭제")
                            print(f"  - 요청(G열): '{requested_method}' (택배,등기,소포 아님)")
                            rows_to_delete.append(idx)
                            delete_count += 1
                
                # 삭제할 행들을 역순으로 삭제 (인덱스가 변경되지 않도록)
                for idx in reversed(rows_to_delete):
                    order_df = order_df.drop(idx).reset_index(drop=True)
                
                print(f"✓ 총 {copy_count}개의 배송방법이 복사되었습니다.")
                print(f"✓ 총 {delete_count}개의 행이 삭제되었습니다.")
            else:
                print("! 배송방법 관련 컬럼을 찾을 수 없습니다.")
                if not delivery_method_col:
                    print("  - 배송방법 컬럼 없음")
                if not delivery_request_col:
                    print("  - 배송방법(구매자 요청) 컬럼 없음")
            
            # 4. 송장 파일 읽기
            invoice_df = pd.read_excel(temp_invoice_file, header=6)
            
            # 5. 열 매핑 설정
            column_mapping = {
                'order': {
                    '수취인명': None,
                    '수취인연락처1': None,
                    '통합배송지': None,
                    '주문번호': None
                },
                'invoice': {
                    '등기번호': None,
                    '수취인명': None,
                    '수취인 이동통신': None,
                    '수취인상세주소': None,
                    '고객주문번호': None
                }
            }
            
            # 열 이름 매핑
            for col in order_df.columns:
                col_str = str(col).strip()
                for key in column_mapping['order'].keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        column_mapping['order'][key] = col
            
            for col in invoice_df.columns:
                col_str = str(col).strip()
                for key in column_mapping['invoice'].keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        column_mapping['invoice'][key] = col

            def _normalize_value(value):
                if pd.isna(value):
                    return ''
                return str(value).strip()

            def _normalize_digits(value):
                if pd.isna(value):
                    return ''
                return ''.join(ch for ch in str(value) if ch.isdigit())

            order_number_series = None
            if column_mapping['order']['주문번호']:
                order_number_series = order_df[column_mapping['order']['주문번호']].apply(_normalize_digits).str[:13]
            
            # 6. 매칭된 주문 정보 출력 및 운송장번호 업데이트
            print("\n[매칭된 주문 정보]")
            matched_count = 0
            
            for idx, invoice_row in invoice_df.iterrows():
                invoice_name = _normalize_value(invoice_row[column_mapping['invoice']['수취인명']])
                invoice_phone = _normalize_value(invoice_row[column_mapping['invoice']['수취인 이동통신']])
                invoice_number = _normalize_value(invoice_row[column_mapping['invoice']['등기번호']])
                invoice_order_number = ''
                if column_mapping['invoice']['고객주문번호']:
                    invoice_order_number = _normalize_digits(invoice_row[column_mapping['invoice']['고객주문번호']])[:13]

                matching_rows = order_df.iloc[0:0]
                if order_number_series is not None and invoice_order_number:
                    matching_rows = order_df[
                        order_number_series == invoice_order_number
                    ]
                
                if not matching_rows.empty:
                    matched_count += 1
                    print(f"\n[매칭된 주문 {matched_count}]")
                    print(f"수취인명: {invoice_name}")
                    print(f"전화번호: {invoice_phone}")
                    if invoice_order_number:
                        print(f"고객주문번호: {invoice_order_number}")
                    print(f"송장번호: {invoice_number}")
                    print(f"주소: {matching_rows[column_mapping['order']['통합배송지']].iloc[0]}")
                    
                    # 운송장번호 업데이트
                    order_df.loc[matching_rows.index, '송장번호'] = invoice_number
                    print(f"✓ 송장번호 업데이트 완료")
            
            print(f"\n✓ 총 {matched_count}개의 주문이 매칭되었습니다.")
            
            # 7. 결과 파일 저장
            self._save_invoice_file(order_df)
            
            # 8. 임시 파일 정리
            decrypted_order_file.unlink()
            
        except Exception as e:
            error_msg = str(e)
            traceback.print_exc()
            print(f"❌ 일괄 발송 파일 처리 중 오류 발생: {error_msg}")
            raise Exception(f"일괄 발송 파일 처리 중 오류가 발생했습니다: {error_msg}")

    def _process_coupang_invoice(self, temp_order_file, temp_invoice_file):
        """쿠팡 스토어 일괄 발송 파일 처리"""
        print("\n[쿠팡 스토어 일괄 발송 파일 처리 시작]")
        
        try:
            # 1. 쿠팡 주문서 파일 읽기
            print("\n[쿠팡 주문서 파일 읽기]")
            order_df = pd.read_excel(temp_order_file)
            
            # 묶음배송번호, 주문번호를 문자열로 변환하여 정확한 값 유지
            if '묶음배송번호' in order_df.columns:
                order_df['묶음배송번호'] = order_df['묶음배송번호'].astype(str)
            if '주문번호' in order_df.columns:
                order_df['주문번호'] = order_df['주문번호'].astype(str)
            
            # 분리배송 열 추가 (모든 값을 'N'으로 설정)
            order_df['분리배송 Y/N'] = 'N'
            
            # 데이터프레임 정보 출력
            print("\n[주문서 데이터프레임 정보]")
            print(f"행 수: {len(order_df)}")
            print(f"열 수: {len(order_df.columns)}")
            print("열 이름:")
            for col in order_df.columns:
                print(f"- {col}")
            
            # 필요한 컬럼 확인
            required_order_columns = {
                '수취인이름': None,
                '수취인전화번호': None,
                '수취인 주소': None
            }
            
            # 컬럼 매핑
            for col in order_df.columns:
                col_str = str(col).strip()
                for key in required_order_columns.keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        required_order_columns[key] = col
                        print(f"✓ 주문서 '{key}' 열 찾음: {col}")
            
            # 필수 컬럼 확인
            missing_columns = [key for key, value in required_order_columns.items() if value is None]
            if missing_columns:
                raise ValueError(f"주문서에서 다음 컬럼을 찾을 수 없습니다: {', '.join(missing_columns)}")
            
            # 샘플 데이터 출력
            print("\n[주문서 샘플 데이터]")
            sample_data = order_df[[required_order_columns['수취인이름'], 
                                  required_order_columns['수취인전화번호']]].head()
            print(sample_data)
            
            # 2. 우체국 송장 파일 읽기
            print("\n[우체국 송장 파일 읽기]")
            invoice_df = pd.read_excel(temp_invoice_file, header=6)
            
            # 데이터프레임 정보 출력
            print("\n[송장서 데이터프레임 정보]")
            print(f"행 수: {len(invoice_df)}")
            print(f"열 수: {len(invoice_df.columns)}")
            print("열 이름:")
            for col in invoice_df.columns:
                print(f"- {col}")
            
            # 필요한 컬럼 확인
            required_invoice_columns = {
                '등기번호': None,
                '수취인명': None,
                '수취인 전화번호': None
            }
            
            # 컬럼 매핑
            for col in invoice_df.columns:
                col_str = str(col).strip()
                for key in required_invoice_columns.keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        required_invoice_columns[key] = col
                        print(f"✓ 송장서 '{key}' 열 찾음: {col}")
            
            # 필수 컬럼 확인
            missing_columns = [key for key, value in required_invoice_columns.items() if value is None]
            if missing_columns:
                raise ValueError(f"송장서에서 다음 컬럼을 찾을 수 없습니다: {', '.join(missing_columns)}")
            
            # 3. 운송장번호 매칭 및 채우기
            print("\n[운송장번호 매칭 및 채우기]")
            print(
                "[참고] 매칭은 수취인명과 전화번호가 송장·주문서에서 문자열이 동일할 때만 됩니다. "
                "이름만으로 맞추지 않는 이유는 동명이인·같은 이름의 서로 다른 주문이 한 배치에 있을 때 "
                "송장이 엉키면 오배송·CS 위험이 크기 때문입니다."
            )
            
            def _coupang_norm_str(val):
                if pd.isna(val):
                    return ''
                s = str(val).strip()
                if s.lower() == 'nan':
                    return ''
                return s
            
            def _coupang_explain_invoice_unmatched(odf, inv_name, inv_phone):
                nc = required_order_columns['수취인이름']
                pc = required_order_columns['수취인전화번호']
                reasons = []
                if not inv_name:
                    reasons.append("송장 수취인명이 비어 있음")
                if not inv_phone:
                    reasons.append("송장 수취인 전화번호가 비어 있음")
                on = odf[nc].map(_coupang_norm_str)
                same_name = odf[on == inv_name]
                if same_name.empty:
                    reasons.append(
                        "주문서에 동일한 수취인명이 없음 "
                        "(철자·띄어쓰기·괄호·별칭 등 표기 차이 가능)"
                    )
                else:
                    ophones = sorted(
                        {p for p in same_name[pc].map(_coupang_norm_str).tolist() if p}
                    )
                    reasons.append(
                        f"이름은 주문서와 일치하나 전화번호 불일치 "
                        f"(송장: {inv_phone!r}, 주문서: {ophones})"
                    )
                return " / ".join(reasons)
            
            # 운송장번호 컬럼 추가
            order_df['운송장번호'] = ''
            
            # 매칭 카운터
            matched_count = 0
            unmatched_invoice = 0
            
            # 각 송장 행에 대해 매칭 시도
            for idx, invoice_row in invoice_df.iterrows():
                invoice_number = str(invoice_row[required_invoice_columns['등기번호']])
                invoice_name = _coupang_norm_str(invoice_row[required_invoice_columns['수취인명']])
                invoice_phone = _coupang_norm_str(invoice_row[required_invoice_columns['수취인 전화번호']])
                
                # 수취인명과 전화번호로 매칭
                matching_rows = order_df[
                    (order_df[required_order_columns['수취인이름']].map(_coupang_norm_str) == invoice_name) &
                    (order_df[required_order_columns['수취인전화번호']].map(_coupang_norm_str) == invoice_phone)
                ]
                
                if not matching_rows.empty:
                    matched_count += 1
                    print(f"\n[매칭된 주문 {matched_count}]")
                    print(f"수취인명: {invoice_name}")
                    print(f"전화번호: {invoice_phone}")
                    print(f"송장번호: {invoice_number}")
                    print(f"주소: {matching_rows[required_order_columns['수취인 주소']].iloc[0]}")
                    
                    order_df.loc[matching_rows.index, '운송장번호'] = invoice_number
                    print(f"✓ 송장번호 업데이트 완료")
                else:
                    unmatched_invoice += 1
                    why = _coupang_explain_invoice_unmatched(order_df, invoice_name, invoice_phone)
                    print(f"\n[매칭 실패 — 송장 행 {unmatched_invoice}]")
                    print(f"수취인명: {invoice_name or '(비어 있음)'}")
                    print(f"전화번호: {invoice_phone or '(비어 있음)'}")
                    print(f"등기번호: {invoice_number}")
                    print(f"사유: {why}")
            
            print(f"\n✓ 총 {matched_count}개의 주문이 매칭되었습니다.")
            if unmatched_invoice:
                print(f"⚠ 송장 {unmatched_invoice}건은 주문서와 맞는 조합이 없어 운송장번호를 넣지 못했습니다.")
            
            blank_tr = order_df['운송장번호'].map(_coupang_norm_str) == ''
            if blank_tr.any():
                print("\n[운송장번호가 비어 있는 주문 (위 매칭 실패와 대응되는 경우가 많음)]")
                for _, orow in order_df[blank_tr].iterrows():
                    oname = _coupang_norm_str(orow[required_order_columns['수취인이름']])
                    ophone = _coupang_norm_str(orow[required_order_columns['수취인전화번호']])
                    print(f"  - {oname} / {ophone}")
            
            # 4. 결과 파일 저장
            self._save_invoice_file(order_df)
            
        except Exception as e:
            error_msg = str(e)
            traceback.print_exc()
            print(f"❌ 일괄 발송 파일 처리 중 오류 발생: {error_msg}")
            raise Exception(f"일괄 발송 파일 처리 중 오류가 발생했습니다: {error_msg}")

    def _save_invoice_file(self, order_df):
        """일괄 발송 파일 저장"""
        # 결과 파일 저장
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        # 스토어 타입에 따라 다른 파일명 사용
        if self.store_type == "naver":
            output_file = output_dir / f"일괄발송_네이버_{current_time}.xlsx"
        elif self.store_type == "coupang":
            output_file = output_dir / f"일괄발송_쿠팡_{current_time}.xlsx"
        else:
            output_file = output_dir / f"일괄발송_{current_time}.xlsx"
        
        # 택배사 정보 추가
        if self.store_type == "naver":
            order_df['택배사'] = '우체국택배'
        elif self.store_type == "coupang":
            order_df['택배사'] = '우체국'

        order_df = order_df.copy()
        order_df.columns = [
            f"Column_{idx}" if pd.isna(col) else str(col).strip()
            for idx, col in enumerate(order_df.columns)
        ]

        def _display_width(value):
            if pd.isna(value):
                return 0
            return len(str(value))
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            order_df.to_excel(writer, index=False, sheet_name='발송처리')
            
            worksheet = writer.sheets['발송처리']
            workbook = writer.book
            
            # 숫자 형식 설정 (텍스트로 처리)
            text_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'num_format': '@'  # 텍스트 형식으로 저장
            })
            
            # 일반 셀 포맷 설정
            center_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter'
            })
            
            header_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'bold': True
            })
            
            # 열 너비 및 포맷 적용
            for idx, col in enumerate(order_df.columns):
                column_values = order_df.iloc[:, idx]
                max_length = max(
                    column_values.map(_display_width).max(),
                    _display_width(col)
                )
                adjusted_width = max_length * 2 if any('\u3131' <= c <= '\u318E' or '\uAC00' <= c <= '\uD7A3' for c in str(col)) else max_length
                
                # 묶음배송번호와 주문번호 열은 텍스트 형식으로 설정
                if col in ['묶음배송번호', '주문번호']:
                    worksheet.set_column(idx, idx, adjusted_width + 2, text_format)
                else:
                    worksheet.set_column(idx, idx, adjusted_width + 2, center_format)
            
            for col_num, value in enumerate(order_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            worksheet.set_default_row(20)
        
        # 주문서 목록 초기화
        self.clear_list()
        
        # 성공 메시지 표시
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("완료")
        msg.setText("일괄 발송 파일이 생성되었습니다.")
        
        open_location_button = msg.addButton("폴더 열기", QMessageBox.ActionRole)
        open_file_button = msg.addButton("엑셀 열기", QMessageBox.ActionRole)
        close_button = msg.addButton("닫기", QMessageBox.RejectRole)
        
        msg.setDefaultButton(close_button)
        msg.exec()
        
        if msg.clickedButton() == open_location_button:
            if not self.open_file_location(output_file):
                QMessageBox.warning(self, "오류", "파일 위치를 열 수 없습니다.")
        elif msg.clickedButton() == open_file_button:
            if not self.open_file_with_default_app(output_file):
                QMessageBox.warning(self, "오류", "파일을 열 수 없습니다.\n엑셀이 설치되어 있는지 확인해주세요.")

    def select_product_file(self):
        """상품 파일을 선택하는 다이얼로그를 표시합니다."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "상품 파일 선택",
            "input",
            "CSV Files (*.csv)"
        )
        
        if file_path:
            print(f"\n[상품 파일 선택됨] 경로: {file_path}")
            filename = os.path.basename(file_path)
            self.product_file_path = file_path
            self.ui.productFileLabel.setText(filename)
            self.statusBar().showMessage(f"상품 파일 선택됨: {filename}")
            print("✓ 상품 파일이 성공적으로 선택되었습니다.")

    def categorize_products(self):
        """선택된 상품 파일을 분류합니다."""
        if not hasattr(self, 'product_file_path'):
            QMessageBox.warning(self, "경고", "먼저 상품 파일을 선택해주세요.")
            return

        try:
            print(f"\n[상품 분류 시작]")
            
            # CSV 파일 읽기
            df = pd.read_csv(self.product_file_path)
            print(f"CSV 파일 열 목록: {df.columns.tolist()}")
            
            # 대표이미지 URL 열이 있는지 확인
            has_image_column = '대표이미지 URL' in df.columns
            print(f"대표이미지 URL 열 존재: {has_image_column}")
            
            if has_image_column:
                print(f"대표이미지 URL 열 데이터 샘플:")
                print(df[['상품명', '대표이미지 URL']].head())
            
            product_names = df['상품명'].dropna().unique()
            
            # 카테고리 키워드 정의
            category_keywords = {
                "MCU/개발보드": ["STM32", "ESP32", "아두이노", "라즈베리", "개발보드", "KIT"],
                "모터": ["모터", "서보", "BLDC", "코어리스", "JGB25"],
                "센서": ["센서", "엔코더", "AS5600", "AS5048", "자이로", "MPU"],
                "케이블": ["케이블", "USB", "연결선"],
                "통신 모듈": ["I2C", "SPI", "CAN", "RS485", "UART", "이더넷"],
                "디스플레이": ["OLED", "디스플레이", "LCD", "스크린"],
                "구동부품": ["바퀴", "휠", "기어", "샤프트", "축"],
                "전원부품": ["배터리", "전원", "전압", "DC"],
                "기타": []
            }

            # 분류 함수 정의
            def subcategorize_mcu(name):
                if re.search("ESP32", name, re.IGNORECASE):
                    return "ESP32"
                elif re.search("STM32", name, re.IGNORECASE):
                    return "STM32"
                elif re.search("아두이노|Arduino", name, re.IGNORECASE):
                    return "아두이노"
                else:
                    return "기타"

            def categorize_product(name):
                for category, keywords in category_keywords.items():
                    for keyword in keywords:
                        if re.search(keyword, name, re.IGNORECASE):
                            return category
                return "기타"

            # 분류 적용
            result_df = pd.DataFrame({'상품명': product_names})
            result_df['카테고리'] = result_df['상품명'].apply(categorize_product)
            result_df['중분류'] = result_df.apply(
                lambda row: subcategorize_mcu(row['상품명']) if row['카테고리'] == 'MCU/개발보드' else '기타',
                axis=1
            )
            result_df['소분류'] = ''  # 소분류 열 추가
            result_df['비고'] = ''    # 비고 열 추가

            # 결과 정렬
            result_df = result_df[['카테고리', '중분류', '소분류', '상품명', '비고']].sort_values(by=['카테고리', '중분류', '소분류', '상품명'])

            # 테이블 위젯 초기화
            table = self.ui.categoryTableWidget
            table.setRowCount(0)
            table.setSortingEnabled(False)  # 정렬 임시 비활성화

            # 이미지 정보 저장
            self.image_info_list = []
            
            # 결과를 테이블에 표시
            for idx, row in result_df.iterrows():
                row_position = table.rowCount()
                table.insertRow(row_position)
                
                # 각 열에 아이템 추가
                for col, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    table.setItem(row_position, col, item)
                
                # 이미지 열에 버튼 추가
                if has_image_column:
                    product_name = row['상품명']
                    image_url = df[df['상품명'] == product_name]['대표이미지 URL'].iloc[0] if len(df[df['상품명'] == product_name]) > 0 else None
                    
                    if pd.notna(image_url) and str(image_url).strip():
                        print(f"이미지 URL 추가: {product_name} - {image_url}")
                        
                        # 이미지 정보 저장
                        self.image_info_list.append({
                            'product_name': product_name,
                            'image_url': image_url
                        })
                        
                        button = QPushButton("이미지 보기")
                        button.setStyleSheet("""
                            QPushButton {
                                background-color: #4CAF50;
                                color: white;
                                border: none;
                                padding: 5px;
                                border-radius: 3px;
                            }
                            QPushButton:hover {
                                background-color: #45a049;
                            }
                        """)
                        button.clicked.connect(lambda checked, url=image_url, name=product_name, idx=len(self.image_info_list)-1: 
                                             self.show_image(url, name, self.image_info_list, idx))
                        table.setCellWidget(row_position, 4, button)  # 이미지 열 위치 변경
                    else:
                        # 이미지 URL이 없는 경우 빈 셀 추가
                        empty_item = QTableWidgetItem("")
                        table.setItem(row_position, 4, empty_item)  # 이미지 열 위치 변경

            # 열 너비 자동 조정
            table.resizeColumnsToContents()
            table.setSortingEnabled(True)  # 정렬 다시 활성화

            # 분류 결과 저장
            self.categorized_df = result_df
            
            print(f"✓ 상품 분류 완료")
            print(f"총 {len(result_df)}개의 상품이 분류되었습니다.")
            
            # 카테고리별 상품 수 출력
            category_counts = result_df['카테고리'].value_counts()
            print("\n[카테고리별 상품 수]")
            for category, count in category_counts.items():
                print(f"{category}: {count}개")

            self.statusBar().showMessage(f"상품 분류 완료: 총 {len(result_df)}개 상품")

        except Exception as e:
            error_msg = str(e)
            print(f"❌ 상품 분류 중 오류 발생: {error_msg}")
            QMessageBox.critical(
                self,
                "오류",
                f"상품 분류 중 오류가 발생했습니다.\n\n{error_msg}"
            )

    def show_image(self, image_url, product_name, all_images=None, current_index=None):
        """이미지 URL을 받아 다이얼로그로 표시합니다."""
        dialog = ImageDialog(image_url, product_name, self, all_images, current_index, self.ui.categoryTableWidget)
        dialog.exec()

    def export_category_excel(self):
        """분류된 상품 목록을 엑셀 파일로 내보냅니다."""
        if not hasattr(self, 'categorized_df'):
            QMessageBox.warning(self, "경고", "먼저 상품을 분류해주세요.")
            return

        try:
            # 테이블에서 최신 데이터 가져오기
            table = self.ui.categoryTableWidget
            updated_df = self.categorized_df.copy()
            
            # 테이블에서 소분류와 비고 데이터 업데이트
            for row in range(table.rowCount()):
                # 상품명 열 확인
                product_name_item = table.item(row, 3)
                if product_name_item is None:
                    continue
                product_name = product_name_item.text()
                
                # 소분류 열 확인
                subcategory_item = table.item(row, 2)
                subcategory = subcategory_item.text() if subcategory_item is not None else ""
                
                # 비고 열 확인
                note_item = table.item(row, 5)
                note = note_item.text() if note_item is not None else ""
                
                # 데이터프레임에서 해당 상품 찾아 업데이트
                mask = updated_df['상품명'] == product_name
                if any(mask):
                    updated_df.loc[mask, '소분류'] = subcategory
                    updated_df.loc[mask, '비고'] = note
            
            # output 디렉토리 생성
            output_dir = Path("output")
            output_dir.mkdir(exist_ok=True)
            
            # 현재 시간을 파일명에 포함
            current_time = datetime.now().strftime("%Y%m%d%H%M%S")
            output_file = (output_dir / f"Product_Categorized_{current_time}.xlsx").resolve()
            
            print(f"\n[분류 결과 엑셀 파일 생성]")
            print(f"출력 파일: {output_file}")

            # 엑셀 파일로 저장
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                updated_df.to_excel(writer, index=False, sheet_name='Sheet1')
                
                # 워크시트와 워크북 객체 가져오기
                worksheet = writer.sheets['Sheet1']
                workbook = writer.book
                
                # 가운데 정렬을 위한 셀 포맷 설정
                center_format = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                # 헤더 포맷 설정 (가운데 정렬 + 굵게)
                header_format = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter',
                    'bold': True
                })
                
                # 열 너비 자동 조정 및 포맷 적용
                for idx, col in enumerate(updated_df.columns):
                    # 열 이름의 길이와 데이터의 최대 길이 계산
                    max_length = max(
                        updated_df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    # 한글은 2배의 너비가 필요하므로 조정
                    adjusted_width = max_length * 2 if any('\u3131' <= c <= '\u318E' or '\uAC00' <= c <= '\uD7A3' for c in str(col)) else max_length
                    worksheet.set_column(idx, idx, adjusted_width + 2, center_format)
                
                # 헤더에 포맷 적용
                for col_num, value in enumerate(updated_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # 전체 행 높이 조정
                worksheet.set_default_row(20)

            print(f"✓ 분류 결과 파일이 생성되었습니다.")
            print(f"  - 파일 위치: {output_file}")
            print(f"  - 행 수: {len(updated_df)}")

            self.statusBar().showMessage(f"분류 결과 파일 생성 완료: {output_file}")
            
            # 성공 메시지 표시 (커스텀 버튼 포함)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("완료")
            msg.setText("분류 결과 파일이 생성되었습니다.")
            msg.setInformativeText(f"파일 위치:\n{output_file}")
            
            # 버튼 추가
            open_location_button = msg.addButton("폴더 열기", QMessageBox.ActionRole)
            open_file_button = msg.addButton("엑셀 열기", QMessageBox.ActionRole)
            close_button = msg.addButton("닫기", QMessageBox.RejectRole)
            
            msg.setDefaultButton(close_button)
            
            # 메시지 박스 표시
            clicked_button = msg.exec()
            
            # 버튼 클릭 처리
            if msg.clickedButton() == open_location_button:
                if not self.open_file_location(output_file):
                    QMessageBox.warning(self, "오류", "파일 위치를 열 수 없습니다.")
            elif msg.clickedButton() == open_file_button:
                if not self.open_file_with_default_app(output_file):
                    QMessageBox.warning(self, "오류", "파일을 열 수 없습니다.\n엑셀이 설치되어 있는지 확인해주세요.")

        except Exception as e:
            error_msg = str(e)
            print(f"❌ 분류 결과 파일 생성 중 오류 발생: {error_msg}")
            QMessageBox.critical(
                self,
                "오류",
                f"분류 결과 파일 생성 중 오류가 발생했습니다.\n\n{error_msg}"
            )

    def open_output_folder(self):
        """output 폴더를 생성하고 엽니다."""
        try:
            # output 폴더 경로 설정
            output_dir = Path("output")
            
            # 폴더가 없으면 생성
            output_dir.mkdir(exist_ok=True)
            
            # 폴더 열기
            if sys.platform == 'win32':
                os.startfile(str(output_dir))
            elif sys.platform == 'darwin':  # macOS
                subprocess.call(('open', str(output_dir)))
            else:  # linux
                subprocess.call(('xdg-open', str(output_dir)))
                
            self.statusBar().showMessage("output 폴더가 열렸습니다.", 2000)
            
        except Exception as e:
            QMessageBox.warning(self, "오류", f"폴더를 열 수 없습니다: {str(e)}")

def main():
    print("프로그램 시작")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    print("메인 윈도우 표시")
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 
