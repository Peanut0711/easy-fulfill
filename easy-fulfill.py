#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import re
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
                              QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QWidget)
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QIODevice, Qt, QSize, QUrl
from PySide6.QtGui import QPixmap, QImage, QIcon, QAction, QDesktopServices
import requests
from io import BytesIO
import warnings
import logging
import json


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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        print("초기화 시작")
        self.selected_file_path = None
        self.store_type = None
        self.orders = {}  # orders 변수를 인스턴스 변수로 초기화
        self.is_order_file_valid = False  # 주문서 파일 유효성 플래그
        self.is_invoice_file_valid = False  # 송장 파일 유효성 플래그
        
        # 인덱스 파일 경로 설정
        self.index_file_path = Path("database") / "order_index.json"
        
        # 인덱스 값 초기화
        self.current_idx_naver = 1
        self.current_idx_coupang = 1
        self.current_idx_gmarket = 1
        
        self.load_ui()
        self.setup_connections()
        self.setup_status_bar()
        
        # 인덱스 값 로드
        self.load_index_values()
        
        print("초기화 완료")

    def load_index_values(self):
        """저장된 인덱스 값을 로드합니다."""
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

    def save_index_values(self):
        """현재 인덱스 값을 저장합니다."""
        try:
            # database 디렉토리 생성
            self.index_file_path.parent.mkdir(exist_ok=True)
            
            # 기존 데이터 로드
            index_data = {}
            if self.index_file_path.exists() and self.index_file_path.stat().st_size > 0:
                try:
                    with open(self.index_file_path, 'r', encoding='utf-8') as f:
                        index_data = json.load(f)
                except json.JSONDecodeError:
                    print("! 인덱스 파일이 손상되었습니다. 새로 생성합니다.")
                    index_data = {}
            
            # 오늘 날짜의 데이터 업데이트
            today = date.today().strftime('%Y-%m-%d')
            if today not in index_data:
                index_data[today] = {}
            
            # 인덱스 값 저장
            index_data[today]['naver'] = self.current_idx_naver
            index_data[today]['gmarket'] = self.current_idx_gmarket
            index_data[today]['coupang'] = self.current_idx_coupang
            
            # 파일에 저장
            with open(self.index_file_path, 'w', encoding='utf-8') as f:
                json.dump(index_data, f, ensure_ascii=False, indent=2)
            
            print(f"✓ 인덱스 값 저장 완료 (날짜: {today})")
            print(f"  - 네이버: {self.current_idx_naver}")
            print(f"  - 지마켓: {self.current_idx_gmarket}")
            print(f"  - 쿠팡: {self.current_idx_coupang}")
        except Exception as e:
            print(f"! 인덱스 값 저장 중 오류 발생: {str(e)}")
            # 오류 발생 시 파일이 손상되지 않도록 기존 파일 백업
            if self.index_file_path.exists():
                backup_path = self.index_file_path.with_suffix('.json.bak')
                try:
                    shutil.copy2(self.index_file_path, backup_path)
                    print(f"! 기존 인덱스 파일을 백업했습니다: {backup_path}")
                except Exception as backup_error:
                    print(f"! 백업 파일 생성 실패: {str(backup_error)}")

    def update_naver_index(self):
        """네이버 인덱스 값을 업데이트하고 저장합니다."""
        self.current_idx_naver += 1
        if hasattr(self.ui, 'lineEdit_idx_naver'):
            self.ui.lineEdit_idx_naver.setText(str(self.current_idx_naver))
        self.save_index_values()

    def update_coupang_index(self):
        """쿠팡 인덱스 값을 업데이트하고 저장합니다."""
        self.current_idx_coupang += 1
        if hasattr(self.ui, 'lineEdit_idx_coupang'):
            self.ui.lineEdit_idx_coupang.setText(str(self.current_idx_coupang))
        self.save_index_values()

    def update_gmarket_index(self):
        """지마켓 인덱스 값을 업데이트하고 저장합니다."""
        self.current_idx_gmarket += 1
        if hasattr(self.ui, 'lineEdit_idx_gmarket'):
            self.ui.lineEdit_idx_gmarket.setText(str(self.current_idx_gmarket))
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
        
        # 노션 홈페이지 액션
        notionAction = QAction(QIcon('image/notion-icon.png'), '노션 홈페이지로 이동', self)        
        notionAction.setStatusTip('노션 홈페이지로 이동')
        notionAction.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://www.notion.so")))
        toolbar.addAction(notionAction)
        
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
        # 발송 처리 탭 버튼 연결
        self.ui.pushButton_load_invoice.clicked.connect(self.load_invoice_file)
        self.ui.pushButton_generate_invoice.clicked.connect(self.generate_invoice_file)
        
        # 인덱스 입력 필드 연결
        if hasattr(self.ui, 'lineEdit_idx_naver'):
            self.ui.lineEdit_idx_naver.textChanged.connect(self.on_naver_index_changed)
        if hasattr(self.ui, 'lineEdit_idx_coupang'):
            self.ui.lineEdit_idx_coupang.textChanged.connect(self.on_coupang_index_changed)
        if hasattr(self.ui, 'lineEdit_idx_gmarket'):
            self.ui.lineEdit_idx_gmarket.textChanged.connect(self.on_gmarket_index_changed)
        
        # 메뉴 동작 연결
        self.ui.actionOpenExcel.triggered.connect(self.select_excel_file)
        self.ui.actionExit.triggered.connect(self.close)
        self.ui.actionAbout.triggered.connect(self.show_about)
        print("버튼과 메뉴 연결 완료")

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
                    scaled_pixmap = pixmap.scaled(logo_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    self.ui.label_logo.setPixmap(scaled_pixmap)
                    self.ui.label_logo.setAlignment(Qt.AlignCenter)
                else:
                    print(f"! 로고 파일을 찾을 수 없습니다: {logo_path}")
                
                # 스토어 타입에 따라 다른 처리 메서드 호출
                try:
                    if self.store_type == "naver":
                        print("✓ 네이버 스토어 파일 처리 시작")
                        self.process_naver_excel_file()
                        self.is_order_file_valid = True  # 파일 처리 성공 시 플래그 설정
                    elif self.store_type == "coupang":
                        print("✓ 쿠팡 스토어 파일 처리 시작")
                        self.process_coupang_excel_file()
                        self.is_order_file_valid = True  # 파일 처리 성공 시 플래그 설정
                    elif self.store_type == "gmarket":
                        print("✓ 지마켓 스토어 파일 처리 시작")
                        self.process_gmarket_excel_file()
                        self.is_order_file_valid = True  # 파일 처리 성공 시 플래그 설정
                except Exception as e:
                    self.is_order_file_valid = False
                    print(f"❌ 파일 처리 중 오류 발생: {str(e)}")
                    QMessageBox.warning(self, "오류", f"파일 처리 중 오류가 발생했습니다: {str(e)}")
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
            
            # store_database.xlsx 파일 읽기
            store_db_path = Path("database") / "store_database.xlsx"
            if not store_db_path.exists():
                raise FileNotFoundError("store_database.xlsx 파일을 찾을 수 없습니다.")
            
            print("\n[store_database.xlsx 파일 읽기]")
            print(f"파일 경로: {store_db_path}")
            store_df = pd.read_excel(store_db_path, sheet_name=0)  # 첫 번째 시트

            # 데이터프레임 정보 출력
            print("\n[데이터프레임 정보]")
            print(f"행 수: {len(store_df)}")
            print(f"열 수: {len(store_df.columns)}")
            print("열 이름:")
            for col in store_df.columns:
                print(f"- {col}")

            # 상품번호와 상품코드 매핑 생성
            product_mapping = {}
            print("\n[상품번호-상품코드 매핑 생성]")
            print("\n[데이터 샘플]")
            print(store_df[['상품번호(스마트스토어)', '상품코드']].head())

            for idx, row in store_df.iterrows():
                # 상품번호에서 .0 제거
                product_number = str(row['상품번호(스마트스토어)']).strip()
                if product_number.endswith('.0'):
                    product_number = product_number[:-2]
                
                # 상품코드 처리
                product_code = str(row['상품코드']).strip()
                if product_code == 'nan':
                    product_code = '        '
                
                print(f"\n처리 중인 행 {idx}:")
                print(f"원본 상품번호: {row['상품번호(스마트스토어)']}")
                print(f"원본 상품코드: {row['상품코드']}")
                print(f"변환된 상품번호: {product_number}")
                print(f"변환된 상품코드: {product_code}")
                
                if product_number and product_number != 'nan':
                    if product_number in product_mapping:
                        print(f"경고: 중복된 상품번호 발견! {product_number}")
                        print(f"기존 매핑: {product_mapping[product_number]}")
                        print(f"새로운 매핑: {product_code}")
                    product_mapping[product_number] = product_code
                    print(f"매핑 추가: {product_number} -> {product_mapping[product_number]}")

            
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
                '배송방법(구매자 요청)': None  # 배송 방법 열 추가
            }
            
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
                    '상품목록': []
                }
                
                # 패턴에 해당하는 모든 주문의 상품 정보 추가
                for _, row in df.iterrows():
                    order_number = str(row[required_columns['주문번호']])
                    if order_number in order_numbers:
                        # 상품 정보
                        product_name = str(row[required_columns['상품명']])
                        quantity = int(row[required_columns['수량']]) if not pd.isna(row[required_columns['수량']]) else 1
                        option = str(row[required_columns['옵션정보']]) if not pd.isna(row[required_columns['옵션정보']]) else "없음"
                        
                        # 상품번호 가져오기
                        product_number = str(row[required_columns['상품번호']]).strip()
                        print(f"\n[상품번호 매칭]")
                        print(f"상품명: {product_name}")
                        print(f"원본 상품번호: {row[required_columns['상품번호']]}")
                        print(f"변환된 상품번호: {product_number}")
                        # print(f"매핑 딕셔너리 키 목록: {list(product_mapping.keys())}")
                        product_code = product_mapping.get(product_number, '        ')
                        print(f"매칭된 상품코드: {product_code}")
                        
                        # 상품 정보 추가
                        self.orders[pattern]['상품수'] += 1
                        self.orders[pattern]['상품목록'].append({
                            '상품명': product_name,
                            '수량': quantity,
                            '옵션': option,
                            '상품코드': product_code
                        })
                        
                        # 수취인 정보가 다른 경우 경고
                        if (self.orders[pattern]['수취인명'] != str(row[required_columns['수취인명']]) or
                            self.orders[pattern]['수취인연락처1'] != str(row[required_columns['수취인연락처1']]) or
                            self.orders[pattern]['통합배송지'] != str(row[required_columns['통합배송지']])):
                            print(f"! 주문번호 패턴 {pattern}의 수취인 정보가 다릅니다:")
                            print(f"  - 기존: {self.orders[pattern]['수취인명']} / {self.orders[pattern]['수취인연락처1']} / {self.orders[pattern]['통합배송지']}")
                            print(f"  - 새로운: {row[required_columns['수취인명']]} / {row[required_columns['수취인연락처1']]} / {row[required_columns['통합배송지']]}")
            
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
                # 배송 방법이 '택배,등기,소포'인 경우에만 기존 형식으로 표시
                if info['배송방법'] == '택배,등기,소포':
                    markdown_text += f"[ ] {self.current_idx_naver}.{info['수취인명']}\n"
                else:
                    # 그 외의 경우 배송 방법을 함께 표시
                    markdown_text += f"[ ] {self.current_idx_naver}.{info['수취인명']} **({info['배송방법']})**\n"
                
                self.update_naver_index()
                if hasattr(self.ui, 'lineEdit_idx_naver'):
                    self.ui.lineEdit_idx_naver.setText(str(self.current_idx_naver))
                
                # 상품 목록 표시
                for product in info['상품목록']:
                    product_name = product['상품명']
                    quantity = product['수량']
                    option = product['옵션']
                    product_code = product['상품코드']
                    
                    markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {product_name} ( 옵션 : {option} )\n"
                
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
            QMessageBox.critical(
                self,
                "오류",
                "엑셀 파일 처리 중 오류가 발생했습니다.\n\n"
                f"{error_msg}\n\n"
                "다음 사항을 확인해주세요:\n"
                "1. 파일이 손상되지 않았는지\n"
                "2. 다른 프로그램에서 파일을 열고 있지 않은지\n"
                "3. 파일을 다시 저장하거나 다른 형식(.xlsx)으로 변환해보세요."
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
            
            # store_database.xlsx 파일 읽기
            try:
                db_path = Path("database") / "store_database.xlsx"
                if not db_path.exists():
                    raise FileNotFoundError("store_database.xlsx 파일을 찾을 수 없습니다.")
                
                # 두 번째 시트 읽기 (시트 이름이 날짜로 변경될 수 있으므로 인덱스로 접근)
                db_df = pd.read_excel(db_path, sheet_name=1, header=1)
                print("\n[상품 데이터베이스 로드 완료]")
                print("데이터베이스 열 목록:")
                for col in db_df.columns:
                    print(f"- {col}")
                
                # 상품코드와 옵션ID 매핑 생성
                product_code_map = {}
                for _, row in db_df.iterrows():
                    option_id_raw = row['옵션 ID']
                    product_code = str(row['상품코드']).strip()
                    
                    # NaN 값 처리
                    if pd.isna(option_id_raw):
                        continue
                    
                    # float로 읽힌 경우 정수로 변환 후 문자열로 변환
                    if isinstance(option_id_raw, float):
                        option_id = str(int(option_id_raw))
                    else:
                        option_id = str(option_id_raw).strip()
                    
                    if option_id and option_id != 'nan':
                        product_code_map[option_id] = product_code
                
                print(f"✓ {len(product_code_map)}개의 상품 매핑 정보를 로드했습니다.")
                
            except Exception as e:
                print(f"! 상품 데이터베이스 로드 중 오류 발생: {str(e)}")
                QMessageBox.warning(self, "경고", "상품 데이터베이스 로드 중 오류가 발생했습니다.")
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
                '옵션ID': None  # 공백 제거
            }
            
            for col in df.columns:
                col_str = str(col).strip()
                for key in required_columns.keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        required_columns[key] = col
                        print(f"✓ '{key}' 열을 찾았습니다: {col}")
            
            # 필수 열이 모두 있는지 확인
            missing_columns = [key for key, value in required_columns.items() if value is None]
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
                        '상품목록': []
                    }
                
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
                    # float로 읽힌 경우 정수로 변환 후 문자열로 변환
                    if isinstance(option_id_raw, float):
                        option_id = str(int(option_id_raw))
                    else:
                        option_id = str(option_id_raw).strip()
                
                product_code = product_code_map.get(option_id, '')  # 매칭되는 상품코드가 없으면 빈 문자열
                
                self.orders[order_number]['상품목록'].append({
                    '상품명': product_name,
                    '옵션': option,
                    '수량': quantity,
                    '상품코드': product_code
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
                # print(f"[주문 처리 시작] 주문번호: {order_number}")                
                # key = (info['수취인이름'], info['수취인전화번호'], order_number)
                markdown_text += f"[ ] {self.current_idx_coupang}.{info['수취인이름']}\n"
                self.update_coupang_index()
                if hasattr(self.ui, 'lineEdit_idx_coupang'):
                    self.ui.lineEdit_idx_coupang.setText(str(self.current_idx_coupang))
                
                # 상품 목록 표시
                for product in info['상품목록']:
                    product_name = product['상품명']
                    quantity = product['수량']
                    option = product['옵션']
                    product_code = product['상품코드']                    
                    
                    markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {product_name} ( 옵션 : {option} )\n"
                
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
            QMessageBox.critical(
                self,
                "오류",
                f"엑셀 파일 처리 중 오류가 발생했습니다.\n\n{error_msg}"
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
                '우편번호': None
            }
            
            for col in df.columns:
                col_str = str(col).strip()
                for key in required_columns.keys():
                    if col_str == key:  # 정확히 일치하는 경우에만 매칭
                        required_columns[key] = col
                        print(f"✓ '{key}' 열을 찾았습니다: {col}")
            
            # 필수 열이 모두 있는지 확인
            missing_columns = [key for key, value in required_columns.items() if value is None]
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
                    self.orders[order_number] = {
                        '수령인명': str(row[required_columns['수령인명']]),
                        '주소': str(row[required_columns['주소']]),
                        '수령인 전화번호': str(row[required_columns['수령인 전화번호']]) if not pd.isna(row[required_columns['수령인 전화번호']]) else '',
                        '수령인 휴대폰': str(row[required_columns['수령인 휴대폰']]) if not pd.isna(row[required_columns['수령인 휴대폰']]) else '',
                        '배송시 요구사항': str(row[required_columns['배송시 요구사항']]) if not pd.isna(row[required_columns['배송시 요구사항']]) else '',
                        '우편번호': str(row[required_columns['우편번호']]) if not pd.isna(row[required_columns['우편번호']]) else '',
                        '상품목록': []
                    }
                
                # 상품 정보 추가
                product_name = str(row[required_columns['상품명']])
                option = str(row[required_columns['옵션']])
                quantity = int(row[required_columns['수량']]) if not pd.isna(row[required_columns['수량']]) else 1
                
                # 옵션ID로 상품코드 찾기
                # option_id = str(row[required_columns['옵션ID']]).strip()
                # product_code = product_code_map.get(option_id, '')  # 매칭되는 상품코드가 없으면 빈 문자열
                
                self.orders[order_number]['상품목록'].append({
                    '상품명': product_name,
                    '옵션': option,
                    '수량': quantity,
                    # '상품코드': product_code
                })
            
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
            
            for order_number, info in self.orders.items():
                # print(f"[주문 처리 시작] 주문번호: {order_number}")                
                # key = (info['수취인이름'], info['수취인전화번호'], order_number)
                markdown_text += f"[ ] {self.current_idx_gmarket}.{info['수령인명']}\n"
                self.update_gmarket_index()
                if hasattr(self.ui, 'lineEdit_idx_gmarket'):
                    self.ui.lineEdit_idx_gmarket.setText(str(self.current_idx_gmarket))
                
                # 상품 목록 표시
                for product in info['상품목록']:
                    product_name = product['상품명']
                    quantity = product['수량']
                    option = product['옵션']
                    # product_code = product['상품코드'] 
                    product_code = 'PASS'
                   
                                   
                    
                    # markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {product_name} ( 옵션 : {option} )\n"
                    markdown_text += f"▶ [{product_code}]**[ {quantity} 개 ]** - {product_name} ( 옵션 : {option} )\n"
                
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
            QMessageBox.critical(
                self,
                "오류",
                f"엑셀 파일 처리 중 오류가 발생했습니다.\n\n{error_msg}"
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
        
        # 발송처리 탭 텍스트 초기화
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
                    '통합배송지': None
                },
                'invoice': {
                    '등기번호': None,
                    '수취인명': None,
                    '수취인 이동통신': None,
                    '수취인상세주소': None
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
            
            # 6. 매칭된 주문 정보 출력 및 운송장번호 업데이트
            print("\n[매칭된 주문 정보]")
            matched_count = 0
            
            for idx, invoice_row in invoice_df.iterrows():
                invoice_name = str(invoice_row[column_mapping['invoice']['수취인명']])
                invoice_phone = str(invoice_row[column_mapping['invoice']['수취인 이동통신']])
                invoice_number = str(invoice_row[column_mapping['invoice']['등기번호']])
                
                matching_rows = order_df[
                    (order_df[column_mapping['order']['수취인명']] == invoice_name) &
                    (order_df[column_mapping['order']['수취인연락처1']] == invoice_phone)
                ]
                
                if not matching_rows.empty:
                    matched_count += 1
                    print(f"\n[매칭된 주문 {matched_count}]")
                    print(f"수취인명: {invoice_name}")
                    print(f"전화번호: {invoice_phone}")
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
            
            # 운송장번호 컬럼 추가
            order_df['운송장번호'] = ''
            
            # 매칭 카운터
            matched_count = 0
            
            # 각 송장 행에 대해 매칭 시도
            for idx, invoice_row in invoice_df.iterrows():
                invoice_number = str(invoice_row[required_invoice_columns['등기번호']])
                invoice_name = str(invoice_row[required_invoice_columns['수취인명']]).strip()
                invoice_phone = str(invoice_row[required_invoice_columns['수취인 전화번호']]).strip()
                
                # 수취인명과 전화번호로 매칭
                matching_rows = order_df[
                    (order_df[required_order_columns['수취인이름']].str.strip() == invoice_name) &
                    (order_df[required_order_columns['수취인전화번호']].str.strip() == invoice_phone)
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
            
            print(f"\n✓ 총 {matched_count}개의 주문이 매칭되었습니다.")
            
            # 4. 결과 파일 저장
            self._save_invoice_file(order_df)
            
        except Exception as e:
            error_msg = str(e)
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
                max_length = max(
                    order_df[col].astype(str).apply(len).max(),
                    len(str(col))
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
        open_file_button = msg.addButton("파일 열기", QMessageBox.ActionRole)
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
            open_file_button = msg.addButton("파일 열기", QMessageBox.ActionRole)
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