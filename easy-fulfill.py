#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import re
from datetime import datetime
from pathlib import Path
import pandas as pd
import numpy as np
import subprocess
import tempfile
import msoffcrypto
from PySide6.QtWidgets import (QApplication, QMainWindow, QFileDialog, QMessageBox, 
                              QInputDialog, QLineEdit, QTableWidgetItem, QLabel, 
                              QDialog, QVBoxLayout, QHBoxLayout, QPushButton)
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QIODevice, Qt, QSize
from PySide6.QtGui import QPixmap, QImage
import requests
from io import BytesIO
import warnings

# 이미지 경고 메시지 숨기기
warnings.filterwarnings("ignore", category=UserWarning, module="PIL.PngImagePlugin")

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
        self.orders = {}  # orders 변수를 인스턴스 변수로 초기화
        self.load_ui()
        self.setup_connections()
        self.setup_status_bar()
        print("초기화 완료")
        
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
        self.setMenuBar(window.menubar)
        self.setStatusBar(window.statusbar)
        self.setWindowTitle(window.windowTitle())
        self.resize(window.size())
        
    def setup_status_bar(self):
        """상태바를 초기화하고 기본 메시지를 설정합니다."""
        self.statusBar().showMessage("준비")
        
    def setup_connections(self):
        """버튼과 메뉴 동작을 연결합니다."""
        # 버튼 연결
        self.ui.selectFileButton.clicked.connect(self.select_excel_file)
        self.ui.generateButton.clicked.connect(self.generate_work_order)
        self.ui.exportInvoiceButton.clicked.connect(self.export_invoice_excel)
        
        # 상품 분류 탭 버튼 연결
        self.ui.selectProductFileButton.clicked.connect(self.select_product_file)
        self.ui.categorizeButton.clicked.connect(self.categorize_products)
        self.ui.exportCategoryButton.clicked.connect(self.export_category_excel)
        
        # 메뉴 동작 연결
        self.ui.actionOpenExcel.triggered.connect(self.select_excel_file)
        self.ui.actionExit.triggered.connect(self.close)
        self.ui.actionAbout.triggered.connect(self.show_about)
        print("버튼과 메뉴 연결 완료")
        
    def is_valid_filename(self, filename):
        """파일명이 올바른 형식인지 검사합니다."""
        print(f"\n[파일명 검증 시작] 파일명: {filename}")
        
        # 네이버 스토어 파일 형식 검사
        naver_pattern = r'^스마트스토어_(전체|선택)주문발주발송관리_(\d{8})_(\d{4})\.xlsx$'
        naver_match = re.match(naver_pattern, filename)
        
        # 쿠팡 스토어 파일 형식 검사
        coupang_pattern = r'^DeliveryList\((\d{4}-\d{2}-\d{2})\)_\(0\)\.xlsx$'
        coupang_match = re.match(coupang_pattern, filename)
        
        if naver_match:
            # 날짜와 시간 유효성 검사
            date_str = naver_match.group(2)
            time_str = naver_match.group(3)
            
            try:
                # 날짜 형식 검증 (YYYYMMDD)
                date = datetime.strptime(date_str, '%Y%m%d')
                print(f"✓ 날짜 형식 검증 완료: {date.strftime('%Y년 %m월 %d일')}")
                
                # 시간 형식 검증 (HHMM)
                hour = int(time_str[:2])
                minute = int(time_str[2:])
                
                if hour < 0 or hour > 23 or minute < 0 or minute > 59:
                    print(f"❌ 잘못된 시간 형식: {hour:02d}:{minute:02d}")
                    return False, None
                    
                print(f"✓ 시간 형식 검증 완료: {hour:02d}시 {minute:02d}분")
                print("✓ 네이버 스토어 파일명 검증 성공")
                return True, "naver"
                
            except ValueError:
                print("❌ 날짜/시간 형식이 올바르지 않습니다.")
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
        else:
            print("❌ 파일명 형식이 올바르지 않습니다.")
            print("   - 네이버 스토어 형식: 스마트스토어_전체주문발주발송관리_YYYYMMDD_HHMM.xlsx")
            print("   - 쿠팡 스토어 형식: DeliveryList(YYYY-MM-DD)_(0).xlsx")
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
        """엑셀 파일을 선택하는 다이얼로그를 표시합니다."""
        file_path, _ = QFileDialog.getOpenFileName(
        self,
        "엑셀 파일 선택",
        "input",
        "Excel Files (*.xlsx *.xls)"
        )                
        if file_path:
            print(f"\n[파일 선택됨] 경로: {file_path}")
            filename = os.path.basename(file_path)
            
            is_valid, store_type = self.is_valid_filename(filename)
            
            if is_valid:
                self.selected_file_path = file_path
                self.ui.filePathLabel.setText(filename)
                self.statusBar().showMessage(f"파일 선택됨: {filename}")
                print("✓ 파일이 성공적으로 선택되었습니다.")
                
                # 스토어 타입에 따라 다른 처리 메서드 호출
                if store_type == "naver":
                    print("✓ 네이버 스토어 파일 처리 시작")
                    self.process_naver_excel_file()
                elif store_type == "coupang":
                    print("✓ 쿠팡 스토어 파일 처리 시작")
                    self.process_coupang_excel_file()
            else:
                QMessageBox.warning(
                    self,
                    "잘못된 파일명",
                    "올바른 파일명 형식이 아닙니다.\n\n"
                    "네이버 스토어: 스마트스토어_전체주문발주발송관리_20240405_1509.xlsx\n"
                    "쿠팡 스토어: DeliveryList(2025-04-06)_(0).xlsx"
                )
                self.selected_file_path = None
                self.ui.filePathLabel.setText("선택된 파일 없음")
                self.statusBar().showMessage("잘못된 파일명")
                print("❌ 파일 선택이 취소되었습니다.")
        else:
            print("\n[알림] 파일 선택이 취소되었습니다.")
            
    def process_naver_excel_file(self):
        """네이버 스토어 엑셀 파일에서 주문번호를 처리합니다."""
        try:
            print(f"\n[네이버 스토어 엑셀 파일 처리 시작] 파일: {self.selected_file_path}")
            
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
                print(f"- 열 이름: {list(df.columns)}")
                
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
                        print(f"- 열 이름: {list(df.columns)}")
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
                '수량': None
            }
            
            print("\n[열 정보]")
            print(f"감지된 열 목록: {', '.join(str(col) for col in df.columns)}")
            
            for col in df.columns:
                col_str = str(col)
                for key in required_columns.keys():
                    if key in col_str:
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
                    '배송메세지': str(first_order[required_columns['배송메세지']]).replace('nan', ''),
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
                        
                        # 상품 정보 추가
                        self.orders[pattern]['상품수'] += 1  # orders를 self.orders로 변경
                        self.orders[pattern]['상품목록'].append({  # orders를 self.orders로 변경
                            '상품명': product_name,
                            '수량': quantity
                        })
                        
                        # 수취인 정보가 다른 경우 경고
                        if (self.orders[pattern]['수취인명'] != str(row[required_columns['수취인명']]) or  # orders를 self.orders로 변경
                            self.orders[pattern]['수취인연락처1'] != str(row[required_columns['수취인연락처1']]) or
                            self.orders[pattern]['통합배송지'] != str(row[required_columns['통합배송지']])):
                            print(f"! 주문번호 패턴 {pattern}의 수취인 정보가 다릅니다:")
                            print(f"  - 기존: {self.orders[pattern]['수취인명']} / {self.orders[pattern]['수취인연락처1']} / {self.orders[pattern]['통합배송지']}")  # orders를 self.orders로 변경
                            print(f"  - 새로운: {row[required_columns['수취인명']]} / {row[required_columns['수취인연락처1']]} / {row[required_columns['통합배송지']]}")
            
            # 주문 정보 출력
            print("\n[주문 정보 출력]")
            print(f"총 {len(self.orders)}개의 주문이 있습니다.")
            
            for pattern, info in self.orders.items():
                print(f"\n주문번호 패턴: {pattern}")
                print(f"주문번호 목록: {', '.join(info['주문번호목록'])}")
                print(f"수취인명: {info['수취인명']}")
                print(f"수취인연락처: {info['수취인연락처1']}")
                print(f"통합배송지: {info['통합배송지']}")
                print(f"구매자연락처: {info['구매자연락처']}")
                print(f"배송메세지: {info['배송메세지']}")
                print(f"상품수: {info['상품수']}")
                print("\n[상품 목록]")
                for product in info['상품목록']:
                    print(f"- {product['상품명']} x {product['수량']}")
                print("-" * 50)
                
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
        # 쿠팡 스토어 파일 처리 로직은 추후 구현 예정
        QMessageBox.information(
            self,
            "알림",
            "쿠팡 스토어 파일 처리 기능은 아직 구현되지 않았습니다.\n"
            "추후 업데이트를 기대해주세요."
        )
        print("❌ 쿠팡 스토어 파일 처리 기능은 아직 구현되지 않았습니다.")

    def generate_work_order(self):
        """작업지시서 생성 버튼 클릭 시 실행되는 함수입니다."""
        if not self.selected_file_path:
            QMessageBox.warning(self, "경고", "먼저 엑셀 파일을 선택해주세요.")
            return
            
        print("작업지시서 생성됨")
        self.statusBar().showMessage("작업지시서 생성 완료")
        # 여기에 실제 작업지시서 생성 로직이 추가될 예정입니다.
        
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
            
            # 수취인 정보 기준으로 주문 통합 (시간 무관)
            consolidated_orders = {}
            
            # 주문 통합 처리
            for pattern, info in self.orders.items():
                # 수취인명과 연락처로 키 생성 (주소는 제외 - 같은 사람이 다른 주소로 배송 요청할 수 있음)
                key = (info['수취인명'], info['수취인연락처1'])
                
                if key not in consolidated_orders:
                    consolidated_orders[key] = {
                        '수취인명': info['수취인명'],
                        '수취인연락처1': info['수취인연락처1'],
                        '통합배송지': info['통합배송지'],  # 첫 번째 주문의 주소 사용
                        '구매자연락처': info['구매자연락처'],
                        '배송메세지': info['배송메세지'],
                        '주문번호목록': info['주문번호목록'],
                        '상품목록': info['상품목록'].copy(),
                        '주문시간': min(info['주문번호목록'])  # 가장 빠른 주문 시간 기준
                    }
                else:
                    # 기존 주문에 상품 정보 추가
                    consolidated_orders[key]['주문번호목록'].extend(info['주문번호목록'])
                    consolidated_orders[key]['상품목록'].extend(info['상품목록'])
                    
                    # 주문시간이 더 빠른 경우 주소와 배송메시지 업데이트
                    earliest_order = min(info['주문번호목록'])
                    if earliest_order < consolidated_orders[key]['주문시간']:
                        consolidated_orders[key]['통합배송지'] = info['통합배송지']
                        consolidated_orders[key]['배송메세지'] = info['배송메세지']
                        consolidated_orders[key]['주문시간'] = earliest_order
                    # 기존 주문이 더 빠르고 현재 주문에 배송메시지가 있는 경우에만 업데이트
                    elif info['배송메세지'].strip():
                        consolidated_orders[key]['배송메세지'] = info['배송메세지']

            # 통합된 주문 정보로 송장 데이터 생성
            for info in consolidated_orders.values():
                invoice_data.append({
                    '받는분성명': info['수취인명'],
                    '받는분전화번호': info['수취인연락처1'],
                    '받는분기타연락처': '',
                    '받는분주소(전체, 분할)': info['통합배송지'],
                    '상품명': '전자제품',
                    '내품수량': '1',
                    '배송메세지1': info['배송메세지']
                })
            
            # 데이터프레임 생성
            df_invoice = pd.DataFrame(invoice_data)
            
            # 열 순서 지정
            columns = [
                '받는분성명',
                '받는분전화번호',
                '받는분기타연락처',
                '받는분주소(전체, 분할)',
                '상품명',
                '내품수량',
                '배송메세지1'
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
            print(f"❌ 송장 엑셀 파일 생성 중 오류 발생: {error_msg}")
            QMessageBox.critical(
                self,
                "오류",
                f"송장 엑셀 파일 생성 중 오류가 발생했습니다.\n\n{error_msg}"
            )

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

def main():
    print("프로그램 시작")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    print("메인 윈도우 표시")
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 