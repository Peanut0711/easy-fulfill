"""
구글 스프레드시트와 실시간 DB 동기화 스크립트

인증: google-oauth/credentials.json + token.json (OAuth, easy-fulfill과 동일 경로·google_sheets_oauth.py)

주요 기능:
1. CSV/XLSX 파일에서 실시간 DB 읽기
2. 구글 스프레드시트에서 기존 DB 읽기
3. 옵션ID 기준으로 신제품 찾기
4. 신제품 데이터를 스프레드시트에 APPEND
"""

import gspread
import pandas as pd
from pathlib import Path
import sys

from google_sheets_oauth import get_authorized_gspread_client
import io
import codecs

# 터미널 인코딩 설정 (Windows 한글 깨짐 해결)
if sys.platform == 'win32':
    try:
        # Windows에서 UTF-8 출력 강제
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
    except:
        pass

# ===== 설정 =====
# 스프레드시트 인증: 프로젝트 루트 google-oauth/ (credentials.json, token.json) — google_sheets_oauth.py
SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"

# 네이버/쿠팡 설정
NAVER_CONFIG = {
    "sheet_index": 0,  # 1번 시트
    "header_row": 1,   # 1행이 헤더
    "option_id_column": "상품번호(스마트스토어)",  # 네이버는 옵션ID가 없고 상품번호를 사용
    "db_dir": "database",
    "file_pattern": "Product_*.csv"
}

COUPANG_CONFIG = {
    "sheet_index": 1,  # 2번 시트
    "header_row": 2,   # 2행이 헤더
    "option_id_column": "옵션 ID",
    "db_dir": "database",
    "file_pattern": "price_inventory_*.xlsx"
}


# ===== 함수 정의 =====

def get_latest_file_from_pattern(directory, file_pattern):
    """
    디렉토리에서 패턴에 맞는 최신 파일 찾기
    
    Args:
        directory: 디렉토리 경로
        file_pattern: 파일 패턴 (예: "Product_*.csv", "price_inventory_*.xlsx")
    
    Returns:
        Path: 최신 파일 경로 (없으면 None)
    """
    from glob import glob
    import os
    
    # glob 패턴 사용
    pattern = os.path.join(directory, file_pattern)
    files = glob(pattern)
    
    if not files:
        return None
    
    # 수정 시간 기준으로 정렬하여 가장 최신 파일 반환
    latest_file = max(files, key=os.path.getmtime)
    return Path(latest_file)


def get_spreadsheet_headers(worksheet, header_row_num):
    """
    스프레드시트에서 헤더 행 가져오기
    
    Args:
        worksheet: gspread 워크시트 객체
        header_row_num: 헤더가 있는 행 번호 (1-based)
    
    Returns:
        list: 헤더 리스트
    """
    headers = worksheet.row_values(header_row_num)
    return headers


def get_all_spreadsheet_data(worksheet, header_row_num):
    """
    스프레드시트에서 모든 데이터 가져오기
    
    Args:
        worksheet: gspread 워크시트 객체
        header_row_num: 헤더가 있는 행 번호 (1-based)
    
    Returns:
        list of dict: 데이터 딕셔너리 리스트
    """
    if header_row_num == 1:
        # 일반적인 경우 (1행이 헤더)
        return worksheet.get_all_records()
    else:
        # 특수한 경우 (쿠팡처럼 2행이 헤더)
        # 중복 헤더 처리 필요
        all_values = worksheet.get_all_values()
        headers = all_values[header_row_num - 1]  # header_row_num은 1-based
        
        # 중복 헤더 처리
        unique_headers = []
        used_headers = set()
        column_indices_to_keep = []
        
        for idx, header in enumerate(headers):
            header_name = header if header else f"빈컬럼_{idx}"
            
            if not header or header.strip() == "":
                if f"빈컬럼_{idx}" not in used_headers:
                    unique_headers.append(f"빈컬럼_{idx}")
                    used_headers.add(f"빈컬럼_{idx}")
                    column_indices_to_keep.append(idx)
            elif header_name not in used_headers:
                unique_headers.append(header_name)
                used_headers.add(header_name)
                column_indices_to_keep.append(idx)
        
        # 데이터 변환
        data_records = []
        for row_idx in range(header_row_num, len(all_values)):
            row_data = all_values[row_idx]
            record = {}
            for i, header_name in enumerate(unique_headers):
                col_idx = column_indices_to_keep[i]
                if col_idx < len(row_data):
                    record[header_name] = row_data[col_idx]
                else:
                    record[header_name] = ""
            data_records.append(record)
        
        return data_records


def read_realtime_db(file_path, header_row=0):
    """
    CSV/XLSX 파일에서 실시간 DB 읽기
    
    Args:
        file_path: 파일 경로
        header_row: 헤더가 있는 행 번호 (0-based, 기본값: 0)
    
    Returns:
        list of dict: 데이터 딕셔너리 리스트
    """
    path = Path(file_path)
    
    if path.suffix.lower() == '.csv':
        df = pd.read_csv(file_path, header=header_row)
    elif path.suffix.lower() in ['.xlsx', '.xls']:
        df = pd.read_excel(file_path, header=header_row)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {path.suffix}")
    
    # 딕셔너리 리스트로 변환
    return df.to_dict('records')


def normalize_option_id(value):
    """
    옵션 ID를 비교용 문자열로 통일한다.
    Excel(pandas)은 큰 정수를 float으로 읽어 85284690572.0 이 되고,
    스프레드시트는 문자열 85284690572 이라 str()만 하면 불일치로 신제품 오판정된다.
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, str):
        s = value.strip()
        if not s or s.lower() == "nan":
            return ""
    try:
        f = float(value)
        if f.is_integer():
            return str(int(f))
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def format_option_id_for_log(value):
    """로그 출력용: 정수형이면 .0 없이, 없으면 N/A."""
    return normalize_option_id(value) or "N/A"


def format_quantity_for_log(value):
    """로그용: 가격·재고 등은 정수로 표시, 비어 있거나 숫자가 아니면 N/A 또는 원문."""
    if value is None or value == "" or value == "N/A":
        return "N/A"
    try:
        if pd.isna(value):
            return "N/A"
    except (TypeError, ValueError):
        pass
    if isinstance(value, str):
        s = value.strip()
        if not s or s.upper() == "N/A" or s.lower() == "nan":
            return "N/A"
        value = s
    try:
        f = float(value)
        if f != f:  # NaN
            return "N/A"
        return str(int(round(f)))
    except (TypeError, ValueError):
        t = str(value).strip()
        return t if t else "N/A"


def find_new_products(realtime_data, spreadsheet_data, option_id_column):
    """
    옵션ID 기준으로 신제품 찾기
    
    Args:
        realtime_data: 실시간 DB 데이터 (딕셔너리 리스트)
        spreadsheet_data: 스프레드시트 데이터 (딕셔너리 리스트)
        option_id_column: 옵션ID 컬럼명
    
    Returns:
        list of dict: 신제품 리스트
    """
    # 기존 옵션ID 집합 생성
    existing_option_ids = set()
    for row in spreadsheet_data:
        option_id = normalize_option_id(row.get(option_id_column, ''))
        if option_id:
            existing_option_ids.add(option_id)
    
    # 신제품 필터링
    new_products = []
    for row in realtime_data:
        option_id = normalize_option_id(row.get(option_id_column, ''))
        if option_id and option_id not in existing_option_ids:
            new_products.append(row)
    
    return new_products


def align_to_spreadsheet_headers(realtime_dict, spreadsheet_headers):
    """
    실시간 DB 데이터를 스프레드시트 헤더 순서로 정렬
    
    Args:
        realtime_dict: 실시간 DB 딕셔너리
        spreadsheet_headers: 스프레드시트 헤더 리스트
    
    Returns:
        list: 정렬된 데이터 리스트
    """
    import math
    
    aligned_row = []
    for header in spreadsheet_headers:
        # 상품코드 컬럼은 사용자가 직접 창고 위치를 입력하는 곳이므로
        # 신규 제품 추가 시 반드시 빈칸으로 설정
        if header == "상품코드":
            value = ""
        else:
            value = realtime_dict.get(header, "")
        
        # NaN 처리: Google Sheets에서 NaN을 JSON으로 변환할 수 없음
        if isinstance(value, float) and math.isnan(value):
            value = ""
        
        aligned_row.append(value)
    
    return aligned_row


def apply_banded_rows(spreadsheet, sheet_id, header_row_num, last_data_row, num_columns):
    """
    교차 색상 적용 (banded rows)
    
    Args:
        spreadsheet: gspread 스프레드시트 객체
        sheet_id: 시트 ID
        header_row_num: 헤더 행 번호 (1-based)
        last_data_row: 마지막 데이터 행 번호 (1-based)
        num_columns: 컬럼 개수
    """
    try:
        # 교차 색상 적용 범위: 헤더 다음 행부터 마지막 데이터 행까지
        start_row_index = header_row_num  # 헤더 다음 행 (0-based)
        end_row_index = last_data_row  # 마지막 데이터 행 다음 (0-based)
        
        # bandedRange 생성 요청
        requests = [{
            'addBanding': {
                'bandedRange': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': start_row_index,
                        'endRowIndex': end_row_index,
                        'startColumnIndex': 0,
                        'endColumnIndex': num_columns
                    },
                    'rowProperties': {
                        'headerColor': {
                            'red': 1.0,
                            'green': 1.0,
                            'blue': 1.0
                        },
                        'firstBandColor': {
                            'red': 1.0,
                            'green': 1.0,
                            'blue': 1.0
                        },
                        'secondBandColor': {
                            'red': 0.95,
                            'green': 0.95,
                            'blue': 0.95
                        }
                    }
                }
            }
        }]
        
        spreadsheet.batch_update({'requests': requests})
        print(f"✅ 교차 색상 적용 완료 (행 {start_row_index + 1}-{end_row_index})")
        
    except Exception as e:
        print(f"⚠️  교차 색상 적용 중 오류 발생: {str(e)}")


def append_to_spreadsheet(worksheet, new_data_list, header_row_num):
    """
    스프레드시트에 데이터 추가 (서식 포함)
    
    Args:
        worksheet: gspread 워크시트 객체
        new_data_list: 추가할 데이터 리스트 (각 항목은 리스트 형태)
        header_row_num: 헤더 행 번호 (1-based)
    """
    if not new_data_list:
        return
    
    # 데이터 추가 전 마지막 데이터 행 번호 저장
    all_values = worksheet.get_all_values()
    last_data_row = len(all_values)  # 마지막 행 번호 (1-based)
    
    # 데이터 추가
    worksheet.append_rows(new_data_list)
    
    # 추가된 행의 시작과 끝 행 번호 계산
    added_rows_start = last_data_row + 1
    added_rows_end = last_data_row + len(new_data_list)
    
    print(f"✅ {len(new_data_list)}개의 새 제품이 추가되었습니다 (행 {added_rows_start}-{added_rows_end})")
    
    # 서식 복사: 마지막 데이터 행의 서식을 새로 추가된 행에 복사
    try:
        # 마지막 데이터 행 찾기 (헤더 제외)
        source_row = header_row_num + 1  # 헤더 다음 행부터 시작
        
        if last_data_row > header_row_num:
            # 실제 데이터가 있는 경우, 마지막 데이터 행의 서식 사용
            source_row = last_data_row
        
        # Google Sheets API를 사용하여 서식 복사
        spreadsheet = worksheet.spreadsheet
        sheet_id = worksheet.id
        
        # 서식 복사 요청 (copyPaste 사용)
        requests = [{
            'copyPaste': {
                'source': {
                    'sheetId': sheet_id,
                    'startRowIndex': source_row - 1,  # 0-based
                    'endRowIndex': source_row,  # 0-based (다음 행 전까지)
                    'startColumnIndex': 0,
                    'endColumnIndex': len(new_data_list[0]) if new_data_list else 1
                },
                'destination': {
                    'sheetId': sheet_id,
                    'startRowIndex': added_rows_start - 1,  # 0-based
                    'endRowIndex': added_rows_end,  # 0-based
                    'startColumnIndex': 0,
                    'endColumnIndex': len(new_data_list[0]) if new_data_list else 1
                },
                'pasteType': 'PASTE_FORMAT',
                'pasteOrientation': 'NORMAL'
            }
        }]
        
        spreadsheet.batch_update({'requests': requests})
        print(f"✅ 서식 복사 완료 (행 {source_row} → 행 {added_rows_start}-{added_rows_end})")
        
    except Exception as e:
        print(f"⚠️  서식 복사 중 오류 발생 (데이터는 정상 추가됨): {str(e)}")


def append_to_spreadsheet_coupang(worksheet, new_data_list, header_row_num):
    """
    쿠팡 스프레드시트에 데이터 추가 (데이터 추가 → 교차 색상 → 서식 복사 순서)
    
    Args:
        worksheet: gspread 워크시트 객체
        new_data_list: 추가할 데이터 리스트 (각 항목은 리스트 형태)
        header_row_num: 헤더 행 번호 (1-based)
    """
    if not new_data_list:
        return
    
    # 데이터 추가 전 마지막 데이터 행 번호 저장
    all_values = worksheet.get_all_values()
    last_data_row = len(all_values)  # 마지막 행 번호 (1-based)
    
    # 헤더 행의 실제 컬럼 수 파악 (스프레드시트의 실제 컬럼 수)
    header_row_values = worksheet.row_values(header_row_num)
    num_columns = len(header_row_values) if header_row_values else (len(new_data_list[0]) if new_data_list else 1)
    
    # 1단계: 데이터 추가
    worksheet.append_rows(new_data_list)
    
    # 추가된 행의 시작과 끝 행 번호 계산
    added_rows_start = last_data_row + 1
    added_rows_end = last_data_row + len(new_data_list)
    
    print(f"✅ {len(new_data_list)}개의 새 제품이 추가되었습니다 (행 {added_rows_start}-{added_rows_end})")
    print(f"ℹ️  교차 색상 적용 범위: 헤더 행({header_row_num}) 다음부터 행 {added_rows_end}까지, 컬럼 {num_columns}개")
    
    # Google Sheets API 객체 준비
    spreadsheet = worksheet.spreadsheet
    sheet_id = worksheet.id
    
    # 2단계: 교차 색상 적용
    try:
        # Google Sheets API는 0-based 인덱스를 사용하며, endRowIndex는 exclusive입니다
        # 헤더 다음 행부터 마지막 데이터 행까지 적용
        # 예: header_row_num=2 (2행이 헤더), added_rows_end=999 (999행까지 데이터)
        #   → start_row_index = 2 (0-based, 3행부터 시작)
        #   → end_row_index = 999 (0-based exclusive, 999행 포함)
        start_row_index = header_row_num  # 헤더 다음 행 (0-based)
        end_row_index = added_rows_end  # 마지막 데이터 행 (0-based, exclusive이므로 999행 포함)
        
        # 기존 bandedRange 설정 확인
        existing_banded_range = None
        existing_banded_range_id = None
        try:
            # Google Sheets API를 통해 스프레드시트 메타데이터 가져오기
            try:
                from googleapiclient.discovery import build
            except ImportError:
                # googleapiclient가 없으면 기본값 사용
                raise ImportError("googleapiclient 패키지가 필요합니다. pip install google-api-python-client")
            
            # worksheet와 동일한 OAuth 자격증명 (gspread HTTPClient.auth)
            credentials = worksheet.client.auth
            
            # Google Sheets API 서비스 생성
            service = build('sheets', 'v4', credentials=credentials)
            response = service.spreadsheets().get(
                spreadsheetId=spreadsheet.id,
                fields='sheets.properties,sheets.bandedRanges'
            ).execute()
            
            # 현재 시트의 bandedRanges 찾기
            sheets_data = response.get('sheets', [])
            for sheet_data in sheets_data:
                sheet_props = sheet_data.get('properties', {})
                if sheet_props.get('sheetId') == sheet_id:
                    # 시트의 bandedRanges 확인
                    banded_ranges = sheet_data.get('bandedRanges', [])
                    if banded_ranges:
                        # 첫 번째 bandedRange 사용 (일반적으로 하나만 있음)
                        existing_banded_range = banded_ranges[0].get('bandedRange', {})
                        existing_banded_range_id = banded_ranges[0].get('bandedRangeId')
                        break
        except Exception as e:
            # 기존 설정 읽기 실패 - 기본값 사용
            print(f"ℹ️  기존 교차 색상 설정을 확인할 수 없습니다. 기본 설정을 사용합니다: {str(e)}")
        
        # 기존 bandedRange가 있고 범위를 업데이트할 수 있는 경우
        if existing_banded_range and existing_banded_range_id:
            try:
                # 기존 bandedRange의 색상 설정 가져오기
                row_props = existing_banded_range.get('rowProperties', {})
                
                # 기존 설정을 유지하면서 범위만 업데이트
                requests = [{
                    'updateBanding': {
                        'bandedRangeId': existing_banded_range_id,
                        'bandedRange': {
                            'range': {
                                'sheetId': sheet_id,
                                'startRowIndex': start_row_index,
                                'endRowIndex': end_row_index,
                                'startColumnIndex': 0,
                                'endColumnIndex': num_columns
                            },
                            'rowProperties': row_props  # 기존 색상 설정 유지
                        },
                        'fields': 'range,rowProperties'
                    }
                }]
                
                spreadsheet.batch_update({'requests': requests})
                # 0-based를 1-based로 변환하여 표시
                print(f"✅ 교차 색상 업데이트 완료 (기존 설정 유지, 행 {start_row_index + 1}-{end_row_index}, 컬럼 A-{chr(65 + num_columns - 1)})")
            except Exception as e:
                # 업데이트 실패 - 새로 생성
                print(f"ℹ️  기존 교차 색상 업데이트 실패, 새로 생성합니다: {str(e)}")
                existing_banded_range = None
        
        # 기존 bandedRange가 없거나 업데이트 실패한 경우 - 기본 설정으로 새로 생성
        if not existing_banded_range:
            # 기본 설정: 머릿글 체크 해제, 흰색/회색 교차
            # Google Sheets API에서 bands 필드는 별도로 지정하지 않고, 
            # rowProperties의 색상이 지정되면 자동으로 교차 색상이 적용됩니다.
            requests = [{
                'addBanding': {
                    'bandedRange': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': start_row_index,
                            'endRowIndex': end_row_index,
                            'startColumnIndex': 0,
                            'endColumnIndex': num_columns
                        },
                        'rowProperties': {
                            'headerColor': {
                                'red': 1.0,
                                'green': 1.0,
                                'blue': 1.0
                            },
                            'firstBandColor': {
                                'red': 1.0,
                                'green': 1.0,
                                'blue': 1.0
                            },
                            'secondBandColor': {
                                'red': 0.95,
                                'green': 0.95,
                                'blue': 0.95
                            }
                        }
                    }
                }
            }]
            
            spreadsheet.batch_update({'requests': requests})
            # 0-based를 1-based로 변환하여 표시하고, 컬럼 범위도 표시
            col_end_letter = chr(65 + num_columns - 1) if num_columns <= 26 else 'Z'  # A-Z까지만 간단히 표시
            print(f"✅ 교차 색상 적용 완료 (기본 설정, 행 {start_row_index + 1}-{end_row_index}, 컬럼 A-{col_end_letter})")
        
    except Exception as e:
        print(f"⚠️  교차 색상 적용 중 오류 발생 (데이터는 정상 추가됨): {str(e)}")
    
    # 3단계: 서식 복사
    try:
        # 마지막 데이터 행 찾기 (헤더 제외)
        source_row = header_row_num + 1  # 헤더 다음 행부터 시작
        
        if last_data_row > header_row_num:
            # 실제 데이터가 있는 경우, 마지막 데이터 행의 서식 사용
            source_row = last_data_row
        
        # 서식 복사 요청 (copyPaste 사용)
        requests = [{
            'copyPaste': {
                'source': {
                    'sheetId': sheet_id,
                    'startRowIndex': source_row - 1,  # 0-based
                    'endRowIndex': source_row,  # 0-based (다음 행 전까지)
                    'startColumnIndex': 0,
                    'endColumnIndex': num_columns
                },
                'destination': {
                    'sheetId': sheet_id,
                    'startRowIndex': added_rows_start - 1,  # 0-based
                    'endRowIndex': added_rows_end,  # 0-based
                    'startColumnIndex': 0,
                    'endColumnIndex': num_columns
                },
                'pasteType': 'PASTE_FORMAT',
                'pasteOrientation': 'NORMAL'
            }
        }]
        
        spreadsheet.batch_update({'requests': requests})
        print(f"✅ 서식 복사 완료 (행 {source_row} → 행 {added_rows_start}-{added_rows_end})")
        
    except Exception as e:
        print(f"⚠️  서식 복사 중 오류 발생 (데이터는 정상 추가됨): {str(e)}")


def sync_naver(realtime_file_path=None, test_mode=True, test_count=None):
    """
    네이버 실시간 DB와 스프레드시트 동기화
    
    Args:
        realtime_file_path: 실시간 DB 파일 경로 (None이면 최신 파일 자동 선택)
        test_mode: 테스트 모드 여부 (True면 추가하지 않음, False면 추가함)
        test_count: 테스트 추가 개수 (None이면 모든 제품, 숫자면 지정 개수만)
    """
    print("\n" + "="*80)
    print("🔵 네이버 동기화 시작")
    print("="*80)
    
    # 1단계: 실시간 DB 읽기
    if realtime_file_path is None:
        # 최신 파일 자동 선택
        latest_file = get_latest_file_from_pattern(
            NAVER_CONFIG["db_dir"],
            NAVER_CONFIG["file_pattern"]
        )
        if latest_file is None:
            print(f"❌ 오류: {NAVER_CONFIG['db_dir']}에서 {NAVER_CONFIG['file_pattern']} 파일을 찾을 수 없습니다.")
            return
        realtime_file_path = latest_file
    
    print(f"📂 실시간 DB 읽기: {realtime_file_path}")
    # 네이버는 CSV 파일이므로 첫 번째 행(0)이 헤더
    realtime_data = read_realtime_db(str(realtime_file_path), header_row=0)
    print(f"✅ 실시간 DB: {len(realtime_data)}개 제품")
    
    # 2단계: 스프레드시트 데이터 읽기
    print("\n📊 스프레드시트에서 기존 데이터 읽기...")
    gc = get_authorized_gspread_client()
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    worksheet = spreadsheet.get_worksheet(NAVER_CONFIG["sheet_index"])
    
    spreadsheet_data = get_all_spreadsheet_data(worksheet, NAVER_CONFIG["header_row"])
    print(f"✅ 기존 DB: {len(spreadsheet_data)}개 제품")
    
    # 3단계: 헤더 가져오기
    spreadsheet_headers = get_spreadsheet_headers(worksheet, NAVER_CONFIG["header_row"])
    
    # 4단계: 신제품 찾기
    print(f"\n🔍 신제품 검색 중... (옵션ID: {NAVER_CONFIG['option_id_column']})")
    
    # 기존 제품 상세 로깅
    print("\n📋 기존 제품 샘플 (처음 10개):")
    for i, row in enumerate(spreadsheet_data[:10]):
        option_id = format_option_id_for_log(row.get(NAVER_CONFIG['option_id_column'], ''))
        product_name = row.get('상품명', 'N/A')
        price = format_quantity_for_log(row.get('판매가', 'N/A'))
        stock = format_quantity_for_log(row.get('재고수량', 'N/A'))
        print(f"  [{i+1}] 상품번호: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    
    # 실시간 제품 샘플 로깅
    print("\n📋 실시간 DB 샘플 (처음 10개):")
    for i, row in enumerate(realtime_data[:10]):
        option_id = format_option_id_for_log(row.get(NAVER_CONFIG['option_id_column'], ''))
        product_name = row.get('상품명', 'N/A')
        price = format_quantity_for_log(row.get('판매가', 'N/A'))
        stock = format_quantity_for_log(row.get('재고수량', 'N/A'))
        print(f"  [{i+1}] 상품번호: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    
    # 신제품 찾기
    new_products = find_new_products(
        realtime_data, 
        spreadsheet_data, 
        NAVER_CONFIG["option_id_column"]
    )
    print(f"\n✅ 신제품 발견: {len(new_products)}개")
    
    # 신제품 상세 로깅
    if new_products:
        print("\n🆕 신제품 목록 (전부 표시):")
        for i, product in enumerate(new_products):
            option_id = format_option_id_for_log(product.get(NAVER_CONFIG['option_id_column'], ''))
            product_name = product.get('상품명', 'N/A')
            price = format_quantity_for_log(product.get('판매가', 'N/A'))
            stock = format_quantity_for_log(product.get('재고수량', 'N/A'))
            print(f"  [{i+1}] 상품번호: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    
    # 5단계: 데이터 변환 및 추가
    if new_products:
        # 테스트 개수 제한 (한 개씩만 추가 테스트용)
        products_to_add = new_products
        if test_count is not None and test_count > 0:
            products_to_add = new_products[:test_count]
            print(f"\n🧪 테스트 모드: {test_count}개 제품만 추가합니다.")
        
        if test_mode:
            print(f"\n📝 추가 예정 제품 수: {len(products_to_add)}개")
            print("⚠️  테스트 모드: 실제로 스프레드시트에 추가하지 않습니다.")
        else:
            print(f"\n📝 추가할 제품 수: {len(products_to_add)}개")
            new_rows = []
            for product in products_to_add:
                aligned_row = align_to_spreadsheet_headers(product, spreadsheet_headers)
                new_rows.append(aligned_row)
            append_to_spreadsheet(worksheet, new_rows, NAVER_CONFIG["header_row"])
    else:
        print("\n✅ 새 제품이 없습니다. 동기화 완료!")
    
    print("\n" + "="*80)


def sync_coupang(realtime_file_path=None, test_mode=True, test_count=None):
    """
    쿠팡 실시간 DB와 스프레드시트 동기화
    
    Args:
        realtime_file_path: 실시간 DB 파일 경로 (None이면 최신 파일 자동 선택)
        test_mode: 테스트 모드 여부 (True면 추가하지 않음, False면 추가함)
        test_count: 테스트 추가 개수 (None이면 모든 제품, 숫자면 지정 개수만)
    """
    print("\n" + "="*80)
    print("🟡 쿠팡 동기화 시작")
    print("="*80)
    
    # 1단계: 실시간 DB 읽기
    if realtime_file_path is None:
        # 최신 파일 자동 선택
        latest_file = get_latest_file_from_pattern(
            COUPANG_CONFIG["db_dir"],
            COUPANG_CONFIG["file_pattern"]
        )
        if latest_file is None:
            print(f"❌ 오류: {COUPANG_CONFIG['db_dir']}에서 {COUPANG_CONFIG['file_pattern']} 파일을 찾을 수 없습니다.")
            return
        realtime_file_path = latest_file
    
    print(f"📂 실시간 DB 읽기: {realtime_file_path}")
    # 쿠팡은 XLSX 파일이고 3행(인덱스 2)이 헤더
    realtime_data = read_realtime_db(str(realtime_file_path), header_row=2)
    print(f"✅ 실시간 DB: {len(realtime_data)}개 제품")
    
    # 2단계: 스프레드시트 데이터 읽기
    print("\n📊 스프레드시트에서 기존 데이터 읽기...")
    gc = get_authorized_gspread_client()
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    worksheet = spreadsheet.get_worksheet(COUPANG_CONFIG["sheet_index"])
    
    spreadsheet_data = get_all_spreadsheet_data(worksheet, COUPANG_CONFIG["header_row"])
    print(f"✅ 기존 DB: {len(spreadsheet_data)}개 제품")
    
    # 3단계: 헤더 가져오기
    spreadsheet_headers = get_spreadsheet_headers(worksheet, COUPANG_CONFIG["header_row"])
    
    # 4단계: 신제품 찾기
    print(f"\n🔍 신제품 검색 중... (옵션ID: {COUPANG_CONFIG['option_id_column']})")
    
    # 기존 제품 상세 로깅
    print("\n📋 기존 제품 샘플 (처음 10개):")
    for i, row in enumerate(spreadsheet_data[:10]):
        option_id = format_option_id_for_log(row.get(COUPANG_CONFIG['option_id_column'], ''))
        product_name = row.get('쿠팡 노출 상품명', 'N/A')
        price = format_quantity_for_log(row.get('판매가격', 'N/A'))
        stock = format_quantity_for_log(row.get('잔여수량(재고)', 'N/A'))
        print(f"  [{i+1}] 옵션ID: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    
    # 실시간 제품 샘플 로깅
    print("\n📋 실시간 DB 샘플 (처음 10개):")
    for i, row in enumerate(realtime_data[:10]):
        option_id = format_option_id_for_log(row.get(COUPANG_CONFIG['option_id_column'], ''))
        product_name = row.get('쿠팡 노출 상품명', 'N/A')
        price = format_quantity_for_log(row.get('판매가격', 'N/A'))
        stock = format_quantity_for_log(row.get('잔여수량(재고)', 'N/A'))
        print(f"  [{i+1}] 옵션ID: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    
    # 신제품 찾기
    new_products = find_new_products(
        realtime_data, 
        spreadsheet_data, 
        COUPANG_CONFIG["option_id_column"]
    )
    print(f"\n✅ 신제품 발견: {len(new_products)}개")
    
    # 신제품 상세 로깅
    if new_products:
        print("\n🆕 신제품 목록 (전부 표시):")
        for i, product in enumerate(new_products):
            option_id = format_option_id_for_log(product.get(COUPANG_CONFIG['option_id_column'], ''))
            product_name = product.get('쿠팡 노출 상품명', 'N/A')
            price = format_quantity_for_log(product.get('판매가격', 'N/A'))
            stock = format_quantity_for_log(product.get('잔여수량(재고)', 'N/A'))
            print(f"  [{i+1}] 옵션ID: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    
    # 5단계: 데이터 변환 및 추가
    if new_products:
        # 테스트 개수 제한 (한 개씩만 추가 테스트용)
        products_to_add = new_products
        if test_count is not None and test_count > 0:
            products_to_add = new_products[:test_count]
            print(f"\n🧪 테스트 모드: {test_count}개 제품만 추가합니다.")
        
        if test_mode:
            print(f"\n📝 추가 예정 제품 수: {len(products_to_add)}개")
            print("⚠️  테스트 모드: 실제로 스프레드시트에 추가하지 않습니다.")
        else:
            print(f"\n📝 추가할 제품 수: {len(products_to_add)}개")
            new_rows = []
            for product in products_to_add:
                aligned_row = align_to_spreadsheet_headers(product, spreadsheet_headers)
                new_rows.append(aligned_row)
            append_to_spreadsheet_coupang(worksheet, new_rows, COUPANG_CONFIG["header_row"])
    else:
        print("\n✅ 새 제품이 없습니다. 동기화 완료!")
    
    print("\n" + "="*80)


# ===== 메인 실행 =====

if __name__ == "__main__":
    print("="*80)
    print("🛒 상품 DB 동기화 스크립트")
    print("="*80)
    print("\n이 스크립트는 database 폴더에서 최신 파일을 자동으로 찾아 동기화합니다.")
    print("="*80)
    
    # 테스트 모드 설정
    # 로깅만 하려면: test_mode=True, test_count=None
    # 한 개씩만 추가 테스트하려면: test_mode=False, test_count=1
    # 전체 추가하려면: test_mode=False, test_count=None
    
    print("\n" + "="*80)
    print("🧪 테스트 모드: 네이버 1개, 쿠팡 1개 제품 추가 테스트")
    print("="*80)
    
    # 네이버: 1개 제품만 실제 추가
    # sync_naver(test_mode=False, test_count=1)
    sync_naver(test_mode=False, test_count=None)
    
    # 쿠팡: 1개 제품만 실제 추가
    # sync_coupang(test_mode=False, test_count=1)
    sync_coupang(test_mode=False, test_count=None)
    
    print("\n" + "="*80)
    print("✅ 테스트 완료!")
    print("\n💡 모든 제품 추가하려면 test_mode=False, test_count=None으로 변경하세요")

