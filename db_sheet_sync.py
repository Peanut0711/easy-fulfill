"""
로컬 database 폴더 내보내기(CSV/XLSX)와 구글 스프레드시트(신규 행 append) 동기화.

인증: google-oauth/credentials.json + token.json — google_sheets_oauth.py
easy-fulfill UI·database-sync CLI 모두 이 모듈을 사용합니다.
"""

import gspread
import pandas as pd
from pathlib import Path

from google_sheets_oauth import get_authorized_gspread_client

# 기본 스프레드시트 ID (앱·CLI에서 동일 값 사용)
DEFAULT_SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"

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


def apply_banded_rows(spreadsheet, sheet_id, header_row_num, last_data_row, num_columns, log=print):
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
        log(f"✅ 교차 색상 적용 완료 (행 {start_row_index + 1}-{end_row_index})")
        
    except Exception as e:
        log(f"⚠️  교차 색상 적용 중 오류 발생: {str(e)}")


def append_to_spreadsheet(
    worksheet, new_data_list, header_row_num, log=print, *, verbose_log=True
):
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
    
    log(f"✅ {len(new_data_list)}개의 새 제품이 추가되었습니다 (행 {added_rows_start}-{added_rows_end})")
    
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
        if verbose_log:
            log(f"✅ 서식 복사 완료 (행 {source_row} → 행 {added_rows_start}-{added_rows_end})")
        
    except Exception as e:
        log(f"⚠️  서식 복사 중 오류 발생 (데이터는 정상 추가됨): {str(e)}")


def append_to_spreadsheet_coupang(
    worksheet, new_data_list, header_row_num, log=print, *, verbose_log=True
):
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
    
    log(f"✅ {len(new_data_list)}개의 새 제품이 추가되었습니다 (행 {added_rows_start}-{added_rows_end})")
    if verbose_log:
        log(f"ℹ️  교차 색상 적용 범위: 헤더 행({header_row_num}) 다음부터 행 {added_rows_end}까지, 컬럼 {num_columns}개")
    
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
            if verbose_log:
                log(f"ℹ️  기존 교차 색상 설정을 확인할 수 없습니다. 기본 설정을 사용합니다: {str(e)}")
        
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
                if verbose_log:
                    # 0-based를 1-based로 변환하여 표시
                    log(f"✅ 교차 색상 업데이트 완료 (기존 설정 유지, 행 {start_row_index + 1}-{end_row_index}, 컬럼 A-{chr(65 + num_columns - 1)})")
            except Exception as e:
                # 업데이트 실패 - 새로 생성
                if verbose_log:
                    log(f"ℹ️  기존 교차 색상 업데이트 실패, 새로 생성합니다: {str(e)}")
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
            if verbose_log:
                # 0-based를 1-based로 변환하여 표시하고, 컬럼 범위도 표시
                col_end_letter = chr(65 + num_columns - 1) if num_columns <= 26 else 'Z'  # A-Z까지만 간단히 표시
                log(f"✅ 교차 색상 적용 완료 (기본 설정, 행 {start_row_index + 1}-{end_row_index}, 컬럼 A-{col_end_letter})")
        
    except Exception as e:
        log(f"⚠️  교차 색상 적용 중 오류 발생 (데이터는 정상 추가됨): {str(e)}")
    
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
        if verbose_log:
            log(f"✅ 서식 복사 완료 (행 {source_row} → 행 {added_rows_start}-{added_rows_end})")
        
    except Exception as e:
        log(f"⚠️  서식 복사 중 오류 발생 (데이터는 정상 추가됨): {str(e)}")


def _empty_channel_result(channel):
    return {
        "ok": False,
        "channel": channel,
        "error": None,
        "skipped_missing_file": False,
        "resolved_file": None,
        "realtime_count": 0,
        "sheet_count": 0,
        "new_products_count": 0,
        "planned_append": 0,
        "appended": 0,
    }


def _log_db_sync_job_summary(log, do_naver, do_coupang, out, *, verbose_log=True):
    """선택한 채널별로 진행/건너뜀/실패를 한 블록으로 정리 (UI 로그 가독성)."""
    requested = []
    if do_naver:
        requested.append(("네이버", "naver"))
    if do_coupang:
        requested.append(("쿠팡", "coupang"))
    done_labels = []
    skipped_labels = []
    failed_labels = []
    for label, key in requested:
        ch = out.get(key)
        if not ch:
            continue
        if ch.get("skipped_missing_file"):
            skipped_labels.append(label)
        elif ch.get("ok"):
            done_labels.append(label)
        else:
            failed_labels.append(label)
    log("\n" + ("=" * 80 if verbose_log else "─" * 40))
    log("📌 작업 요약")
    if done_labels:
        log(f"   ✓ 동기화 진행: {', '.join(done_labels)}")
    if skipped_labels:
        log(
            f"   ⏭ 로컬 파일 없음 → 건너뜀: {', '.join(skipped_labels)} "
            "(다른 채널은 정상적으로 처리되었습니다)"
        )
    if failed_labels:
        log(f"   ✗ 처리 실패: {', '.join(failed_labels)}")
    if not done_labels and not failed_labels and skipped_labels:
        log("   → 선택한 채널 모두 해당 패턴 파일이 없어 시트 반영 없음")
    log("=" * 80 if verbose_log else "─" * 40)


def sync_naver(
    spreadsheet_id,
    realtime_file_path=None,
    test_mode=True,
    test_count=None,
    log=print,
    *,
    verbose_log=True,
):
    """네이버 실시간 DB와 스프레드시트 동기화. dict 결과는 UI/CLI 공통 사용."""
    r = _empty_channel_result("naver")
    if verbose_log:
        log("\n" + "=" * 80)
        log("🔵 네이버 동기화 시작")
        log("=" * 80)
    else:
        log("\n🔵 네이버 동기화")

    if realtime_file_path is None:
        latest_file = get_latest_file_from_pattern(
            NAVER_CONFIG["db_dir"],
            NAVER_CONFIG["file_pattern"],
        )
        if latest_file is None:
            log(
                f"⏭️ 네이버 건너뜀: {NAVER_CONFIG['db_dir']} 폴더에 "
                f"{NAVER_CONFIG['file_pattern']} 에 맞는 파일이 없습니다."
            )
            log(
                "   (네이버만 건너뜁니다. 쿠팡 등 다른 선택 채널 동기화는 계속 진행됩니다.)"
            )
            r["skipped_missing_file"] = True
            r["ok"] = True
            if verbose_log:
                log("\n" + "=" * 80)
            return r
        realtime_file_path = latest_file

    r["resolved_file"] = str(realtime_file_path)
    log(f"📂 실시간 DB 읽기: {realtime_file_path}")
    realtime_data = read_realtime_db(str(realtime_file_path), header_row=0)
    r["realtime_count"] = len(realtime_data)
    log(f"✅ 실시간 DB: {len(realtime_data)}개 제품")

    log("\n📊 스프레드시트에서 기존 데이터 읽기...")
    gc = get_authorized_gspread_client()
    spreadsheet = gc.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.get_worksheet(NAVER_CONFIG["sheet_index"])

    spreadsheet_data = get_all_spreadsheet_data(worksheet, NAVER_CONFIG["header_row"])
    r["sheet_count"] = len(spreadsheet_data)
    log(f"✅ 기존 DB: {len(spreadsheet_data)}개 제품")

    spreadsheet_headers = get_spreadsheet_headers(worksheet, NAVER_CONFIG["header_row"])

    if verbose_log:
        log(f"\n🔍 신제품 검색 중... (키: {NAVER_CONFIG['option_id_column']})")
        log("\n📋 기존 제품 샘플 (처음 10개):")
        for i, row in enumerate(spreadsheet_data[:10]):
            option_id = format_option_id_for_log(row.get(NAVER_CONFIG["option_id_column"], ""))
            product_name = row.get("상품명", "N/A")
            price = format_quantity_for_log(row.get("판매가", "N/A"))
            stock = format_quantity_for_log(row.get("재고수량", "N/A"))
            log(f"  [{i+1}] 상품번호: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")

        log("\n📋 실시간 DB 샘플 (처음 10개):")
        for i, row in enumerate(realtime_data[:10]):
            option_id = format_option_id_for_log(row.get(NAVER_CONFIG["option_id_column"], ""))
            product_name = row.get("상품명", "N/A")
            price = format_quantity_for_log(row.get("판매가", "N/A"))
            stock = format_quantity_for_log(row.get("재고수량", "N/A"))
            log(f"  [{i+1}] 상품번호: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    else:
        log("\n🔍 신제품 비교 중...")

    new_products = find_new_products(
        realtime_data,
        spreadsheet_data,
        NAVER_CONFIG["option_id_column"],
    )
    r["new_products_count"] = len(new_products)
    log(f"\n✅ 신제품 발견: {len(new_products)}개")

    if new_products:
        if verbose_log:
            log("\n🆕 신제품 목록 (전부 표시):")
            for i, product in enumerate(new_products):
                option_id = format_option_id_for_log(product.get(NAVER_CONFIG["option_id_column"], ""))
                product_name = product.get("상품명", "N/A")
                price = format_quantity_for_log(product.get("판매가", "N/A"))
                stock = format_quantity_for_log(product.get("재고수량", "N/A"))
                log(f"  [{i+1}] 상품번호: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
        else:
            log(
                f"\n🆕 신제품 {len(new_products)}건 — 행 목록은 「로그 출력(상세)」를 켠 뒤 다시 실행하면 표시됩니다."
            )

    if new_products:
        products_to_add = new_products
        if test_count is not None and test_count > 0:
            products_to_add = new_products[:test_count]
            log(f"\n🧪 건수 제한: 상위 {test_count}개만 대상으로 합니다.")
        r["planned_append"] = len(products_to_add)

        if test_mode:
            log(f"\n📝 추가 예정 제품 수: {len(products_to_add)}개")
            log("⚠️  미리보기: 스프레드시트에 쓰지 않습니다.")
        else:
            log(f"\n📝 추가할 제품 수: {len(products_to_add)}개")
            new_rows = []
            for product in products_to_add:
                aligned_row = align_to_spreadsheet_headers(product, spreadsheet_headers)
                new_rows.append(aligned_row)
            append_to_spreadsheet(
                worksheet,
                new_rows,
                NAVER_CONFIG["header_row"],
                log=log,
                verbose_log=verbose_log,
            )
            r["appended"] = len(new_rows)
    else:
        log("\n✅ 새 제품이 없습니다. 동기화 완료!")

    r["ok"] = r["error"] is None
    if verbose_log:
        log("\n" + "=" * 80)
    return r


def sync_coupang(
    spreadsheet_id,
    realtime_file_path=None,
    test_mode=True,
    test_count=None,
    log=print,
    *,
    verbose_log=True,
):
    """쿠팡 실시간 DB와 스프레드시트 동기화."""
    r = _empty_channel_result("coupang")
    if verbose_log:
        log("\n" + "=" * 80)
        log("🟡 쿠팡 동기화 시작")
        log("=" * 80)
    else:
        log("\n🟡 쿠팡 동기화")

    if realtime_file_path is None:
        latest_file = get_latest_file_from_pattern(
            COUPANG_CONFIG["db_dir"],
            COUPANG_CONFIG["file_pattern"],
        )
        if latest_file is None:
            log(
                f"⏭️ 쿠팡 건너뜀: {COUPANG_CONFIG['db_dir']} 폴더에 "
                f"{COUPANG_CONFIG['file_pattern']} 에 맞는 파일이 없습니다."
            )
            log(
                "   (쿠팡만 건너뜁니다. 네이버 등 다른 선택 채널 동기화는 계속 진행됩니다.)"
            )
            r["skipped_missing_file"] = True
            r["ok"] = True
            if verbose_log:
                log("\n" + "=" * 80)
            return r
        realtime_file_path = latest_file

    r["resolved_file"] = str(realtime_file_path)
    log(f"📂 실시간 DB 읽기: {realtime_file_path}")
    realtime_data = read_realtime_db(str(realtime_file_path), header_row=2)
    r["realtime_count"] = len(realtime_data)
    log(f"✅ 실시간 DB: {len(realtime_data)}개 제품")

    log("\n📊 스프레드시트에서 기존 데이터 읽기...")
    gc = get_authorized_gspread_client()
    spreadsheet = gc.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.get_worksheet(COUPANG_CONFIG["sheet_index"])

    spreadsheet_data = get_all_spreadsheet_data(worksheet, COUPANG_CONFIG["header_row"])
    r["sheet_count"] = len(spreadsheet_data)
    log(f"✅ 기존 DB: {len(spreadsheet_data)}개 제품")

    spreadsheet_headers = get_spreadsheet_headers(worksheet, COUPANG_CONFIG["header_row"])

    if verbose_log:
        log(f"\n🔍 신제품 검색 중... (키: {COUPANG_CONFIG['option_id_column']})")
        log("\n📋 기존 제품 샘플 (처음 10개):")
        for i, row in enumerate(spreadsheet_data[:10]):
            option_id = format_option_id_for_log(row.get(COUPANG_CONFIG["option_id_column"], ""))
            product_name = row.get("쿠팡 노출 상품명", "N/A")
            price = format_quantity_for_log(row.get("판매가격", "N/A"))
            stock = format_quantity_for_log(row.get("잔여수량(재고)", "N/A"))
            log(f"  [{i+1}] 옵션ID: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")

        log("\n📋 실시간 DB 샘플 (처음 10개):")
        for i, row in enumerate(realtime_data[:10]):
            option_id = format_option_id_for_log(row.get(COUPANG_CONFIG["option_id_column"], ""))
            product_name = row.get("쿠팡 노출 상품명", "N/A")
            price = format_quantity_for_log(row.get("판매가격", "N/A"))
            stock = format_quantity_for_log(row.get("잔여수량(재고)", "N/A"))
            log(f"  [{i+1}] 옵션ID: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
    else:
        log("\n🔍 신제품 비교 중...")

    new_products = find_new_products(
        realtime_data,
        spreadsheet_data,
        COUPANG_CONFIG["option_id_column"],
    )
    r["new_products_count"] = len(new_products)
    log(f"\n✅ 신제품 발견: {len(new_products)}개")

    if new_products:
        if verbose_log:
            log("\n🆕 신제품 목록 (전부 표시):")
            for i, product in enumerate(new_products):
                option_id = format_option_id_for_log(product.get(COUPANG_CONFIG["option_id_column"], ""))
                product_name = product.get("쿠팡 노출 상품명", "N/A")
                price = format_quantity_for_log(product.get("판매가격", "N/A"))
                stock = format_quantity_for_log(product.get("잔여수량(재고)", "N/A"))
                log(f"  [{i+1}] 옵션ID: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")
        else:
            log(
                f"\n🆕 신제품 {len(new_products)}건 — 행 목록은 「로그 출력(상세)」를 켠 뒤 다시 실행하면 표시됩니다."
            )

    if new_products:
        products_to_add = new_products
        if test_count is not None and test_count > 0:
            products_to_add = new_products[:test_count]
            log(f"\n🧪 건수 제한: 상위 {test_count}개만 대상으로 합니다.")
        r["planned_append"] = len(products_to_add)

        if test_mode:
            log(f"\n📝 추가 예정 제품 수: {len(products_to_add)}개")
            log("⚠️  미리보기: 스프레드시트에 쓰지 않습니다.")
        else:
            log(f"\n📝 추가할 제품 수: {len(products_to_add)}개")
            new_rows = []
            for product in products_to_add:
                aligned_row = align_to_spreadsheet_headers(product, spreadsheet_headers)
                new_rows.append(aligned_row)
            append_to_spreadsheet_coupang(
                worksheet,
                new_rows,
                COUPANG_CONFIG["header_row"],
                log=log,
                verbose_log=verbose_log,
            )
            r["appended"] = len(new_rows)
    else:
        log("\n✅ 새 제품이 없습니다. 동기화 완료!")

    r["ok"] = r["error"] is None
    if verbose_log:
        log("\n" + "=" * 80)
    return r


def run_db_sheet_sync_job(
    spreadsheet_id,
    *,
    do_naver,
    do_coupang,
    test_mode,
    test_count,
    naver_path=None,
    coupang_path=None,
    verbose_log=True,
):
    """UI 스레드 밖에서 호출. test_count: 0이면 무제한(None과 동일). verbose_log: 샘플·상세 append 로그."""
    logs = []

    def log(msg):
        logs.append(str(msg))

    out = {
        "ok": True,
        "logs": logs,
        "test_mode": test_mode,
        "naver": None,
        "coupang": None,
        "error": None,
    }
    naver_path_arg = None if not naver_path else str(naver_path)
    coupang_path_arg = None if not coupang_path else str(coupang_path)
    tc = None if test_count is None or test_count <= 0 else int(test_count)

    if verbose_log:
        log("—" * 40)
        if test_mode:
            log("🧪 미리보기: 스프레드시트에 행을 추가·수정하지 않습니다.")
        else:
            log("📝 실반영: 신규 상품이 있으면 시트에 행이 추가됩니다.")
        log("—" * 40)
    else:
        if test_mode:
            log("▶ 미리보기 · 시트 미적용")
        else:
            log("▶ 실반영 · 신규 행 시트 추가 가능")

    try:
        if do_naver:
            out["naver"] = sync_naver(
                spreadsheet_id,
                realtime_file_path=naver_path_arg,
                test_mode=test_mode,
                test_count=tc,
                log=log,
                verbose_log=verbose_log,
            )
            if not out["naver"].get("ok"):
                out["ok"] = False
        if do_coupang:
            out["coupang"] = sync_coupang(
                spreadsheet_id,
                realtime_file_path=coupang_path_arg,
                test_mode=test_mode,
                test_count=tc,
                log=log,
                verbose_log=verbose_log,
            )
            if not out["coupang"].get("ok"):
                out["ok"] = False
        _log_db_sync_job_summary(log, do_naver, do_coupang, out, verbose_log=verbose_log)
    except Exception as e:
        out["ok"] = False
        out["error"] = str(e)
        log(f"❌ 예외: {e}")
    return out
