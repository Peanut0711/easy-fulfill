"""
네이버·쿠팡 상품 목록(API 또는 로컬 database 폴더 내보내기 CSV/XLSX)과 구글 스프레드시트(신규 행 append) 동기화.

인증: google-oauth/credentials.json + token.json — google_sheets_oauth.py
easy-fulfill UI·database-sync CLI 모두 이 모듈을 사용합니다.
"""

import re
from datetime import date, datetime

from pathlib import Path

from shared.lazy_import import deferred_module
from shared.paths import APP_ROOT

# 앱 시작 시 경로 표시만 할 때는 pandas·gspread·Google 인증 모듈을 불러오지 않는다.
pd = deferred_module("pandas")
gspread = deferred_module("gspread")


def get_authorized_gspread_client(*args, **kwargs):
    from .google_sheets_oauth import get_authorized_gspread_client as authorize
    return authorize(*args, **kwargs)

# 기본 스프레드시트 ID (앱·CLI에서 동일 값 사용)
DEFAULT_SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"

# 네이버/쿠팡 설정
NAVER_CONFIG = {
    "sheet_index": 0,  # 1번 시트
    "header_row": 1,   # 1행이 헤더
    "option_id_column": "상품번호(스마트스토어)",  # 네이버는 옵션ID가 없고 상품번호를 사용
    "db_dir": str(APP_ROOT / "database"),
    # 구 스마트스토어 내보내기: Product_*.csv / 신규: 스마트스토어상품_YYYYMMDD_HHMMSS.csv
    "file_patterns": ["Product_*.csv", "스마트스토어상품_*.csv"],
}

COUPANG_CONFIG = {
    "sheet_index": 1,  # 2번 시트
    "header_row": 1,   # 1행이 헤더(예전엔 2행). 실제 헤더 행은 detect_header_row로 찾는다
    "option_id_column": "옵션 ID",
    "db_dir": str(APP_ROOT / "database"),
    "file_pattern": "price_inventory_*.xlsx"
}


# 상품 목록을 가져올 곳: API(기본) 또는 database 폴더의 내보내기 파일
SOURCE_API = "api"
SOURCE_FILE = "file"

# 네이버 API 상태 코드 → 내보내기 파일·시트에 쓰인 글자
NAVER_STATUS_LABELS = {
    "SALE": "판매중", "OUTOFSTOCK": "품절", "SUSPENSION": "판매중지", "WAIT": "판매대기",
    "CLOSE": "판매종료", "PROHIBITION": "판매금지", "UNADMISSION": "승인대기",
    "REJECTION": "승인거부", "DELETE": "삭제",
}
NAVER_DISPLAY_LABELS = {"ON": "전시중", "SUSPENSION": "전시중지", "WAIT": "전시대기"}
COUPANG_OFFER_LABELS = {"NEW": "새상품", "REFURBISHED": "리퍼", "USED_BEST": "중고(최상)",
                        "USED_GOOD": "중고(상)", "USED_NORMAL": "중고(중)"}


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


def get_latest_file_from_patterns(directory, file_patterns):
    """
    디렉토리에서 여러 glob 패턴에 맞는 파일 중 수정 시각이 가장 최근인 경로 반환.

    Args:
        directory: 디렉토리 경로
        file_patterns: glob 문자열 리스트 (예: ["Product_*.csv", "스마트스토어상품_*.csv"])

    Returns:
        Path | None
    """
    from glob import glob
    import os

    files = []
    for fp in file_patterns:
        files.extend(glob(os.path.join(directory, fp)))
    if not files:
        return None
    files = list(dict.fromkeys(files))
    latest_file = max(files, key=os.path.getmtime)
    return Path(latest_file)


def try_parse_db_sync_filename_date(filename: str):
    """
    네이버·쿠팡 DB 내보내기 파일명에서 (년, 월, 일) 추출. 인식 실패 시 None.

    - Product_YYYYMMDD_… / 스마트스토어상품_YYYYMMDD_… .csv
    - price_inventory_YYMMDD.xlsx · price_inventory_YYYYMMDD.xlsx
    """
    m = re.match(
        r"^(?:Product_|스마트스토어상품_)(\d{8})_\d{6}\.csv$",
        filename,
    )
    if m:
        s = m.group(1)
        return int(s[:4]), int(s[4:6]), int(s[6:8])
    m = re.match(r"^price_inventory_(\d{6})\.xlsx$", filename, re.I)
    if m:
        s = m.group(1)
        return 2000 + int(s[:2]), int(s[2:4]), int(s[4:6])
    m = re.match(r"^price_inventory_(\d{8})\.xlsx$", filename, re.I)
    if m:
        s = m.group(1)
        return int(s[:4]), int(s[4:6]), int(s[6:8])
    return None


def db_sync_file_date(path: Path):
    """파일 날짜 ``date``: 파일 이름의 날짜, 없으면 수정 시각. 둘 다 모르면 None."""
    ymd = try_parse_db_sync_filename_date(path.name)
    if ymd is not None:
        try:
            return date(*ymd)
        except ValueError:
            pass
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).date()
    except OSError:
        return None


def format_db_sync_label_line(channel_title: str, path: Path | None, today: date | None = None) -> str:
    """
    「설정 › DB 동기화」 파일 한 줄: 「네이버 DB · Product_20261117_093000.csv · 11-17 (오늘)」.

    날짜는 파일 이름의 날짜(없으면 수정 시각)와 며칠 전인지(오늘·어제·n일 전)다. 전체 경로는 화면에서
    툴팁으로 보인다. 파일이 없으면 「네이버 DB · 파일 없음」.

    Args:
        channel_title: 예) "네이버 DB", "쿠팡 DB"
        path: 최신 파일 경로 (없으면 None)
        today: 기준 날짜(테스트용, 기본은 오늘)
    """
    if path is None:
        return f"{channel_title} · 파일 없음"
    file_date = db_sync_file_date(path)
    if file_date is None:
        return f"{channel_title} · {path.name} · 날짜 모름"
    days = ((today or date.today()) - file_date).days
    if days == 0:
        ago = "오늘"
    elif days == 1:
        ago = "어제"
    elif days > 1:
        ago = f"{days}일 전"
    else:
        ago = "미래 날짜"
    return f"{channel_title} · {path.name} · {file_date:%m-%d} ({ago})"


def detect_header_row(values, default_header_row, first_header="상품코드", max_rows=3):
    """A열이 「상품코드」인 행(위에서 max_rows행 안)을 헤더 행(1-based)으로 본다. 없으면 기본값.

    운영 시트 쿠팡 탭은 처음 내보내기 파일처럼 2행이 헤더였다가 지금은 1행이 헤더다(2026-11-19 확인).
    """
    for index, row in enumerate(values[:max_rows]):
        if row and str(row[0]).strip() == first_header:
            return index + 1
    return default_header_row


def records_from_values(values, header_row_num):
    """시트 값(get_all_values)에서 헤더 다음 행들을 dict 목록으로 만든다.

    같은 헤더가 두 번 나오면 뒤의 것은 버리고(쿠팡 탭의 「판매가격」 두 칸), 빈 헤더는 「빈컬럼_n」으로 둔다.
    """
    headers = values[header_row_num - 1] if len(values) >= header_row_num else []
    unique_headers = []
    column_indices_to_keep = []
    used_headers = set()
    for idx, header in enumerate(headers):
        name = header if header and header.strip() else f"빈컬럼_{idx}"
        if name in used_headers:
            continue
        used_headers.add(name)
        unique_headers.append(name)
        column_indices_to_keep.append(idx)

    records = []
    for row in values[header_row_num:]:
        record = {}
        for name, col_idx in zip(unique_headers, column_indices_to_keep):
            record[name] = row[col_idx] if col_idx < len(row) else ""
        records.append(record)
    return records


def read_sheet_table(worksheet, default_header_row):
    """탭을 한 번 읽어 (헤더 목록, 데이터 dict 목록, 헤더 행 번호)를 반환한다."""
    values = worksheet.get_all_values()
    header_row = detect_header_row(values, default_header_row)
    headers = values[header_row - 1] if len(values) >= header_row else []
    return headers, records_from_values(values, header_row), header_row


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
            existing_option_ids.add(option_id)  # 같은 키가 목록에 두 번 있어도 한 번만 추가
            new_products.append(row)
    
    return new_products


def align_to_spreadsheet_headers(realtime_dict, spreadsheet_headers):
    """
    실시간 DB 데이터를 스프레드시트 헤더 순서로 정렬

    같은 헤더가 다시 나오면 pandas가 붙이는 이름(「판매가격.1」)의 값을 쓴다.

    Args:
        realtime_dict: 실시간 DB 딕셔너리
        spreadsheet_headers: 스프레드시트 헤더 리스트

    Returns:
        list: 정렬된 데이터 리스트
    """
    import math

    aligned_row = []
    seen = {}
    for header in spreadsheet_headers:
        count = seen.get(header, 0)
        seen[header] = count + 1
        # 상품코드 컬럼은 사용자가 직접 창고 위치를 입력하는 곳이므로
        # 신규 제품 추가 시 반드시 빈칸으로 설정
        if header == "상품코드" or not header:
            value = ""
        else:
            value = realtime_dict.get(header if count == 0 else f"{header}.{count}", "")

        # NaN 처리: Google Sheets에서 NaN을 JSON으로 변환할 수 없음
        if value is None or (isinstance(value, float) and math.isnan(value)):
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
        "source": None,
        "skipped_missing_file": False,
        "resolved_file": None,
        "realtime_count": 0,
        "sheet_count": 0,
        "new_products_count": 0,
        "planned_append": 0,
        "appended": 0,
        "skipped_concurrent": 0,
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


# ===== API로 상품 목록 가져오기 =====

def _api_datetime_text(value):
    """API 시각(2026-11-19T12:32:04.102+09:00)을 내보내기 파일 형식(2026-11-19 12:32:04)으로."""
    text = str(value or "")
    return text[:19].replace("T", " ") if len(text) >= 19 else text


def _int_or_blank(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return ""


def naver_rows_from_api(products):
    """네이버 채널 상품(API 응답 dict)을 시트 헤더 이름의 dict로 바꾼다.

    시트의 열은 스마트스토어 내보내기 파일과 같다. 주문 처리에 쓰는 키는 「상품번호(스마트스토어)」이고,
    API에 없는 칸(포인트·배송 속성 등)은 비워 둔다.
    """
    rows = []
    for product in products:
        if product.get("channelServiceType") not in (None, "STOREFARM"):
            continue
        product_no = _int_or_blank(product.get("channelProductNo"))
        if product_no == "":
            continue
        categories = [part.strip() for part in str(product.get("wholeCategoryName") or "").split(">")]
        categories += [""] * (4 - len(categories))
        image = product.get("representativeImage") or {}
        rows.append({
            "판매자상품코드": product.get("sellerManagementCode") or "",
            "상품명": str(product.get("name") or "").strip(),
            "대표이미지 URL": image.get("url") or "",
            "상품번호(스마트스토어)": product_no,
            "채널": "스마트스토어",
            "판매상태": NAVER_STATUS_LABELS.get(product.get("statusType"), product.get("statusType") or ""),
            "전시상태": NAVER_DISPLAY_LABELS.get(
                product.get("channelProductDisplayStatusType"),
                product.get("channelProductDisplayStatusType") or ""),
            "재고수량": _int_or_blank(product.get("stockQuantity")),
            "판매가": _int_or_blank(product.get("salePrice")),
            "할인가": _int_or_blank(product.get("discountedPrice")),
            "기본배송비": _int_or_blank(product.get("deliveryFee")),
            "반품배송비": _int_or_blank(product.get("returnFee")),
            "교환배송비": _int_or_blank(product.get("exchangeFee")),
            "대분류": categories[0],
            "중분류": categories[1],
            "소분류": categories[2],
            "세분류": categories[3],
            "제조사명": product.get("manufacturerName") or "",
            "브랜드명": product.get("brandName") or "",
            "상품등록일": _api_datetime_text(product.get("regDate")),
            "최종수정일": _api_datetime_text(product.get("modifiedDate")),
        })
    return rows


def coupang_rows_from_api(seller_products):
    """쿠팡 등록 상품 상세(API 응답 data)를 옵션마다 시트 헤더 이름의 dict로 바꾼다.

    시트의 열은 쿠팡 가격·재고 내보내기 파일과 같다. 주문 처리에 쓰는 키는 「옵션 ID」(vendorItemId)와
    「Product ID」(VP 상품번호)다. 승인 전이라 옵션 ID가 없는 옵션은 건너뛴다.
    재고·판매상태는 상세 응답에 없어 비워 둔다.
    """
    rows = []
    for product in seller_products:
        display_name = str(product.get("displayProductName") or product.get("sellerProductName") or "").strip()
        for item in product.get("items") or []:
            option_id = _int_or_blank(item.get("vendorItemId"))
            if option_id == "":
                continue
            item_name = str(item.get("itemName") or "").strip()
            rows.append({
                "쿠팡 노출 상품명": ", ".join(part for part in (display_name, item_name) if part),
                "업체상품 ID": _int_or_blank(product.get("sellerProductId")),
                "Product ID": _int_or_blank(product.get("productId")),
                "옵션 ID": option_id,
                "상품상태": COUPANG_OFFER_LABELS.get(item.get("offerCondition"), item.get("offerCondition") or ""),
                "바코드": item.get("barcode") or "",
                "업체상품코드": item.get("externalVendorSku") or "",
                "업체 등록 상품명": str(product.get("sellerProductName") or "").strip(),
                "등록 옵션명": item_name,
                "판매가격": _int_or_blank(item.get("salePrice")),
                "할인율기준가": _int_or_blank(item.get("originalPrice")),
                "승인상태": product.get("statusName") or "",
            })
    return rows


def _read_shared_config(spreadsheet):
    """공유 「설정」 탭의 키/값. API 키는 여기에만 있다(코드·저장소에 두지 않는다)."""
    from tracking.repository import config_values_map_from_rows

    return config_values_map_from_rows(spreadsheet.worksheet("설정").get_all_values())


def _require_keys(config, keys, label):
    values = [str(config.get(key) or "").strip() for key in keys]
    if not all(values):
        raise RuntimeError(f"{label} API 키가 등록되지 않았습니다. 「⚙ 설정 › 연동 › API 키」에서 등록하세요.")
    return values


def load_naver_products_from_api(config, sheet_records, log):
    from integrations import naver_commerce

    client_id, client_secret = _require_keys(config, ("naver_client_id", "naver_client_secret"), "네이버")
    token = naver_commerce.get_access_token(client_id, client_secret)
    products = naver_commerce.fetch_channel_products(token)
    log(f"🌐 네이버 API: 스마트스토어 상품 {len(products)}개")
    return naver_rows_from_api(products)


def load_coupang_products_from_api(config, sheet_records, log):
    """시트 「업체상품 ID」에 없는 등록 상품만 상세 조회해 옵션 행을 만든다.

    상세 조회는 상품마다 한 번(약 0.7초)이라 전체를 읽으면 수 분이 걸린다. 이미 시트에 있는 상품에 나중에
    추가된 옵션은 이 방식으로는 찾지 못한다(docs/known-risks.md).
    """
    from integrations import coupang_commerce

    vendor_id, access_key, secret_key = _require_keys(
        config, ("coupang_vendor_id", "coupang_access_key", "coupang_secret_key"), "쿠팡")
    listed = coupang_commerce.fetch_seller_products(vendor_id, access_key, secret_key)
    known = {normalize_option_id(row.get("업체상품 ID")) for row in sheet_records}
    known.discard("")
    targets = []
    for product in listed:
        seller_product_id = normalize_option_id(product.get("sellerProductId"))
        if seller_product_id and seller_product_id not in known:
            known.add(seller_product_id)
            targets.append(seller_product_id)
    log(f"🌐 쿠팡 API: 등록 상품 {len(listed)}개 중 시트에 없는 상품 {len(targets)}개 상세 조회")
    details = [coupang_commerce.fetch_seller_product(spid, access_key, secret_key) for spid in targets]
    rows = coupang_rows_from_api(details)
    missing_option = sum(
        1 for detail in details for item in (detail.get("items") or []) if not item.get("vendorItemId"))
    if missing_option:
        log(f"   (옵션 ID가 아직 없는 옵션 {missing_option}개는 건너뜀 — 승인 뒤 다시 실행)")
    return rows


# ===== 채널 동기화 =====

CHANNELS = {
    "naver": {
        "label": "네이버", "icon": "🔵", "config": NAVER_CONFIG, "file_header_row": 0,
        "id_label": "상품번호", "name": "상품명", "price": "판매가", "stock": "재고수량",
        "api_loader": load_naver_products_from_api, "append": "append_to_spreadsheet",
    },
    "coupang": {
        "label": "쿠팡", "icon": "🟡", "config": COUPANG_CONFIG, "file_header_row": 2,
        "id_label": "옵션ID", "name": "쿠팡 노출 상품명", "price": "판매가격", "stock": "잔여수량(재고)",
        "api_loader": load_coupang_products_from_api, "append": "append_to_spreadsheet_coupang",
    },
}


def _find_latest_channel_file(channel):
    config = CHANNELS[channel]["config"]
    if "file_patterns" in config:
        return get_latest_file_from_patterns(config["db_dir"], config["file_patterns"])
    return get_latest_file_from_pattern(config["db_dir"], config["file_pattern"])


def _log_product_lines(log, title, rows, spec, key_column):
    log(title)
    for i, row in enumerate(rows):
        option_id = format_option_id_for_log(row.get(key_column, ""))
        product_name = row.get(spec["name"], "N/A")
        price = format_quantity_for_log(row.get(spec["price"], "N/A"))
        stock = format_quantity_for_log(row.get(spec["stock"], "N/A"))
        log(f"  [{i+1}] {spec['id_label']}: {option_id}, 상품명: {product_name}, 가격: {price}, 재고: {stock}")


def sync_channel(
    channel,
    spreadsheet_id,
    *,
    source=SOURCE_FILE,
    realtime_file_path=None,
    test_mode=True,
    test_count=None,
    log=print,
    verbose_log=True,
):
    """네이버·쿠팡 상품 목록(API 또는 내보내기 파일)과 공유 시트를 비교해 시트에 없는 상품만 행으로 추가한다.

    기존 행은 고치지 않는다. 신규 행의 「상품코드」는 비워 두고 작업자가 실물 위치를 적는다.
    """
    spec = CHANNELS[channel]
    config = spec["config"]
    key_column = config["option_id_column"]
    r = _empty_channel_result(channel)
    r["source"] = source
    if verbose_log:
        log("\n" + "=" * 80)
        log(f"{spec['icon']} {spec['label']} 동기화 시작")
        log("=" * 80)
    else:
        log(f"\n{spec['icon']} {spec['label']} 동기화")

    if source == SOURCE_FILE:
        if realtime_file_path is None:
            realtime_file_path = _find_latest_channel_file(channel)
            if realtime_file_path is None:
                patterns = config.get("file_patterns") or [config["file_pattern"]]
                log(f"⏭️ {spec['label']} 건너뜀: {config['db_dir']} 폴더에 "
                    f"다음 패턴에 맞는 파일이 없습니다: {', '.join(patterns)}")
                log(f"   ({spec['label']}만 건너뜁니다. 다른 선택 채널 동기화는 계속 진행됩니다.)")
                r["skipped_missing_file"] = True
                r["ok"] = True
                if verbose_log:
                    log("\n" + "=" * 80)
                return r
        r["resolved_file"] = str(realtime_file_path)

    log("\n📊 스프레드시트에서 기존 데이터 읽기...")
    gc = get_authorized_gspread_client()
    spreadsheet = gc.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.get_worksheet(config["sheet_index"])
    spreadsheet_headers, spreadsheet_data, header_row = read_sheet_table(worksheet, config["header_row"])
    r["sheet_count"] = len(spreadsheet_data)
    log(f"✅ 기존 DB: {len(spreadsheet_data)}개 제품")

    if source == SOURCE_API:
        realtime_data = spec["api_loader"](_read_shared_config(spreadsheet), spreadsheet_data, log)
    else:
        log(f"📂 실시간 DB 읽기: {realtime_file_path}")
        realtime_data = read_realtime_db(str(realtime_file_path), header_row=spec["file_header_row"])
    r["realtime_count"] = len(realtime_data)
    log(f"✅ 가져온 상품: {len(realtime_data)}개")

    if verbose_log:
        log(f"\n🔍 신제품 검색 중... (키: {key_column})")
        _log_product_lines(log, "\n📋 기존 제품 샘플 (처음 10개):", spreadsheet_data[:10], spec, key_column)
        _log_product_lines(log, "\n📋 가져온 상품 샘플 (처음 10개):", realtime_data[:10], spec, key_column)
    else:
        log("\n🔍 신제품 비교 중...")

    new_products = find_new_products(realtime_data, spreadsheet_data, key_column)
    r["new_products_count"] = len(new_products)
    log(f"\n✅ 신제품 발견: {len(new_products)}개")
    if not new_products:
        log("\n✅ 새 제품이 없습니다. 동기화 완료!")
        r["ok"] = True
        if verbose_log:
            log("\n" + "=" * 80)
        return r

    # 신규 목록은 짧으므로 상세 로그가 꺼져 있어도 보여 준다(작업자가 상품코드를 채울 대상).
    _log_product_lines(log, "\n🆕 신제품 목록:", new_products, spec, key_column)

    products_to_add = new_products
    if test_count is not None and test_count > 0:
        products_to_add = new_products[:test_count]
        log(f"\n🧪 건수 제한: 상위 {test_count}개만 대상으로 합니다.")
    r["planned_append"] = len(products_to_add)

    if test_mode:
        log(f"\n📝 추가 예정 제품 수: {len(products_to_add)}개")
        log("⚠️  미리보기: 스프레드시트에 쓰지 않습니다.")
    else:
        # 가져오는 사이 다른 PC가 같은 상품을 먼저 추가했을 수 있으므로 쓰기 직전에 키를 다시 읽는다.
        _, latest_data, header_row = read_sheet_table(worksheet, config["header_row"])
        fresh = find_new_products(products_to_add, latest_data, key_column)
        r["skipped_concurrent"] = len(products_to_add) - len(fresh)
        if r["skipped_concurrent"]:
            log(f"\nℹ️  그사이 다른 PC가 추가한 {r['skipped_concurrent']}개는 제외합니다.")
        log(f"\n📝 추가할 제품 수: {len(fresh)}개")
        new_rows = [align_to_spreadsheet_headers(product, spreadsheet_headers) for product in fresh]
        append = globals()[spec["append"]]
        append(worksheet, new_rows, header_row, log=log, verbose_log=verbose_log)
        r["appended"] = len(new_rows)

    r["ok"] = True
    if verbose_log:
        log("\n" + "=" * 80)
    return r


def sync_naver(spreadsheet_id, realtime_file_path=None, test_mode=True, test_count=None, log=print, *,
               verbose_log=True, source=SOURCE_FILE):
    """네이버 상품과 스프레드시트 동기화. dict 결과는 UI/CLI 공통 사용."""
    return sync_channel("naver", spreadsheet_id, source=source, realtime_file_path=realtime_file_path,
                        test_mode=test_mode, test_count=test_count, log=log, verbose_log=verbose_log)


def sync_coupang(spreadsheet_id, realtime_file_path=None, test_mode=True, test_count=None, log=print, *,
                 verbose_log=True, source=SOURCE_FILE):
    """쿠팡 상품과 스프레드시트 동기화."""
    return sync_channel("coupang", spreadsheet_id, source=source, realtime_file_path=realtime_file_path,
                        test_mode=test_mode, test_count=test_count, log=log, verbose_log=verbose_log)


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
    source=SOURCE_FILE,
):
    """UI 스레드 밖에서 호출. test_count: 0이면 무제한(None과 동일). verbose_log: 샘플·상세 append 로그.

    source: SOURCE_API면 네이버·쿠팡 API로, SOURCE_FILE이면 database 폴더 파일(naver_path·coupang_path)로 가져온다.
    한 채널이 실패해도 다른 채널은 계속한다.
    """
    logs = []

    def log(msg):
        logs.append(str(msg))

    out = {
        "ok": True,
        "logs": logs,
        "test_mode": test_mode,
        "source": source,
        "naver": None,
        "coupang": None,
        "error": None,
    }
    tc = None if test_count is None or test_count <= 0 else int(test_count)
    where = "API" if source == SOURCE_API else "파일"

    if verbose_log:
        log("—" * 40)
        if test_mode:
            log(f"🧪 미리보기({where}): 스프레드시트에 행을 추가·수정하지 않습니다.")
        else:
            log(f"📝 실반영({where}): 신규 상품이 있으면 시트에 행이 추가됩니다.")
        log("—" * 40)
    else:
        log(f"▶ 미리보기 · 시트 미적용 · {where}" if test_mode else f"▶ 실반영 · 신규 행 시트 추가 가능 · {where}")

    errors = []
    for channel, enabled, path in (("naver", do_naver, naver_path), ("coupang", do_coupang, coupang_path)):
        if not enabled:
            continue
        try:
            out[channel] = sync_channel(
                channel,
                spreadsheet_id,
                source=source,
                realtime_file_path=None if source == SOURCE_API or not path else str(path),
                test_mode=test_mode,
                test_count=tc,
                log=log,
                verbose_log=verbose_log,
            )
        except Exception as e:
            result = _empty_channel_result(channel)
            result["source"] = source
            result["error"] = str(e)
            out[channel] = result
            errors.append(str(e))
            log(f"❌ {CHANNELS[channel]['label']} 실패: {e}")
        if not out[channel].get("ok"):
            out["ok"] = False
    if errors:
        out["error"] = "\n".join(errors)
    _log_db_sync_job_summary(log, do_naver, do_coupang, out, verbose_log=verbose_log)
    return out
