import gspread

# api-key 폴더의 JSON 파일 경로
credential_path = "api-key/beaming-figure-476816-r5-7dd9d6f34342.json"

try:
    # gspread의 service_account 메서드 사용 (자동으로 필요한 스코프 처리)
    gc = gspread.service_account(filename=credential_path)
    
    print("✅ 인증 성공!")
    
    # 스프레드시트 ID로 직접 접근 (Drive API 없이 접근 가능)
    spreadsheet_id = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"
    sheet = gc.open_by_key(spreadsheet_id)
    print("✅ 스프레드시트 'store_database' 연결 성공!")
    
    # 사용 가능한 워크시트 목록 출력
    worksheets = sheet.worksheets()
    print(f"\n📋 사용 가능한 워크시트 목록 ({len(worksheets)}개):")
    for i, ws in enumerate(worksheets, 1):
        print(f"  {i}. {ws.title}")
    
    # ===== 1번 시트 (네이버 DB) 테스트 =====
    print("\n" + "="*60)
    print("📊 1번 시트 (네이버 DB) 테스트")
    print("="*60)
    
    # 시트 번호로 직접 접근 (인덱스는 0부터 시작하므로 0 = 1번 시트)
    ws_naver = sheet.get_worksheet(0)
    print(f"✅ 1번 시트 '{ws_naver.title}' 선택 성공!")
    
    # 데이터 가져오기
    all_records_naver = ws_naver.get_all_records()
    total_rows_naver = len(all_records_naver)
    
    print(f"✅ 네이버 DB: 총 {total_rows_naver}행 불러옴")
    
    # 첫 몇 행 데이터 확인 (옵션)
    if total_rows_naver > 0:
        print(f"\n📊 네이버 DB 첫 번째 행 데이터:")
        first_record = all_records_naver[0]
        print(f"  상품코드: {first_record.get('상품코드', 'N/A')}")
        print(f"  상품명: {first_record.get('상품명', 'N/A')}")
        print(f"  판매가: {first_record.get('판매가', 'N/A')}")
    
    # ===== 2번 시트 (쿠팡 DB) 테스트 =====
    print("\n" + "="*60)
    print("📊 2번 시트 (쿠팡 DB) 테스트")
    print("="*60)
    
    # 시트 번호로 직접 접근 (인덱스는 0부터 시작하므로 1 = 2번 시트)
    ws_coupang = sheet.get_worksheet(1)
    print(f"✅ 2번 시트 '{ws_coupang.title}' 선택 성공!")
    
    # 쿠팡 시트는 헤더가 2행부터 시작
    # 헤더 행 전체 확인 (컬럼 인덱스와 함께)
    
    def get_column_letter(n):
        """숫자를 Excel 컬럼 문자로 변환 (0 -> A, 1 -> B, ...)"""
        result = ""
        n += 1  # 0-based를 1-based로 변환
        while n > 0:
            n -= 1
            result = chr(65 + (n % 26)) + result
            n //= 26
        return result
    
    # 2행 헤더 전체 가져오기
    header_row = ws_coupang.row_values(2)
    
    print(f"\n📋 쿠팡 시트 헤더 행 전체 분석 (총 {len(header_row)}개 컬럼):")
    print("="*80)
    
    # 헤더명 -> 인덱스 리스트 매핑 (중복 찾기용)
    header_map = {}
    for idx, header in enumerate(header_row):
        header_name = header if header else "(빈값)"
        if header_name not in header_map:
            header_map[header_name] = []
        header_map[header_name].append(idx)
    
    # 각 컬럼 출력 및 중복 표시
    for idx, header in enumerate(header_row):
        col_letter = get_column_letter(idx)
        header_name = header if header else "(빈값)"
        duplicate_marker = ""
        
        # 중복 확인
        if header_name in header_map and len(header_map[header_name]) > 1:
            duplicate_marker = f" ⚠️ 중복! ({len(header_map[header_name])}개 중 {header_map[header_name].index(idx) + 1}번째)"
        
        print(f"  {col_letter:4s} | {idx:3d}번 컬럼 | {header_name}{duplicate_marker}")
    
    # 중복된 헤더 요약
    print("\n" + "="*80)
    print("🔍 중복된 헤더 요약:")
    duplicates_found = False
    skipped_headers = {}  # 무시된 헤더 추적
    for header_name, indices in header_map.items():
        if len(indices) > 1:
            duplicates_found = True
            col_letters = [get_column_letter(idx) for idx in indices]
            print(f"  '{header_name}': {len(indices)}개 중복 -> {', '.join(col_letters)}")
            print(f"    → 첫 번째 컬럼({col_letters[0]})만 사용, 나머지 {len(indices)-1}개 무시")
            # 무시될 컬럼들 저장
            for idx in indices[1:]:  # 첫 번째 이후 모든 인덱스
                skipped_headers[idx] = header_name
    
    if not duplicates_found:
        print("  ✅ 중복된 헤더가 없습니다.")
    
    print("\n" + "="*80)
    
    # ===== 중복 헤더 무시하고 데이터 가져오기 =====
    print("\n📊 중복 헤더 무시하고 데이터 로딩 테스트:")
    print("="*80)
    
    # 모든 데이터 가져오기
    all_values = ws_coupang.get_all_values()
    
    if len(all_values) < 2:
        print("⚠️  데이터가 없습니다.")
    else:
        # 헤더 처리: 중복된 경우 첫 번째 것만 사용
        unique_headers = []
        used_headers = set()  # 이미 사용된 헤더명 추적
        column_indices_to_keep = []  # 유지할 컬럼 인덱스
        
        for idx, header in enumerate(header_row):
            header_name = header if header else f"빈컬럼_{idx}"
            
            # 빈 헤더는 "빈컬럼_인덱스"로 처리
            if not header or header.strip() == "":
                if f"빈컬럼_{idx}" not in used_headers:
                    unique_headers.append(f"빈컬럼_{idx}")
                    used_headers.add(f"빈컬럼_{idx}")
                    column_indices_to_keep.append(idx)
            # 중복 체크: 첫 번째 것만 사용
            elif header_name not in used_headers:
                unique_headers.append(header_name)
                used_headers.add(header_name)
                column_indices_to_keep.append(idx)
            # 중복된 경우 무시
            else:
                print(f"  ⏭️  컬럼 {get_column_letter(idx)} ({idx}번) '{header_name}' 무시됨 (중복)")
        
        print(f"\n✅ 유효한 헤더: {len(unique_headers)}개")
        print(f"⏭️  무시된 컬럼: {len(header_row) - len(unique_headers)}개")
        
        # 데이터 행을 딕셔너리로 변환 (헤더 행 제외: 인덱스 2부터)
        data_records = []
        for row_idx in range(2, len(all_values)):  # 3행부터 (인덱스 2)
            row_data = all_values[row_idx]
            record = {}
            for i, header_name in enumerate(unique_headers):
                col_idx = column_indices_to_keep[i]
                if col_idx < len(row_data):
                    record[header_name] = row_data[col_idx]
                else:
                    record[header_name] = ""
            data_records.append(record)
        
        total_rows_coupang = len(data_records)
        print(f"✅ 쿠팡 DB: 총 {total_rows_coupang}행 불러옴")
        
        # 첫 번째 행 데이터 확인
        if total_rows_coupang > 0:
            print(f"\n📊 쿠팡 DB 첫 번째 행 데이터 (처음 10개 필드):")
            first_record = data_records[0]
            for i, (key, value) in enumerate(list(first_record.items())[:10]):
                print(f"  {key}: {value}")
    
except FileNotFoundError:
    print(f"❌ 오류: 인증 파일을 찾을 수 없습니다: {credential_path}")
except gspread.exceptions.SpreadsheetNotFound:
    print("❌ 오류: 'store_database' 스프레드시트를 찾을 수 없습니다.")
except IndexError:
    print("❌ 오류: 해당 번호의 시트를 찾을 수 없습니다. (시트가 존재하지 않음)")
except Exception as e:
    print(f"❌ 오류 발생: {type(e).__name__}: {str(e)}")

