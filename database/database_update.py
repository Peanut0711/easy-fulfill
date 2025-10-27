import pandas as pd
import os
import glob
import re
from datetime import datetime
import logging
import sys
import io
import shutil

# 터미널 인코딩 설정 (Windows 한글 깨짐 해결)
if sys.platform == 'win32':
    try:
        # Windows에서 UTF-8 출력 강제
        import codecs
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
    except:
        pass

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 로깅 핸들러에도 UTF-8 인코딩 설정
for handler in logger.handlers:
    if hasattr(handler, 'stream') and hasattr(handler.stream, 'buffer'):
        handler.stream = io.TextIOWrapper(handler.stream.buffer, encoding='utf-8')

class DatabaseUpdater:
    def __init__(self):
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.db_file = os.path.join(self.current_dir, 'store_database.xlsx')
        self.output_file = os.path.join(self.current_dir, 'store_database_updated.xlsx')
        
        # 스토어별 설정
        self.stores = {
            'naver': {
                'pattern': 'Product_YYYYMMDD_HHMMSS.csv',
                'sheet_name': '네이버 스토어 DB',
                'key_column': '상품번호(스마트스토어)',
                'channel': '네이버'
            },
            'coupang': {
                'pattern': 'price_inventory_YYYYMMDD.xlsx',
                'sheet_name': '쿠팡 스토어 DB',
                'key_column': 'Product ID',
                'channel': '쿠팡'
            }
        }
    
    def find_latest_store_file(self, store_type):
        """최신 날짜의 스토어 파일을 찾는 함수 (파일명의 날짜를 기준으로 정렬)"""
        try:
            if store_type == 'naver':
                # Product_로 시작하고 .csv로 끝나는 파일 찾기
                pattern = os.path.join(self.current_dir, 'Product_*.csv')
                files = glob.glob(pattern)
                if not files:
                    return None
                
                # 파일명에서 날짜 추출 (Product_20251027_104406.csv -> 20251027)
                file_dates = []
                for file in files:
                    filename = os.path.basename(file)
                    # Product_YYYYMMDD_HHMMSS.csv 패턴 매칭
                    match = re.search(r'Product_(\d{8})_(\d{6})\.csv', filename)
                    if match:
                        date_str = match.group(1)  # YYYYMMDD
                        time_str = match.group(2)  # HHMMSS
                        datetime_str = f"{date_str}{time_str}"  # YYYYMMDDHHMMSS
                        file_dates.append((file, datetime_str))
                
                if file_dates:
                    # 날짜+시간을 기준으로 정렬하여 최신 파일 반환
                    file_dates.sort(key=lambda x: x[1], reverse=True)
                    return file_dates[0][0]
                
                # 날짜 추출 실패 시 기존 방식 사용 (수정 시간 기준)
                files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return files[0]
            
            elif store_type == 'coupang':
                # price_inventory_로 시작하고 .xlsx로 끝나는 파일 찾기
                pattern = os.path.join(self.current_dir, 'price_inventory_*.xlsx')
                files = glob.glob(pattern)
                if not files:
                    return None
                
                # 파일명에서 날짜 추출 (price_inventory_250925.xlsx -> 250925)
                file_dates = []
                for file in files:
                    filename = os.path.basename(file)
                    # price_inventory_YYMMDD.xlsx 패턴 매칭
                    match = re.search(r'price_inventory_(\d{6})\.xlsx', filename)
                    if match:
                        date_str = match.group(1)  # YYMMDD
                        # 20YYMMDD 형식으로 변환
                        full_date = f"20{date_str}"  # 250925 -> 20250925
                        file_dates.append((file, full_date))
                
                if file_dates:
                    # 날짜를 기준으로 정렬하여 최신 파일 반환
                    file_dates.sort(key=lambda x: x[1], reverse=True)
                    return file_dates[0][0]
                
                # 날짜 추출 실패 시 기존 방식 사용 (수정 시간 기준)
                files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return files[0]
            
            return None
            
        except Exception as e:
            logger.error(f"파일 탐색 중 오류 발생 ({store_type}): {str(e)}")
            return None
    
    def read_store_data(self, file_path, store_type):
        """스토어 데이터를 읽는 함수"""
        try:
            if store_type == 'naver':
                # CSV 파일 읽기 (UTF-8 인코딩)
                df = pd.read_csv(file_path, encoding='utf-8')
            elif store_type == 'coupang':
                # 엑셀 파일 읽기 (3번째 행을 헤더로 사용)
                df = pd.read_excel(file_path, header=2)
            else:
                logger.error(f"지원하지 않는 스토어 타입: {store_type}")
                return None
            
            logger.info(f"{store_type} 스토어 데이터 읽기 완료: {len(df)}개 상품")
            return df
            
        except UnicodeDecodeError:
            # UTF-8로 읽기 실패 시 다른 인코딩 시도
            try:
                if store_type == 'naver':
                    df = pd.read_csv(file_path, encoding='cp949')
                    logger.info(f"{store_type} 스토어 데이터 읽기 완료 (cp949): {len(df)}개 상품")
                    return df
            except Exception as e:
                logger.error(f"CSV 파일 인코딩 오류 ({store_type}): {str(e)}")
                return None
        except Exception as e:
            logger.error(f"스토어 데이터 읽기 오류 ({store_type}): {str(e)}")
            return None
    
    def read_database(self, expected_sheet_name):
        """데이터베이스 시트를 읽는 함수 (시트 이름 자동 탐지)"""
        try:
            # 먼저 설정된 시트 이름으로 시도
            try:
                df = pd.read_excel(self.db_file, sheet_name=expected_sheet_name)
                logger.info(f"데이터베이스 시트 읽기 완료 ({expected_sheet_name}): {len(df)}개 상품")
                return df
            except Exception:
                # 시트 이름을 찾을 수 없으면 첫 번째 시트 사용
                df = pd.read_excel(self.db_file, sheet_name=0)
                logger.info(f"데이터베이스 시트 읽기 완료 (첫 번째 시트): {len(df)}개 상품")
                return df
            
        except Exception as e:
            logger.error(f"데이터베이스 읽기 오류 ({expected_sheet_name}): {str(e)}")
            return None
    
    def compare_and_update(self, store_df, db_df, store_type):
        """스토어 데이터와 DB를 비교하여 업데이트하는 함수"""
        try:
            store_config = self.stores[store_type]
            key_column = store_config['key_column']
            
            # 키 컬럼이 존재하는지 확인
            if key_column not in store_df.columns:
                logger.error(f"스토어 파일에 키 컬럼이 없습니다: {key_column}")
                return None, [], []
            
            if key_column not in db_df.columns:
                logger.error(f"DB 파일에 키 컬럼이 없습니다: {key_column}")
                return None, [], []
            
            # 기존 상품 번호 목록
            existing_products = set(db_df[key_column].dropna().astype(str))
            store_products = set(store_df[key_column].dropna().astype(str))
            
            # 신규 상품 찾기
            new_products = store_products - existing_products
            new_products_df = store_df[store_df[key_column].astype(str).isin(new_products)].copy()
            
            # 서버에는 있지만 DB에 없는 필드명 찾기
            db_columns = set(db_df.columns)
            store_columns = set(store_df.columns)
            missing_fields = store_columns - db_columns
            
            # 신규 상품 정보 출력용 데이터
            new_product_info = []
            if len(new_products_df) > 0:
                for _, row in new_products_df.iterrows():
                    product_info = {
                        '상품번호': row.get(key_column, ''),
                        '상품명': row.get('상품명', ''),
                        '채널': store_config['channel']
                    }
                    new_product_info.append(product_info)
            
            # DB 컬럼 순서에 맞춰 신규 상품 데이터 정렬
            if len(new_products_df) > 0:
                # DB의 컬럼 순서를 기준으로 정렬
                ordered_columns = list(db_df.columns)
                new_products_ordered = pd.DataFrame()
                
                for col in ordered_columns:
                    if col in new_products_df.columns:
                        new_products_ordered[col] = new_products_df[col]
                    else:
                        # DB에는 있지만 스토어 데이터에 없는 컬럼은 빈 값으로 채움
                        new_products_ordered[col] = ''
                
                # 상품코드 컬럼은 빈 값으로 설정
                if '상품코드' in new_products_ordered.columns:
                    new_products_ordered['상품코드'] = ''
                
                # 기존 DB에 신규 상품 추가
                updated_db = pd.concat([db_df, new_products_ordered], ignore_index=True)
            else:
                updated_db = db_df.copy()
            
            return updated_db, list(missing_fields), new_product_info
            
        except Exception as e:
            logger.error(f"데이터 비교 및 업데이트 오류 ({store_type}): {str(e)}")
            return None, [], []
    
    def find_store_sheet(self, store_type):
        """스토어별 시트를 찾는 함수"""
        try:
            xl_file = pd.ExcelFile(self.db_file)
            sheet_names = xl_file.sheet_names
            
            if store_type == 'naver':
                # product로 시작하는 시트 찾기
                for sheet_name in sheet_names:
                    if sheet_name.lower().startswith('product'):
                        return sheet_name
            elif store_type == 'coupang':
                # price_inventory로 시작하는 시트 찾기
                for sheet_name in sheet_names:
                    if sheet_name.lower().startswith('price_inventory'):
                        return sheet_name
            
            # 못 찾으면 첫 번째 시트 반환
            return sheet_names[0] if sheet_names else None
            
        except Exception as e:
            logger.error(f"시트 찾기 오류 ({store_type}): {str(e)}")
            return None
    
    def backup_existing_database(self):
        """기존 데이터베이스 파일을 백업하는 함수"""
        try:
            if not os.path.exists(self.db_file):
                return True  # 파일이 없으면 백업할 필요 없음
            
            # 기존 백업 파일들 찾기
            backup_pattern = os.path.join(self.current_dir, 'store_database_old*.xlsx')
            existing_backups = glob.glob(backup_pattern)
            
            # 다음 백업 번호 결정
            if existing_backups:
                # 기존 백업 파일에서 번호 추출
                backup_numbers = []
                for backup_file in existing_backups:
                    filename = os.path.basename(backup_file)
                    match = re.search(r'store_database_old\((\d+)\)\.xlsx', filename)
                    if match:
                        backup_numbers.append(int(match.group(1)))
                
                if backup_numbers:
                    next_number = max(backup_numbers) + 1
                else:
                    next_number = 1
            else:
                next_number = 1
            
            # 백업 파일명 생성
            backup_filename = f'store_database_old({next_number}).xlsx'
            backup_path = os.path.join(self.current_dir, backup_filename)
            
            # 파일 복사
            shutil.copy2(self.db_file, backup_path)
            logger.info(f"기존 데이터베이스 백업 완료: {backup_filename}")
            print(f"[백업] 기존 데이터베이스가 백업되었습니다: {backup_filename}")
            
            return True
            
        except Exception as e:
            logger.error(f"데이터베이스 백업 오류: {str(e)}")
            print(f"[오류] 데이터베이스 백업 실패: {str(e)}")
            return False
    
    def read_database_with_header(self, sheet_name, store_type):
        """스토어별로 다른 헤더 설정으로 DB를 읽는 함수"""
        try:
            if store_type == 'coupang':
                # 쿠팡은 2번째 행(인덱스 1)을 헤더로 사용
                df = pd.read_excel(self.db_file, sheet_name=sheet_name, header=1)
            else:
                # 네이버는 첫 번째 행을 헤더로 사용
                df = pd.read_excel(self.db_file, sheet_name=sheet_name)
            
            logger.info(f"데이터베이스 시트 읽기 완료 ({sheet_name}): {len(df)}개 상품")
            return df
            
        except Exception as e:
            logger.error(f"데이터베이스 읽기 오류 ({sheet_name}): {str(e)}")
            return None
    
    def update_store_database(self, store_type):
        """특정 스토어의 데이터베이스를 업데이트하는 함수"""
        try:
            store_config = self.stores[store_type]
            
            # 최신 스토어 파일 찾기
            store_file = self.find_latest_store_file(store_type)
            if not store_file:
                logger.warning(f"{store_type} 스토어 파일을 찾을 수 없습니다.")
                return 0, [], [], None
            
            logger.info(f"{store_type} 스토어 최신 파일: {os.path.basename(store_file)}")
            
            # 스토어 데이터 읽기
            store_df = self.read_store_data(store_file, store_type)
            if store_df is None:
                return 0, [], [], None
            
            # 스토어별 시트 찾기
            sheet_name = self.find_store_sheet(store_type)
            if not sheet_name:
                logger.error(f"{store_type} 스토어 시트를 찾을 수 없습니다.")
                return 0, [], [], None
            
            # DB 데이터 읽기 (스토어별 헤더 설정 적용)
            db_df = self.read_database_with_header(sheet_name, store_type)
            if db_df is None:
                return 0, [], [], None
            
            # 데이터 비교 및 업데이트
            updated_db, missing_fields, new_product_info = self.compare_and_update(
                store_df, db_df, store_type
            )
            
            if updated_db is None:
                return 0, [], [], None
            
            return len(new_product_info), missing_fields, new_product_info, updated_db
            
        except Exception as e:
            logger.error(f"스토어 데이터베이스 업데이트 오류 ({store_type}): {str(e)}")
            return 0, [], [], None
    
    def save_updated_database(self, naver_db, coupang_db):
        """업데이트된 데이터베이스를 저장하는 함수"""
        try:
            # 기존 데이터베이스 백업
            if not self.backup_existing_database():
                return False
            
            with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                if naver_db is not None:
                    naver_db.to_excel(writer, sheet_name='네이버 스토어 DB', index=False)
                if coupang_db is not None:
                    # 쿠팡 DB는 첫 번째 행을 공란으로 두고 저장
                    coupang_db.to_excel(writer, sheet_name='쿠팡 스토어 DB', index=False, startrow=1)
            
            # 업데이트된 파일을 원래 이름으로 변경
            if os.path.exists(self.output_file):
                shutil.move(self.output_file, self.db_file)
                logger.info(f"업데이트된 데이터베이스 저장 완료: {self.db_file}")
                print(f"[성공] 데이터베이스가 업데이트되었습니다: store_database.xlsx")
                return True
            else:
                logger.error("업데이트된 파일이 생성되지 않았습니다.")
                return False
            
        except Exception as e:
            logger.error(f"데이터베이스 저장 오류: {str(e)}")
            return False
    
    def run_update(self):
        """메인 업데이트 실행 함수"""
        try:
            logger.info("데이터베이스 업데이트 시작")
            
            # DB 파일 존재 확인
            if not os.path.exists(self.db_file):
                logger.error(f"데이터베이스 파일을 찾을 수 없습니다: {self.db_file}")
                print(f"[오류] 데이터베이스 파일을 찾을 수 없습니다: {self.db_file}")
                return
            
            # 네이버 스토어 업데이트
            naver_count, naver_missing_fields, naver_new_products, naver_updated_db = self.update_store_database('naver')
            
            # 쿠팡 스토어 업데이트
            coupang_count, coupang_missing_fields, coupang_new_products, coupang_updated_db = self.update_store_database('coupang')
            
            # 결과 출력
            print("\n" + "="*50)
            print("[데이터베이스 업데이트 결과]")
            print("="*50)
            
            # 신규 상품 추가 개수
            if naver_count > 0:
                print(f"[성공] 네이버 스토어 신규 상품 추가: {naver_count}개")
            if coupang_count > 0:
                print(f"[성공] 쿠팡 스토어 신규 상품 추가: {coupang_count}개")
            
            if naver_count == 0 and coupang_count == 0:
                print("[정보] 신규 상품 없음")
            
            # 서버에는 있지만 DB에 없는 필드명
            all_missing_fields = naver_missing_fields + coupang_missing_fields
            if all_missing_fields:
                print(f"\n[경고] 서버에는 있지만 내 DB에 없는 필드명:")
                for field in all_missing_fields:
                    print(f"   - {field}")
            
            # 재고 담당자 상품코드 추가 대상
            all_new_products = naver_new_products + coupang_new_products
            if all_new_products:
                print(f"\n[알림] 재고 담당자 상품코드 추가 대상 ({len(all_new_products)}개):")
                for product in all_new_products:
                    print(f"   - 상품번호: {product['상품번호']}, 상품명: {product['상품명']}, 채널: {product['채널']}")
            
            print("\n" + "="*50)
            
            # 업데이트된 데이터베이스 저장
            if naver_count > 0 or coupang_count > 0:
                # 업데이트된 DB가 있는 경우에만 저장
                if self.save_updated_database(naver_updated_db, coupang_updated_db):
                    print("[성공] 업데이트된 데이터베이스가 저장되었습니다.")
                else:
                    print("[오류] 데이터베이스 저장에 실패했습니다.")
            else:
                print("[정보] 변경사항이 없어 저장하지 않았습니다.")
            
            logger.info("데이터베이스 업데이트 완료")
            
        except Exception as e:
            logger.error(f"메인 업데이트 실행 오류: {str(e)}")
            print(f"❗ 오류 발생: {str(e)}")

def main():
    """메인 실행 함수"""
    updater = DatabaseUpdater()
    updater.run_update()

if __name__ == "__main__":
    main()
