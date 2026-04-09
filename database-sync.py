"""
CLI: database 폴더 최신 파일 → 구글 스프레드시트(신규 행 append).

로직은 db_sheet_sync 모듈을 사용합니다.
"""
import io
import sys

from db_sheet_sync import DEFAULT_SPREADSHEET_ID, sync_coupang, sync_naver

if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
    except Exception:
        pass

if __name__ == "__main__":
    sid = DEFAULT_SPREADSHEET_ID
    print("=" * 80)
    print("🛒 상품 DB 동기화 (CLI)")
    print("=" * 80)
    print("\테스트 모드로 하려면 아래 sync_* 호출에 test_mode=True 로 바꾸세요.\n")

    sync_naver(sid, test_mode=False, test_count=None)
    sync_coupang(sid, test_mode=False, test_count=None)

    print("\n" + "=" * 80)
    print("완료")
