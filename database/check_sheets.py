import pandas as pd

# 엑셀 파일의 시트명 확인
excel_file = pd.ExcelFile('store_database.xlsx')
print("시트명 목록:")
for i, sheet_name in enumerate(excel_file.sheet_names):
    print(f"{i+1}. {sheet_name}")

# 각 시트의 컬럼명도 확인
for sheet_name in excel_file.sheet_names:
    df = pd.read_excel('store_database.xlsx', sheet_name=sheet_name)
    print(f"\n[{sheet_name}] 컬럼명:")
    for col in df.columns:
        print(f"  - {col}")
