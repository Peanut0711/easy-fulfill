"""월별 사은품 기록을 엑셀 한 장으로 내보낸다."""

from pathlib import Path
from tempfile import NamedTemporaryFile

from orders.gift_event import GIFT_EVENT_STORES
from orders.gift_history import MONTHLY_TIER_COUNT, gift_tier_quantities, summarize_gift_month


def export_gift_month(path: Path, records, year: int, month: int, store: str = "") -> None:
    from xlsxwriter import Workbook
    from xlsxwriter.utility import xl_col_to_name

    totals, daily = summarize_gift_month(records, year, month, store)
    count = MONTHLY_TIER_COUNT
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as handle:
        temporary_path = Path(handle.name)
    try:
        with Workbook(temporary_path) as workbook:
            sheet = workbook.add_worksheet("월별 사은품")
            sheet.hide_gridlines(2)
            sheet.set_column("A:A", 15)
            sheet.set_column("B:F", 13)
            sheet.set_row(0, 28)
            sheet.set_row(1, 21)
            sheet.set_row(2, 21)
            sheet.set_row(3, 21)
            sheet.freeze_panes(4, 0)

            base = {"font_name": "Malgun Gothic", "valign": "vcenter"}
            title = workbook.add_format({**base, "bold": True, "font_size": 15, "font_color": "#16324F"})
            subtitle = workbook.add_format({**base, "italic": True, "font_color": "#46566B"})
            header = workbook.add_format({**base, "bold": True, "bg_color": "#16324F", "font_color": "#FFFFFF",
                                          "align": "center", "right": 1, "right_color": "#FFFFFF"})
            date_format = workbook.add_format({**base, "num_format": "yyyy-mm-dd"})
            number = workbook.add_format({**base, "num_format": "#,##0"})
            summary = workbook.add_format({**base, "bold": True, "bg_color": "#DDEAF7", "num_format": "#,##0"})

            store_label = GIFT_EVENT_STORES.get(store, "전체 스토어")
            sheet.write_string("A1", f"{year}년 {month}월 사은품 집계", title)
            sheet.write_string("A2", f"{store_label} · 한국시간 · 공유 시트의 마크다운 기록 기준", subtitle)
            sheet.write_row("A3", ["날짜", *(f"사은품 {tier}" for tier in range(1, count + 1)), "합계"], header)
            sheet.write_string("A4", "월 합계", summary)
            monthly = gift_tier_quantities(totals, count)
            for column, quantity in enumerate(monthly, start=1):
                letter = xl_col_to_name(column)
                sheet.write_formula(3, column, f"=SUM({letter}5:{letter}{len(daily) + 4})", summary, quantity)
            sheet.write_formula("F4", "=SUM(B4:E4)", summary, sum(monthly))

            for row, (day, counts) in enumerate(daily, start=4):
                quantities = gift_tier_quantities(counts, count)
                sheet.write_datetime(row, 0, day, date_format)
                sheet.write_row(row, 1, quantities, number)
                sheet.write_formula(row, 5, f"=SUM(B{row + 1}:E{row + 1})", number, sum(quantities))
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)
