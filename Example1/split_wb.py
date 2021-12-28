from pathlib import Path
import xlwings as xw  # pip install xlwings

BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "data.xlsx"  # !!! CHANGE THE FILE NAME !!!
OUTPUT_DIR = BASE_DIR / "Output"

# Create Output directory
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

with xw.App(visible=False) as app:
    wb = app.books.open(EXCEL_FILE)
    for sheet in wb.sheets:
        wb_new = app.books.add()
        sheet.copy(after=wb_new.sheets[0])
        wb_new.sheets[0].delete()
        wb_new.save(OUTPUT_DIR / f"{sheet.name}.xlsx")
        wb_new.close()