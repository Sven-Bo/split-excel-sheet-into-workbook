from pathlib import Path
import xlwings as xw  # pip install xlwings

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "Files"  # !!! CHANGE THE FILEPATH !!!
OUTPUT_DIR = BASE_DIR / "Output"

# Create Output directory
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

files = list(INPUT_DIR.rglob("*.xls*"))

with xw.App(visible=False) as app:
    for file in files:
        wb = app.books.open(file)
        for sheet in wb.sheets:
            wb_new = app.books.add()
            sheet.copy(after=wb_new.sheets[0])
            wb_new.sheets[0].delete()
            wb_new.save(OUTPUT_DIR / f"{file.stem}_{sheet.name}.xlsx")
            wb_new.close()