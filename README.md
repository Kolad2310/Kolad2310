```
import os
import shutil
import time
import re
import win32com.client as win32


# ========== CONFIG ==========
MASTER_PATH = r"C:\FULL\PATH\Master.xlsx"
OUTPUT_FOLDER = r"C:\FULL\PATH\output_files"

ENTITY_SHEET = "Region & Function"
ENTITY_COLUMN = "C"
ENTITY_START_ROW = 2

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"

SHEETS_TO_KEEP = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]
# ============================


def safe_name(name):
    """Make entity safe for Windows filename"""
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def wait_excel(excel):
    """Wait until Excel finishes refresh/calculation"""
    while excel.CalculationState != 0:
        time.sleep(0.5)


def get_entities():
    """Read entity list once from master"""
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(MASTER_PATH, ReadOnly=True)
    ws = wb.Sheets(ENTITY_SHEET)

    last_row = ws.Cells(ws.Rows.Count, ENTITY_COLUMN).End(-4162).Row
    entities = []

    for r in range(ENTITY_START_ROW, last_row + 1):
        val = ws.Cells(r, ENTITY_COLUMN).Value
        if val:
            entities.append(safe_name(val))

    wb.Close(False)
    excel.Quit()
    return entities


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    entities = get_entities()

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    for entity in entities:
        print(f"\n▶ Processing entity: {entity}")

        # 1️⃣ Copy + rename master immediately
        output_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        if os.path.exists(output_path):
            os.remove(output_path)

        shutil.copy2(MASTER_PATH, output_path)

        # 2️⃣ Open copied (already renamed) file
        wb = excel.Workbooks.Open(output_path, UpdateLinks=1)

        # 3️⃣ Set entity in landing page
        wb.Sheets(LANDING_SHEET).Range(ENTITY_CELL).Value = entity

        # 4️⃣ Refresh
        wb.RefreshAll()
        excel.CalculateFull()
        wait_excel(excel)

        # 5️⃣ Delete unwanted sheets
        for ws in list(wb.Sheets):
            if ws.Name not in SHEETS_TO_KEEP:
                ws.Delete()

        # 6️⃣ Save & close
        wb.Save()
        wb.Close(False)

        print(f"   Saved → {output_path}")

    excel.Quit()
    print("\n✅ All entity files created, renamed, refreshed, and trimmed successfully")


if __name__ == "__main__":
    main()
