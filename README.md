```
import os
import time
import re
import win32com.client as win32


# ============ CONFIG ============
TEMPLATE_PATH = r"C:\FULL\PATH\Template.xlsx"
OUTPUT_FOLDER = r"C:\FULL\PATH\output_value_versions"

ENTITY_SHEET = "Region & Function"
ENTITY_COLUMN = "C"
ENTITY_START_ROW = 2

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"

SHEETS_TO_EXPORT = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]
# ================================


def sanitize(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False

    wb = excel.Workbooks.Open(TEMPLATE_PATH, UpdateLinks=1)

    ws_entity = wb.Sheets(ENTITY_SHEET)
    ws_landing = wb.Sheets(LANDING_SHEET)

    last_row = ws_entity.Cells(
        ws_entity.Rows.Count, ENTITY_COLUMN
    ).End(-4162).Row  # xlUp

    for row in range(ENTITY_START_ROW, last_row + 1):

        entity = ws_entity.Cells(row, ENTITY_COLUMN).Value
        if not entity:
            continue

        entity = sanitize(entity)
        print(f"Processing entity: {entity}")

        # 1️⃣ Set entity
        ws_landing.Range(ENTITY_CELL).Value = entity

        # 2️⃣ Refresh template (SAFE way)
        wb.RefreshAll()
        excel.CalculateFullRebuild()

        # ⛔ Wait until Excel is truly free
        while excel.CalculationState != 0:
            time.sleep(0.5)

        # 3️⃣ Copy sheets (Excel-native, safest)
        wb.Sheets(SHEETS_TO_EXPORT).Copy()
        new_wb = excel.ActiveWorkbook

        # 4️⃣ Convert formulas to values (SAFE)
        for sheet in new_wb.Sheets:
            used = sheet.UsedRange
            used.Value = used.Value

        # 5️⃣ Save
        save_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        if os.path.exists(save_path):
            os.remove(save_path)

        new_wb.SaveAs(save_path, FileFormat=51)
        new_wb.Close(False)

    wb.Close(False)
    excel.Quit()

    print("✅ COMPLETED WITHOUT COM ERRORS")


if __name__ == "__main__":
    main()
