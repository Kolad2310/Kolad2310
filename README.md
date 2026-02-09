```
import os
import time
import re
import win32com.client as win32
import pythoncom


# ========== CONFIG ==========
MASTER_PATH = r"C:\FULL\PATH\Master.xlsx"
OUTPUT_FOLDER = r"C:\FULL\PATH\output_files"

ENTITY_SHEET = "Region & Function"
ENTITY_COLUMN = "C"
ENTITY_START_ROW = 2

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"
# ============================


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def get_entities():
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


def process_entity(entity):
    pythoncom.CoInitialize()

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    wb = excel.Workbooks.Open(MASTER_PATH, UpdateLinks=1)

    # Set entity
    wb.Sheets(LANDING_SHEET).Range(ENTITY_CELL).Value = entity

    # Refresh + full rebuild (SAFE)
    wb.RefreshAll()
    excel.CalculateFullRebuild()

    while excel.CalculationState != 0:
        time.sleep(0.5)

    # Convert entire workbook to values
    for ws in wb.Worksheets:
        ws.Cells.Value = ws.Cells.Value

    out_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
    if os.path.exists(out_path):
        os.remove(out_path)

    wb.SaveAs(out_path, FileFormat=51)

    wb.Close(False)
    excel.Quit()
    pythoncom.CoUninitialize()


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    entities = get_entities()
    print(f"Found {len(entities)} entities")

    for i, entity in enumerate(entities, start=1):
        print(f"[{i}/{len(entities)}] Processing {entity}")
        process_entity(entity)

    print("✅ All files created successfully")


if __name__ == "__main__":
    main()
