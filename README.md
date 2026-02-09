```
import os
import time
import re
import pythoncom
import win32com.client as win32

# ================= CONFIG =================

MASTER_PATH = r"C:\PATH\Master.xlsx"
OUTPUT_FOLDER = r"C:\PATH\Output_Entity_Files"

ENTITIES = ["APAC", "EMEA", "INDIA", "AMERICAS", "UK"]

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"

# ONLY output sheets (do NOT touch inputs)
OUTPUT_SHEETS = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

ROW_CHUNK = 200   # safe chunk size

# =========================================


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def wait_for_calc(excel):
    while excel.CalculationState != 0:
        time.sleep(1)


def freeze_sheet_in_chunks(ws, excel):
    """
    Freeze formulas to values safely using PasteSpecial
    in small row chunks (NO marshaling issues).
    """
    xlPasteValues = -4163

    used = ws.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count

    for r in range(1, rows + 1, ROW_CHUNK):
        r_end = min(r + ROW_CHUNK - 1, rows)
        rng = ws.Range(
            ws.Cells(r, 1),
            ws.Cells(r_end, cols)
        )
        rng.Copy()
        rng.PasteSpecial(Paste=xlPasteValues)

    excel.CutCopyMode = False


def main():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    pythoncom.CoInitialize()

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    for i, entity in enumerate(ENTITIES, start=1):
        entity_safe = safe_name(entity)
        print(f"\n[{i}/{len(ENTITIES)}] Processing {entity_safe}")

        # 1️⃣ Open master
        wb = excel.Workbooks.Open(MASTER_PATH, UpdateLinks=1)

        # 2️⃣ Set entity
        wb.Worksheets(LANDING_SHEET).Range(ENTITY_CELL).Value = entity

        # 3️⃣ Full recalculation (Excel-native)
        excel.CalculateFullRebuild()
        wait_for_calc(excel)

        # 4️⃣ Save copy
        out_path = os.path.join(OUTPUT_FOLDER, f"{entity_safe}.xlsx")
        if os.path.exists(out_path):
            os.remove(out_path)

        wb.SaveCopyAs(out_path)
        wb.Close(False)

        # 5️⃣ Open copied file
        out_wb = excel.Workbooks.Open(out_path)

        # 6️⃣ Freeze ONLY output sheets
        for sheet_name in OUTPUT_SHEETS:
            ws = out_wb.Worksheets(sheet_name)
            freeze_sheet_in_chunks(ws, excel)

        out_wb.Save()
        out_wb.Close(False)

        print(f"✔ Created value file: {out_path}")

    excel.Quit()
    pythoncom.CoUninitialize()
    print("\n✅ ALL FILES CREATED SUCCESSFULLY")


# ================= ENTRY POINT =================

if __name__ == "__main__":
    main()
