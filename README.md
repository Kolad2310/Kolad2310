```
import os
import time
import re
import pythoncom
import win32com.client as win32


# ================= CONFIG =================

MASTER_PATH = r"C:\PATH\Master.xlsx"
OUTPUT_FOLDER = r"C:\PATH\Output_Entity_Files"

ENTITIES = [
    "APAC",
    "EMEA",
    "INDIA",
    "AMERICAS",
    "UK"
]

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"

# Sheets to convert to values (keep formatting)
OUTPUT_SHEETS = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

ROW_CHUNK = 200  # safe chunk size for PasteSpecial

# =========================================


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def wait_for_refresh_and_calc(excel):
    """
    Wait until:
    - Refresh is done
    - Calculation is done
    """
    while True:
        if excel.CalculationState == 0:
            break
        time.sleep(2)


def freeze_sheet_to_values(ws, excel):
    """
    Convert formulas to values using Excel-native PasteSpecial
    in small chunks (NO marshaling errors)
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

        # 3️⃣ Refresh all (Power Query, connections, etc.)
        wb.RefreshAll()

        # 4️⃣ Wait for refresh + calculation
        wait_for_refresh_and_calc(excel)

        # 5️⃣ Save a copy
        out_path = os.path.join(OUTPUT_FOLDER, f"{entity_safe}.xlsx")
        if os.path.exists(out_path):
            os.remove(out_path)

        wb.SaveCopyAs(out_path)

        # Close master without saving
        wb.Close(False)

        # 6️⃣ Open copied file
        out_wb = excel.Workbooks.Open(out_path)

        # 7️⃣ Convert output sheets to values
        for sheet_name in OUTPUT_SHEETS:
            ws = out_wb.Worksheets(sheet_name)
            freeze_sheet_to_values(ws, excel)

        # 8️⃣ Save & close value version
        out_wb.Save()
        out_wb.Close(False)

        print(f"✔ Value file created: {out_path}")

    excel.Quit()
    pythoncom.CoUninitialize()

    print("\n✅ ALL ENTITY FILES CREATED SUCCESSFULLY")


# ================= ENTRY POINT =================

if __name__ == "__main__":
    main()
