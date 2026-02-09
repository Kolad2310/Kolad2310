```
import os
import time
import re
import win32com.client as win32
import pythoncom


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

# =========================================


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def wait_for_calc(excel):
    while excel.CalculationState != 0:
        time.sleep(1)


def convert_workbook_to_values(wb):
    """Freeze entire workbook to values (format preserved)"""
    for ws in wb.Worksheets:
        used = ws.UsedRange
        used.Value = used.Value


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    pythoncom.CoInitialize()

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    for idx, entity in enumerate(ENTITIES, start=1):
        entity_safe = safe_name(entity)
        print(f"\n[{idx}/{len(ENTITIES)}] Processing {entity_safe}")

        wb = excel.Workbooks.Open(MASTER_PATH, UpdateLinks=1)

        # 1️⃣ Set entity
        wb.Worksheets(LANDING_SHEET).Range(ENTITY_CELL).Value = entity

        # 2️⃣ Full Excel-native recalculation
        excel.CalculateFullRebuild()
        wait_for_calc(excel)

        # 3️⃣ Save copy
        out_path = os.path.join(OUTPUT_FOLDER, f"{entity_safe}.xlsx")
        if os.path.exists(out_path):
            os.remove(out_path)

        wb.SaveCopyAs(out_path)

        # 4️⃣ Open copied file and freeze values
        out_wb = excel.Workbooks.Open(out_path)
        convert_workbook_to_values(out_wb)
        out_wb.Save()
        out_wb.Close(False)

        # 5️⃣ Close master WITHOUT saving
        wb.Close(False)

        print(f"✔ Saved value version → {out_path}")

    excel.Quit()
    pythoncom.CoUninitialize()

    print("\n✅ ALL ENTITY FILES CREATED SUCCESSFULLY")


# ================= ENTRY POINT =================

if __name__ == "__main__":
    main()
