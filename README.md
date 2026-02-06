```
import os
import time
import re
import win32com.client as win32


# ============== CONFIG =================

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

# =======================================


def sanitize(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def wait_until_excel_free(excel):
    """Hard wait until Excel finishes refresh/calculation"""
    while excel.CalculationState != 0:
        time.sleep(0.5)


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

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
        print(f"\n▶ Processing entity: {entity}")

        # 1️⃣ Set entity
        ws_landing.Range(ENTITY_CELL).Value = entity

        # 2️⃣ Refresh template
        wb.RefreshAll()
        excel.CalculateFullRebuild()
        wait_until_excel_free(excel)

        # 3️⃣ Create output workbook
        out_wb = excel.Workbooks.Add()

        # Keep exactly one sheet initially
        while out_wb.Sheets.Count > 1:
            out_wb.Sheets(1).Delete()

        # 4️⃣ Copy sheets ONE BY ONE
        for sheet_name in SHEETS_TO_EXPORT:
            print(f"   Copying sheet: {sheet_name}")

            src_ws = wb.Sheets(sheet_name)

            src_ws.Copy(After=out_wb.Sheets(out_wb.Sheets.Count))
            tgt_ws = out_wb.Sheets(out_wb.Sheets.Count)

            # Convert to values immediately
            used = tgt_ws.UsedRange
            used.Value = used.Value

        # Remove initial blank sheet
        out_wb.Sheets(1).Delete()

        # 5️⃣ Save
        save_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        if os.path.exists(save_path):
            os.remove(save_path)

        print(f"   Saving → {save_path}")
        out_wb.SaveAs(save_path, FileFormat=51)
        out_wb.Close(False)

    wb.Close(False)
    excel.Quit()

    print("\n✅ ALL FILES CREATED SEQUENTIALLY WITHOUT COM ERRORS")


if __name__ == "__main__":
    main()
