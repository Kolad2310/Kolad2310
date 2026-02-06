```
import os
import time
import re
import win32com.client as win32


# ========== CONFIG ==========

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
# ============================


def sanitize(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def wait_until_done(excel):
    while excel.CalculationState != 0:
        time.sleep(0.5)


def get_entity_list(excel):
    """Read entity list ONCE (safe)"""
    wb = excel.Workbooks.Open(MASTER_PATH, ReadOnly=True)
    ws = wb.Sheets(ENTITY_SHEET)

    last_row = ws.Cells(ws.Rows.Count, ENTITY_COLUMN).End(-4162).Row
    entities = []

    for r in range(ENTITY_START_ROW, last_row + 1):
        val = ws.Cells(r, ENTITY_COLUMN).Value
        if val:
            entities.append(sanitize(val))

    wb.Close(False)
    return entities


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    # 🔒 Read entities ONCE
    entities = get_entity_list(excel)

    for entity in entities:
        print(f"\n▶ Processing {entity}")

        # 1️⃣ Open master fresh
        wb = excel.Workbooks.Open(MASTER_PATH, UpdateLinks=1)

        ws_landing = wb.Sheets(LANDING_SHEET)

        # 2️⃣ Set entity
        ws_landing.Range(ENTITY_CELL).Value = entity

        # 3️⃣ Refresh
        wb.RefreshAll()
        excel.CalculateFullRebuild()
        wait_until_done(excel)

        # 4️⃣ Create output workbook
        out_wb = excel.Workbooks.Add()

        # Keep one sheet
        while out_wb.Sheets.Count > 1:
            out_wb.Sheets(1).Delete()

        # 5️⃣ Copy sheets ONE BY ONE
        for sheet_name in SHEETS_TO_EXPORT:
            wb.Sheets(sheet_name).Copy(
                After=out_wb.Sheets(out_wb.Sheets.Count)
            )
            tgt = out_wb.Sheets(out_wb.Sheets.Count)
            used = tgt.UsedRange
            used.Value = used.Value

        # Remove initial blank
        out_wb.Sheets(1).Delete()

        # 6️⃣ Save
        save_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        if os.path.exists(save_path):
            os.remove(save_path)

        out_wb.SaveAs(save_path, FileFormat=51)
        out_wb.Close(False)

        # 7️⃣ Close master COMPLETELY
        wb.Close(False)

    excel.Quit()
    print("\n✅ ALL FILES CREATED – NO COM ERRORS")


if __name__ == "__main__":
    main()
