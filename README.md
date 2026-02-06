```
import os
import time
import re
import win32com.client as win32


# ================= CONFIG =================

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

# ==========================================


def sanitize_filename(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def copy_sheet_as_values(src_ws, tgt_wb, sheet_name):
    tgt_ws = tgt_wb.Sheets.Add()
    tgt_ws.Name = sheet_name

    used_range = src_ws.UsedRange
    used_range.Copy()

    tgt_ws.Range("A1").PasteSpecial(Paste=-4163)   # xlPasteValues
    tgt_ws.Range("A1").PasteSpecial(Paste=-4122)   # xlPasteFormats

    # Copy column widths
    for col in range(1, used_range.Columns.Count + 1):
        tgt_ws.Columns(col).ColumnWidth = src_ws.Columns(col).ColumnWidth


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    wb = excel.Workbooks.Open(TEMPLATE_PATH, UpdateLinks=1)

    ws_entity = wb.Sheets(ENTITY_SHEET)
    ws_landing = wb.Sheets(LANDING_SHEET)

    last_row = ws_entity.Cells(
        ws_entity.Rows.Count, ENTITY_COLUMN
    ).End(-4162).Row   # xlUp

    for row in range(ENTITY_START_ROW, last_row + 1):

        entity = ws_entity.Cells(row, ENTITY_COLUMN).Value
        if not entity:
            continue

        entity = sanitize_filename(entity)
        print(f"Processing entity: {entity}")

        # 1️⃣ Set entity
        ws_landing.Range(ENTITY_CELL).Value = entity

        # 2️⃣ Refresh TEMPLATE
        wb.RefreshAll()
        excel.CalculateFull()
        time.sleep(2)   # allow refresh to complete

        # 3️⃣ Create new workbook
        new_wb = excel.Workbooks.Add()

        # Remove default sheets safely
        while new_wb.Sheets.Count > 1:
            new_wb.Sheets(1).Delete()

        new_wb.Sheets(1).Delete()

        # 4️⃣ Copy sheets as VALUES
        for sheet_name in SHEETS_TO_EXPORT:
            copy_sheet_as_values(
                wb.Sheets(sheet_name),
                new_wb,
                sheet_name
            )

        # 5️⃣ Save value version
        save_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        if os.path.exists(save_path):
            os.remove(save_path)

        new_wb.SaveAs(save_path, FileFormat=51)
        new_wb.Close(False)

    wb.Close(False)
    excel.Quit()

    print("✅ All refreshed value-version files created successfully")


if __name__ == "__main__":
    main()
