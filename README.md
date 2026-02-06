```
import os
import time
import re
import win32com.client as win32


# ================= USER CONFIG =================

TEMPLATE_PATH = r"C:\Path\To\Template.xlsx"
OUTPUT_FOLDER = r"C:\Path\To\output_value_versions"

ENTITY_SHEET = "Region & Function"
ENTITY_COLUMN = "C"       # ✅ Entity is in column C
ENTITY_START_ROW = 2      # Header in row 1

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"        # Dropdown cell

SHEETS_TO_EXPORT = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

# ===============================================


def sanitize_filename(name):
    """Remove invalid characters for Windows filenames"""
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def copy_sheet_as_values(source_wb, target_wb, sheet_name):
    src_ws = source_wb.Sheets(sheet_name)
    tgt_ws = target_wb.Sheets.Add()
    tgt_ws.Name = sheet_name

    used_range = src_ws.UsedRange

    # Copy values
    used_range.Copy()
    tgt_ws.Range("A1").PasteSpecial(Paste=-4163)   # xlPasteValues
    tgt_ws.Range("A1").PasteSpecial(Paste=-4122)   # xlPasteFormats

    # Copy column widths
    for col in range(1, used_range.Columns.Count + 1):
        tgt_ws.Columns(col).ColumnWidth = src_ws.Columns(col).ColumnWidth


def main():

    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

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

        entity = str(entity).strip()
        safe_entity = sanitize_filename(entity)

        print(f"Processing: {entity}")

        # Set entity dropdown
        ws_landing.Range(ENTITY_CELL).Value = entity

        # Refresh workbook
        wb.RefreshAll()
        excel.CalculateFull()
        time.sleep(2)

        # Create new workbook
        new_wb = excel.Workbooks.Add()

        # Remove default sheets
        while new_wb.Sheets.Count > 0:
            new_wb.Sheets(1).Delete()

        # Copy required sheets
        for sheet in SHEETS_TO_EXPORT:
            copy_sheet_as_values(wb, new_wb, sheet)

        # Save as Entity.xlsx
        save_path = os.path.join(OUTPUT_FOLDER, f"{safe_entity}.xlsx")
        new_wb.SaveAs(save_path, FileFormat=51)  # xlsx
        new_wb.Close(False)

    wb.Close(False)
    excel.Quit()

    print("✅ All entity value files created successfully")


if __name__ == "__main__":
    main()
