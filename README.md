```
import os
import re
import win32com.client as win32


TEMPLATE_PATH = r"C:\Path\To\Template.xlsx"
OUTPUT_FOLDER = r"C:\Path\To\output_value_versions"

ENTITY_SHEET = "Region & Function"
ENTITY_COLUMN = "C"
ENTITY_START_ROW = 2

SHEETS_TO_COPY = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]


def sanitize_filename(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def main():

    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(TEMPLATE_PATH)
    ws_entity = wb.Sheets(ENTITY_SHEET)

    last_row = ws_entity.Cells(
        ws_entity.Rows.Count, ENTITY_COLUMN
    ).End(-4162).Row  # xlUp

    for row in range(ENTITY_START_ROW, last_row + 1):

        entity = ws_entity.Cells(row, ENTITY_COLUMN).Value
        if not entity:
            continue

        entity = str(entity).strip()
        safe_entity = sanitize_filename(entity)

        print(f"Creating file for: {entity}")

        new_wb = None

        for idx, sheet_name in enumerate(SHEETS_TO_COPY):

            if idx == 0:
                # FIRST sheet: Excel creates a new workbook automatically
                wb.Sheets(sheet_name).Copy()
                new_wb = excel.ActiveWorkbook
            else:
                # NEXT sheets: append to existing workbook
                wb.Sheets(sheet_name).Copy(
                    After=new_wb.Sheets(new_wb.Sheets.Count)
                )

        save_path = os.path.join(OUTPUT_FOLDER, f"{safe_entity}.xlsx")
        new_wb.SaveAs(save_path, FileFormat=51)
        new_wb.Close(False)

    wb.Close(False)
    excel.Quit()

    print("✅ Files created successfully without COM errors")


if __name__ == "__main__":
    main()
