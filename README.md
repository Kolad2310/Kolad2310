```
import os
import re
import win32com.client as win32


# ====== CHANGE ONLY THESE TWO ======
TEMPLATE_PATH = r"C:\FULL\ABSOLUTE\PATH\Template.xlsx"
OUTPUT_FOLDER = r"C:\FULL\ABSOLUTE\PATH\output_value_versions"
# ==================================


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

    # 🔒 HARD guarantee folder exists
    OUTPUT_FOLDER_ABS = os.path.abspath(OUTPUT_FOLDER)
    os.makedirs(OUTPUT_FOLDER_ABS, exist_ok=True)

    print("OUTPUT FOLDER USED:")
    print(OUTPUT_FOLDER_ABS)
    print("-" * 50)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(os.path.abspath(TEMPLATE_PATH))
    ws_entity = wb.Sheets(ENTITY_SHEET)

    last_row = ws_entity.Cells(
        ws_entity.Rows.Count, ENTITY_COLUMN
    ).End(-4162).Row  # xlUp

    for row in range(ENTITY_START_ROW, last_row + 1):

        entity = ws_entity.Cells(row, ENTITY_COLUMN).Value
        if not entity:
            continue

        entity = sanitize_filename(entity)
        print(f"Creating file for entity: {entity}")

        new_wb = None

        for idx, sheet_name in enumerate(SHEETS_TO_COPY):

            if idx == 0:
                wb.Sheets(sheet_name).Copy()
                new_wb = excel.ActiveWorkbook
            else:
                wb.Sheets(sheet_name).Copy(
                    After=new_wb.Sheets(new_wb.Sheets.Count)
                )

        save_path = os.path.join(OUTPUT_FOLDER_ABS, f"{entity}.xlsx")
        print("Saving to:", save_path)

        # 🔒 Force overwrite-safe save
        if os.path.exists(save_path):
            os.remove(save_path)

        new_wb.SaveAs(
            Filename=save_path,
            FileFormat=51  # xlsx
        )

        new_wb.Close(False)

    wb.Close(False)
    excel.Quit()

    print("✅ DONE. Files MUST exist in the output folder now.")


if __name__ == "__main__":
    main()
