```
import os
import shutil
import pandas as pd
import xlwings as xw
import re


# ================= CONFIG =================

INPUT_DATA_FILE = r"C:\PATH\Input_Data.xlsx"
TEMPLATE_FILE   = r"C:\PATH\Template.xlsx"
OUTPUT_FOLDER   = r"C:\PATH\Output_Entity_Files"

ENTITIES = ["APAC", "EMEA", "INDIA", "AMERICAS", "UK"]

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL   = "F1"

INPUT_SHEETS = {
    "P&L": {"entity_col": "E", "header_row": 23},
    "BS":  {"entity_col": "E", "header_row": 21},
    "SD":  {"entity_col": "E", "header_row": 21}
}

OUTPUT_SHEETS = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

# =========================================


def normalize(val):
    if pd.isna(val):
        return ""
    return str(val).strip().upper()


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def read_input_data():
    data = {}
    for sheet in INPUT_SHEETS:
        data[sheet] = pd.read_excel(INPUT_DATA_FILE, sheet_name=sheet)
    return data


def filter_entity_data(df, entity, entity_col_letter):
    col_idx = ord(entity_col_letter.upper()) - ord("A")
    entity_norm = normalize(entity)
    return df[df.iloc[:, col_idx].apply(normalize) == entity_norm]


def write_dataframe_to_sheet(sheet, df, header_row):
    if df.empty:
        return

    start_row = header_row + 1
    num_rows, num_cols = df.shape

    # Clear old data only below headers
    sheet.range(
        (start_row, 1),
        (sheet.cells.last_cell.row, num_cols)
    ).clear_contents()

    # Write data (NO headers)
    sheet.range((start_row, 1)).value = df.values


def freeze_sheet(sheet):
    used = sheet.used_range
    if used:
        used.value = used.value


def process_entities():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    print("📖 Reading input data once...")
    input_data = read_input_data()

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    for idx, entity in enumerate(ENTITIES, start=1):
        entity_safe = safe_name(entity)
        print(f"\n[{idx}/{len(ENTITIES)}] Processing {entity_safe}")

        out_path = os.path.join(OUTPUT_FOLDER, f"{entity_safe}.xlsx")
        shutil.copy2(TEMPLATE_FILE, out_path)

        wb = app.books.open(out_path)

        # 1️⃣ Fill input sheets
        for sheet_name, cfg in INPUT_SHEETS.items():
            filtered_df = filter_entity_data(
                input_data[sheet_name],
                entity,
                cfg["entity_col"]
            )

            write_dataframe_to_sheet(
                wb.sheets[sheet_name],
                filtered_df,
                cfg["header_row"]
            )

        # 2️⃣ Set entity (critical for formulas)
        wb.sheets[LANDING_SHEET].range(ENTITY_CELL).value = entity

        # 3️⃣ Recalculate
        app.calculate()

        # 4️⃣ Freeze output sheets
        for s in OUTPUT_SHEETS:
            freeze_sheet(wb.sheets[s])

        wb.save()
        wb.close()

        print(f"   Saved → {out_path}")

    app.quit()
    print("\n✅ ALL ENTITY FILES CREATED CORRECTLY")


# ================= ENTRY POINT =================

if __name__ == "__main__":
    process_entities()
