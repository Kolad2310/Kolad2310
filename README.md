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

ENTITIES = [
    "APAC",
    "EMEA",
    "INDIA",
    "AMERICAS",
    "UK"
]

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL   = "F1"

# Input sheet config: entity column + start row
INPUT_SHEETS = {
    "P&L": {"entity_col": "E", "start_row": 24},
    "BS":  {"entity_col": "E", "start_row": 22},
    "SD":  {"entity_col": "E", "start_row": 22}
}

# Sheets whose values must be frozen
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
    """Read input sheets once into memory"""
    data = {}
    for sheet in INPUT_SHEETS:
        data[sheet] = pd.read_excel(
            INPUT_DATA_FILE,
            sheet_name=sheet,
            header=0
        )
    return data


def filter_entity_data(df, entity, entity_col_letter):
    """Robust entity filter"""
    col_idx = ord(entity_col_letter.upper()) - ord("A")
    entity_norm = normalize(entity)
    return df[df.iloc[:, col_idx].apply(normalize) == entity_norm]


def write_dataframe_to_sheet(sheet, df, start_row):
    """Write values into template input sheet"""
    if df.empty:
        return

    # Clear old data
    sheet.range(
        (start_row, 1),
        (sheet.cells.last_cell.row, sheet.cells.last_cell.column)
    ).clear_contents()

    # Write new data (values only)
    sheet.range((start_row, 1)).value = df.values


def freeze_sheet(sheet):
    """Convert formulas to values, keep formatting"""
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

        # 1️⃣ Fill input sheets with entity-filtered data
        for sheet_name, cfg in INPUT_SHEETS.items():
            filtered_df = filter_entity_data(
                input_data[sheet_name],
                entity,
                cfg["entity_col"]
            )

            write_dataframe_to_sheet(
                wb.sheets[sheet_name],
                filtered_df,
                cfg["start_row"]
            )

        # 2️⃣ SET ENTITY IN LANDING PAGE (CRITICAL)
        wb.sheets[LANDING_SHEET].range(ENTITY_CELL).value = entity

        # 3️⃣ FORCE CALC (template calc is fast)
        app.calculate()

        # 4️⃣ Freeze output sheets to VALUES
        for sheet_name in OUTPUT_SHEETS:
            freeze_sheet(wb.sheets[sheet_name])

        wb.save()
        wb.close()

        print(f"   Saved → {out_path}")

    app.quit()
    print("\n✅ ALL ENTITY FILES CREATED WITH VALUES")


# ================= ENTRY POINT =================

if __name__ == "__main__":
    process_entities()
