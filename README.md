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

INPUT_SHEETS = {
    "P&L": {"entity_col": "E", "start_row": 24},
    "BS":  {"entity_col": "E", "start_row": 22},
    "SD":  {"entity_col": "E", "start_row": 22}
}

OUTPUT_SHEETS = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

# =========================================


def normalize(val):
    """Normalize entity for safe comparison"""
    if pd.isna(val):
        return ""
    return str(val).strip().upper()


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def read_input_data():
    """Read full sheets including headers"""
    data = {}
    for sheet in INPUT_SHEETS:
        data[sheet] = pd.read_excel(
            INPUT_DATA_FILE,
            sheet_name=sheet,
            header=0
        )
    return data


def filter_entity_data(df, entity, entity_col_letter):
    """Filter rows for one entity (robust match)"""
    col_idx = ord(entity_col_letter.upper()) - ord("A")
    entity_norm = normalize(entity)

    mask = df.iloc[:, col_idx].apply(normalize) == entity_norm
    return df.loc[mask]


def write_dataframe_to_sheet(sheet, df, start_row):
    """Write dataframe INCLUDING headers alignment"""
    if df.empty:
        return

    # Clear old content first (important)
    sheet.range((start_row, 1),
                (sheet.cells.last_cell.row,
                 sheet.cells.last_cell.column)).clear_contents()

    # Write values only (no headers row)
    sheet.range((start_row, 1)).value = df.values


def freeze_sheet(sheet):
    """Convert formulas to values (format preserved)"""
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

        # ---- Fill input sheets ----
        for sheet_name, cfg in INPUT_SHEETS.items():
            filtered = filter_entity_data(
                input_data[sheet_name],
                entity,
                cfg["entity_col"]
            )

            write_dataframe_to_sheet(
                wb.sheets[sheet_name],
                filtered,
                cfg["start_row"]
            )

        # 🔑 Force calculation (template is small → fast)
        app.calculate()

        # ---- Freeze output sheets ----
        for sheet_name in OUTPUT_SHEETS:
            freeze_sheet(wb.sheets[sheet_name])

        wb.save()
        wb.close()

        print(f"   Saved → {out_path}")

    app.quit()
    print("\n✅ ALL ENTITY FILES CREATED SUCCESSFULLY")


# ================= ENTRY POINT =================

if __name__ == "__main__":
    process_entities()
