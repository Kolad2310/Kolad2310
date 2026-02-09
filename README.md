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

# Absolute Excel layout (DO NOT CHANGE)
INPUT_SHEETS = {
    "P&L": {"entity_col_idx": 4, "header_row": 23},  # Col E = index 4
    "BS":  {"entity_col_idx": 4, "header_row": 21},
    "SD":  {"entity_col_idx": 4, "header_row": 21}
}

OUTPUT_SHEETS = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

# =========================================


def normalize(v):
    if pd.isna(v):
        return ""
    return str(v).strip().upper()


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def read_input_raw():
    """Read sheets as raw Excel grids (NO headers)"""
    data = {}
    for sheet in INPUT_SHEETS:
        data[sheet] = pd.read_excel(
            INPUT_DATA_FILE,
            sheet_name=sheet,
            header=None
        )
    return data


def filter_rows(df, entity, entity_col_idx, header_row):
    """
    Filter rows strictly BELOW header_row
    and where column E matches entity
    """
    entity_norm = normalize(entity)

    data_only = df.iloc[header_row:]  # rows below header
    mask = data_only.iloc[:, entity_col_idx].apply(normalize) == entity_norm

    return data_only.loc[mask]


def write_rows(sheet, rows_df, header_row):
    """
    Write raw rows BELOW header row
    preserving exact Excel layout
    """
    start_row = header_row + 1

    # Clear everything below header
    sheet.range(
        (start_row, 1),
        (sheet.cells.last_cell.row, sheet.cells.last_cell.column)
    ).clear_contents()

    if rows_df.empty:
        return

    sheet.range((start_row, 1)).value = rows_df.values


def freeze_sheet(sheet):
    used = sheet.used_range
    if used:
        used.value = used.value


def process_entities():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    print("📖 Reading input data once (raw mode)...")
    input_data = read_input_raw()

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    for i, entity in enumerate(ENTITIES, start=1):
        entity_safe = safe_name(entity)
        print(f"\n[{i}/{len(ENTITIES)}] Processing {entity_safe}")

        out_path = os.path.join(OUTPUT_FOLDER, f"{entity_safe}.xlsx")
        shutil.copy2(TEMPLATE_FILE, out_path)

        wb = app.books.open(out_path)

        # 1️⃣ Fill P&L / BS / SD inputs
        for sheet_name, cfg in INPUT_SHEETS.items():
            rows = filter_rows(
                input_data[sheet_name],
                entity,
                cfg["entity_col_idx"],
                cfg["header_row"]
            )

            write_rows(
                wb.sheets[sheet_name],
                rows,
                cfg["header_row"]
            )

        # 2️⃣ Set entity control cell (CRITICAL)
        wb.sheets[LANDING_SHEET].range(ENTITY_CELL).value = entity

        # 3️⃣ Calculate (template calc is fast)
        app.calculate()

        # 4️⃣ Freeze outputs
        for s in OUTPUT_SHEETS:
            freeze_sheet(wb.sheets[s])

        wb.save()
        wb.close()

        print(f"   ✔ Saved {entity_safe}.xlsx")

    app.quit()
    print("\n✅ ALL ENTITY FILES CREATED WITH VALUES")


# ================= ENTRY POINT =================

if __name__ == "__main__":
    process_entities()
