```
import os
import shutil
import time
import re
import xlwings as xw


# ================= CONFIG =================

MASTER_PATH = r"C:\FULL\PATH\Master.xlsx"
OUTPUT_FOLDER = r"C:\FULL\PATH\output_files"

ENTITY_SHEET = "Region & Function"
ENTITY_COLUMN = "C"
ENTITY_START_ROW = 2

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"

SHEETS_TO_KEEP = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

# =========================================


def safe_name(name: str) -> str:
    """Make entity safe for Windows filename"""
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def get_entities() -> list:
    """Read entity list once from master (fast & safe)"""
    app = xw.App(visible=False)
    app.display_alerts = False

    wb = app.books.open(MASTER_PATH, read_only=True)
    ws = wb.sheets[ENTITY_SHEET]

    last_row = ws.range(
        f"{ENTITY_COLUMN}{ws.cells.last_cell.row}"
    ).end("up").row

    entities = []
    for r in range(ENTITY_START_ROW, last_row + 1):
        val = ws.range(f"{ENTITY_COLUMN}{r}").value
        if val:
            entities.append(safe_name(val))

    wb.close()
    app.quit()
    return entities


def freeze_sheet_to_values(sheet: xw.Sheet):
    """
    Convert a sheet to values ONLY (formatting preserved)
    This is safe because it does NOT touch other sheets.
    """
    used = sheet.used_range
    if used is not None:
        used.value = used.value


def process_entities(entities: list):

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    for idx, entity in enumerate(entities, start=1):
        print(f"[{idx}/{len(entities)}] Processing {entity}")

        out_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        if os.path.exists(out_path):
            os.remove(out_path)

        # 1️⃣ Copy master and rename
        shutil.copy2(MASTER_PATH, out_path)

        # 2️⃣ Open copied file (Excel refreshes on open)
        wb = app.books.open(out_path)

        # 3️⃣ Set entity
        wb.sheets[LANDING_SHEET].range(ENTITY_CELL).value = entity

        # 4️⃣ Wait for heavy calculation to finish
        # (formula-intensive model)
        time.sleep(10)

        # 5️⃣ Freeze required sheets to VALUES
        for sheet_name in SHEETS_TO_KEEP:
            freeze_sheet_to_values(wb.sheets[sheet_name])

        # 6️⃣ Delete all other (reference) sheets
        for sheet in wb.sheets:
            if sheet.name not in SHEETS_TO_KEEP:
                sheet.delete()

        # 7️⃣ Save & close
        wb.save()
        wb.close()

        print(f"     Value version saved → {out_path}")

    app.quit()


# ================= ENTRY POINT =================

if __name__ == "__main__":
    entities = get_entities()
    print(f"Found {len(entities)} entities")
    process_entities(entities)
    print("✅ ALL VALUE FILES CREATED SUCCESSFULLY")
