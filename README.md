```
import os
import shutil
import time
import re
import xlwings as xw


# ========== CONFIG ==========
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
# ============================


def safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def get_entities():
    """Read entity list once (fast & safe)"""
    app = xw.App(visible=False)
    wb = app.books.open(MASTER_PATH, read_only=True)

    ws = wb.sheets[ENTITY_SHEET]
    last_row = ws.range(f"{ENTITY_COLUMN}{ws.cells.last_cell.row}").end("up").row

    entities = []
    for r in range(ENTITY_START_ROW, last_row + 1):
        val = ws.range(f"{ENTITY_COLUMN}{r}").value
        if val:
            entities.append(safe_name(val))

    wb.close()
    app.quit()
    return entities


def main():

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    entities = get_entities()

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    for entity in entities:
        print(f"\n▶ Processing entity: {entity}")

        # 1️⃣ Copy + rename master
        out_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        if os.path.exists(out_path):
            os.remove(out_path)

        shutil.copy2(MASTER_PATH, out_path)

        # 2️⃣ Open copied file
        wb = app.books.open(out_path)

        # 3️⃣ Set entity
        wb.sheets[LANDING_SHEET].range(ENTITY_CELL).value = entity

        # 4️⃣ Refresh (xlwings-safe)
        wb.api.RefreshAll()
        app.api.CalculateFull()

        # Give Excel time to finish background refresh
        time.sleep(2)

        # 5️⃣ Delete unwanted sheets
        for sheet in wb.sheets:
            if sheet.name not in SHEETS_TO_KEEP:
                sheet.delete()

        # 6️⃣ Save & close
        wb.save()
        wb.close()

        print(f"   Saved → {out_path}")

    app.quit()
    print("\n✅ All entity files created successfully using xlwings")


if __name__ == "__main__":
    main()
