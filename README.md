```
import os
import shutil
import time
import re
import xlwings as xw


MASTER_PATH = r"C:\FULL\PATH\Master.xlsx"
OUTPUT_FOLDER = r"C:\FULL\PATH\output_files"

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL = "F1"

SHEETS_TO_KEEP = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]


def safe(name):
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


def main(entities):

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    for entity in entities:
        entity = safe(entity)
        print(f"Processing {entity}")

        out_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
        shutil.copy2(MASTER_PATH, out_path)

        # Open → Excel auto-refreshes & recalculates
        wb = app.books.open(out_path)
        wb.sheets[LANDING_SHEET].range(ENTITY_CELL).value = entity

        # Give Excel time to finish internal calc
        time.sleep(10)

        # Delete unwanted sheets
        for s in wb.sheets:
            if s.name not in SHEETS_TO_KEEP:
                s.delete()

        wb.save()
        wb.close()

    app.quit()
    print("DONE")
