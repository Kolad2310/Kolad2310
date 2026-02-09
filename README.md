```
import os
import shutil
import time
import re
import xlwings as xw


# =====================================================
# CONFIGURATION (CHANGE THESE)
# =====================================================

MASTER_PATH = r"C:\FULL\PATH\Master.xlsx"

BASE_FOLDER = r"C:\FULL\PATH\entity_processing"

PREP_FOLDER  = os.path.join(BASE_FOLDER, "01_prepared_files")
CALC_FOLDER  = os.path.join(BASE_FOLDER, "02_calculated_files")
VALUE_FOLDER = os.path.join(BASE_FOLDER, "03_value_files")

ENTITIES = [
    "APAC",
    "EMEA",
    "AMERICAS",
    "INDIA",
    "UK"
]

LANDING_SHEET = "Landing Page DB"
ENTITY_CELL   = "F1"

SHEETS_TO_KEEP = [
    "Landing Page DB",
    "SSV Perf view",
    "SSV Cost Perf view",
    "By Sector YTD"
]

# =====================================================


def safe_name(name: str) -> str:
    """Make entity safe for Windows filename"""
    return re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())


# =====================================================
# PHASE 1 – PREPARE FILES (FAST)
# =====================================================

def phase_1_prepare_files():
    os.makedirs(PREP_FOLDER, exist_ok=True)

    app = xw.App(visible=False)
    app.display_alerts = False

    print("\n--- PHASE 1: PREPARING FILES ---")

    for entity in ENTITIES:
        entity = safe_name(entity)
        out_path = os.path.join(PREP_FOLDER, f"{entity}.xlsx")

        if os.path.exists(out_path):
            os.remove(out_path)

        shutil.copy2(MASTER_PATH, out_path)

        wb = app.books.open(out_path)
        wb.sheets[LANDING_SHEET].range(ENTITY_CELL).value = entity
        wb.save()
        wb.close()

        print(f"Prepared: {out_path}")

    app.quit()


# =====================================================
# PHASE 2 – FULL RECALCULATION (SLOW)
# =====================================================

def phase_2_recalculate_all():
    os.makedirs(CALC_FOLDER, exist_ok=True)

    app = xw.App(visible=False)
    app.display_alerts = False

    print("\n--- PHASE 2: FULL RECALCULATION ---")

    for file in os.listdir(PREP_FOLDER):
        if not file.lower().endswith(".xlsx"):
            continue

        src = os.path.join(PREP_FOLDER, file)
        dst = os.path.join(CALC_FOLDER, file)

        if os.path.exists(dst):
            os.remove(dst)

        shutil.copy2(src, dst)

        wb = app.books.open(dst)

        # Equivalent of Ctrl + Alt + Shift + F9
        app.api.CalculateFullRebuild()

        # Give Excel time to complete heavy calc
        time.sleep(5)

        wb.save()
        wb.close()

        print(f"Calculated: {dst}")

    app.quit()


# =====================================================
# PHASE 3 – CREATE VALUE VERSIONS (FAST)
# =====================================================

def freeze_sheet_to_values(sheet: xw.Sheet):
    used = sheet.used_range
    if used is not None:
        used.value = used.value


def phase_3_create_value_versions():
    os.makedirs(VALUE_FOLDER, exist_ok=True)

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    print("\n--- PHASE 3: CREATING VALUE VERSIONS ---")

    for file in os.listdir(CALC_FOLDER):
        if not file.lower().endswith(".xlsx"):
            continue

        src = os.path.join(CALC_FOLDER, file)
        dst = os.path.join(VALUE_FOLDER, file)

        if os.path.exists(dst):
            os.remove(dst)

        shutil.copy2(src, dst)

        wb = app.books.open(dst)

        # Freeze required sheets
        for sheet_name in SHEETS_TO_KEEP:
            freeze_sheet_to_values(wb.sheets[sheet_name])

        # Delete all other sheets
        for sheet in wb.sheets:
            if sheet.name not in SHEETS_TO_KEEP:
                sheet.delete()

        wb.save()
        wb.close()

        print(f"Value version created: {dst}")

    app.quit()


# =====================================================
# ENTRY POINT
# =====================================================

if __name__ == "__main__":

    phase_1_prepare_files()
    phase_2_recalculate_all()
    phase_3_create_value_versions()

    print("\n✅ ALL PHASES COMPLETED SUCCESSFULLY")
