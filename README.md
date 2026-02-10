```
import os
import xlwings as xw
from pywintypes import com_error

SOURCE_FOLDER = r"C:\input_excels"
TARGET_FOLDER = r"C:\value_excels"

# 👇 Only these sheets will be processed
L = ["Sheet1", "Summary", "F1 Landing Page DB"]  # <-- edit this

os.makedirs(TARGET_FOLDER, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

for file in os.listdir(SOURCE_FOLDER):
    if not file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        continue

    print(f"\nProcessing file: {file}")
    wb = app.books.open(os.path.join(SOURCE_FOLDER, file))

    for sheet in wb.sheets:
        if sheet.name not in L:
            continue  # ⛔ skip all other sheets

        try:
            # PURE lift-and-shift (Excel-native)
            sheet.api.Cells.Copy()
            sheet.api.Cells.PasteSpecial(Paste=-4122)  # xlPasteFormats
            sheet.api.Cells.PasteSpecial(Paste=-4163)  # xlPasteValues

        except com_error as e:
            print(
                f"❌ COM ERROR\n"
                f"   File  : {file}\n"
                f"   Sheet : {sheet.name}\n"
                f"   Cell  : Cells (entire sheet)\n"
                f"   Error : {e}"
            )

    wb.save(os.path.join(TARGET_FOLDER, file))
    wb.close()

app.quit()

print("\n✅ Lift-and-shift completed (only sheets in list L)")
