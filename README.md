```
import os
import xlwings as xw
from pywintypes import com_error

SOURCE_FOLDER = r"C:\input_excels"
TARGET_FOLDER = r"C:\value_excels"

L = ["Sheet1", "Summary", "F1 Landing Page DB"]  # sheets to process

os.makedirs(TARGET_FOLDER, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

for file in os.listdir(SOURCE_FOLDER):

    # 🚫 skip Excel temp / lock files
    if file.startswith("~$"):
        continue

    if not file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        continue

    print(f"\nProcessing file: {file}")
    wb = app.books.open(os.path.join(SOURCE_FOLDER, file))

    for sheet in wb.sheets:
        if sheet.name not in L:
            continue

        try:
            sheet.api.Cells.Copy()
            sheet.api.Cells.PasteSpecial(Paste=-4122)  # formats
            sheet.api.Cells.PasteSpecial(Paste=-4163)  # values

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

print("\n✅ Done — temp files skipped, lift-and-shift clean")
