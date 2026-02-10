```
import os
import xlwings as xw
import pythoncom
from pywintypes import com_error

SOURCE_FOLDER = r"C:\input_excels"
TARGET_FOLDER = r"C:\value_excels"

os.makedirs(TARGET_FOLDER, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

for file in os.listdir(SOURCE_FOLDER):
    if not file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        continue

    src_path = os.path.join(SOURCE_FOLDER, file)
    tgt_path = os.path.join(TARGET_FOLDER, file)

    print(f"\nProcessing file: {file}")
    wb = app.books.open(src_path)

    for sheet in wb.sheets:
        try:
            used = sheet.used_range
            if used is None:
                continue

            # Copy → Paste formats
            used.copy()
            used.paste(paste="formats")

            # Paste values
            used.value = used.value

        except com_error as e:
            try:
                cell_address = used.address
            except Exception:
                cell_address = "UNKNOWN_RANGE"

            print(
                f"❌ COM ERROR\n"
                f"   File  : {file}\n"
                f"   Sheet : {sheet.name}\n"
                f"   Range : {cell_address}\n"
                f"   Error : {e}"
            )

        except Exception as e:
            print(
                f"❌ PYTHON ERROR\n"
                f"   File  : {file}\n"
                f"   Sheet : {sheet.name}\n"
                f"   Error : {e}"
            )

    wb.save(tgt_path)
    wb.close()

app.quit()

print("\n✅ Processing completed (with diagnostics)")
