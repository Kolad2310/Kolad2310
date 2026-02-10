```
import os
import xlwings as xw

SOURCE_FOLDER = r"C:\input_excels"
TARGET_FOLDER = r"C:\value_excels"

os.makedirs(TARGET_FOLDER, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

for file in os.listdir(SOURCE_FOLDER):
    if file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        src_path = os.path.join(SOURCE_FOLDER, file)
        tgt_path = os.path.join(TARGET_FOLDER, file)

        wb = app.books.open(src_path)

        for sheet in wb.sheets:
            used = sheet.used_range
            if used is not None:
                used.value = used.value  # 🔥 formulas → values

        wb.save(tgt_path)
        wb.close()

app.quit()

print("✅ Value versions created successfully")
