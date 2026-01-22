```
import xlwings as xw
import os

folder_path = r"C:\your\folder\path"   # 🔁 change this

cells_to_update = ["G13", "G14", "G15", "G16"]

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

try:
    for file in os.listdir(folder_path):
        if file.endswith((".xlsx", ".xlsm", ".xls")):
            file_path = os.path.join(folder_path, file)

            wb = app.books.open(file_path)
            sht = wb.sheets["Template"]

            for cell in cells_to_update:
                val = sht.range(cell).value
                if isinstance(val, (int, float)):
                    sht.range(cell).value = val * -1

            wb.save()
            wb.close()

finally:
    app.quit()
