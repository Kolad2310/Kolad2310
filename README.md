```
import xlwings as xw
import os

source_folder = r"C:\source_folder_path"      # 🔁 change this
target_folder = r"C:\target_folder_path"      # 🔁 change this

os.makedirs(target_folder, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

try:
    for file in os.listdir(source_folder):
        if file.endswith((".xlsx", ".xlsm")):
            src_path = os.path.join(source_folder, file)

            name, ext = os.path.splitext(file)
            tgt_file = f"{name}_values{ext}"
            tgt_path = os.path.join(target_folder, tgt_file)

            wb = app.books.open(src_path)

            # -------- Relationship Management --------
            if "Relationship Management" in [s.name for s in wb.sheets]:
                sht = wb.sheets["Relationship Management"]
                for col in ["B", "D"]:
                    rng = sht.range(f"{col}:{col}")
                    rng.value = rng.value  # convert formulas → values

            # -------- Relationship Maintenance --------
            if "Relationship Maintenance" in [s.name for s in wb.sheets]:
                sht = wb.sheets["Relationship Maintenance"]
                rng = sht.range("J:J")
                rng.value = rng.value

            # -------- Structural cost --------
            if "Structural cost" in [s.name for s in wb.sheets]:
                sht = wb.sheets["Structural cost"]
                rng = sht.range("D:D")
                rng.value = rng.value

            wb.save(tgt_path)
            wb.close()

finally:
    app.quit()
