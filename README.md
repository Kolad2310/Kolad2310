```
import xlwings as xw
import os
import time

# =========================
# CONFIGURATION
# =========================
template_path = r"C:\input\template.xlsx"
output_folder = r"C:\output"

sheet_name = "F1 Landing Page DB"
target_cell = "A1"   # cell to replace with list item

entity_list = [
    "India",
    "UK",
    "USA",
    "Germany"
]

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)

# =========================
# PROCESS
# =========================
app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

try:
    for entity in entity_list:
        print(f"Processing: {entity}")

        wb = app.books.open(template_path)

        sht = wb.sheets[sheet_name]

        # Replace value with list item
        sht.range(target_cell).value = entity

        # Refresh all connections / queries
        wb.api.RefreshAll()

        # Give Excel time to complete refresh
        time.sleep(5)

        # Save new file
        output_file = os.path.join(
            output_folder,
            f"F1_Landing_{entity}.xlsx"
        )

        wb.save(output_file)
        wb.close()

        print(f"Saved: {output_file}")

finally:
    app.quit()
