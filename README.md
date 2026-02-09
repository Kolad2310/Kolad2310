```
import pandas as pd
import shutil
import xlwings as xw
import os

MASTER_RESULTS = r"C:\path\Master.xlsx"
TEMPLATE_PATH  = r"C:\path\Presentation_Template.xlsx"
OUTPUT_FOLDER  = r"C:\path\Output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Read calculated results ONCE
df = pd.read_excel(MASTER_RESULTS, sheet_name="MODEL_OUTPUT")

app = xw.App(visible=False)
app.display_alerts = False

entities = df["Entity"].unique()

for entity in entities:
    out_path = os.path.join(OUTPUT_FOLDER, f"{entity}.xlsx")
    shutil.copy2(TEMPLATE_PATH, out_path)

    wb = app.books.open(out_path)

    entity_df = df[df["Entity"] == entity]

    # Example injections (you control mapping)
    wb.sheets["SSV Perf view"].range("B5").value = entity_df["Revenue"].iloc[0]
    wb.sheets["SSV Cost Perf view"].range("C7").value = entity_df["Cost"].iloc[0]
    wb.sheets["By Sector YTD"].range("D10").value = entity_df["YTD_Rev"].iloc[0]

    wb.save()
    wb.close()

app.quit()
