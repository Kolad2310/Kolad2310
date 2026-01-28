```
import os
from openpyxl import load_workbook

folder_path = r"PATH_TO_YOUR_FOLDER"  # <-- change this

for file in os.listdir(folder_path):
    if file.endswith(".xlsx"):
        file_path = os.path.join(folder_path, file)

        # Extract country from filename
        # example: abc_India_temomates.xlsx
        country = file.split("_")[1]

        wb = load_workbook(file_path)

        if "CFO_SingOff" in wb.sheetnames:
            ws = wb["CFO_SingOff"]

            cell_value = ws["A1"].value or ""

            # Keep text till last hyphen and append country
            if "-" in cell_value:
                base_text = cell_value.rsplit("-", 1)[0].strip()
                ws["A1"] = f"{base_text} - {country}"
            else:
                ws["A1"] = f"{cell_value} - {country}"

            wb.save(file_path)
            print(f"Updated: {file}")

        else:
            print(f"CFO_SingOff sheet not found in {file}")
