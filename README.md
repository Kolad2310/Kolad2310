```
import pandas as pd

input_files = [
    r"C:\data\file1.xlsx",
    r"C:\data\file2.xlsx",
    r"C:\data\file3.xlsx",
    r"C:\data\file4.xlsx",
]

output_file = r"C:\data\combined_output.xlsx"

SHEETS = {
    "P&L": 22,   # Excel row 23
    "BS": 20,    # Excel row 21
    "SD": 20
}

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

    for sheet_name, header_row in SHEETS.items():
        all_data = []

        for file in input_files:
            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=header_row
            )

            # ✅ Remove rows where ALL columns are NA
            df = df.dropna(how="all")

            all_data.append(df)

        combined_df = pd.concat(all_data, ignore_index=True)

        combined_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

print("✅ Sheets appended successfully (blank rows skipped)")
