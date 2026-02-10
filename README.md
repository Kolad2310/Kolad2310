```
import pandas as pd
from pathlib import Path

# ---------------- CONFIG ----------------
input_files = [
    r"C:\data\file1.xlsx",
    r"C:\data\file2.xlsx",
    r"C:\data\file3.xlsx",
    r"C:\data\file4.xlsx",
]

output_file = r"C:\data\combined_output.xlsx"

HEADER_ROWS = {
    "P&L": 22,   # Excel row 23
    "BS": 20,    # Excel row 21
    "SD": 20
}
# ----------------------------------------


def process_sheet(sheet_name, header_row):
    """Returns prefix_df, header, combined_data_df"""
    combined_data = []
    prefix_df = None
    header = None

    for i, file in enumerate(input_files):
        # Read entire sheet without headers
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)

        if i == 0:
            # Rows before header → keep as-is
            prefix_df = df_raw.iloc[:header_row]

            # Extract header
            header = df_raw.iloc[header_row].tolist()

        # Extract data below header
        data_df = df_raw.iloc[header_row + 1:].copy()
        data_df.columns = header

        combined_data.append(data_df)

    combined_data_df = pd.concat(combined_data, ignore_index=True)
    return prefix_df, header, combined_data_df


# ---------------- WRITE OUTPUT ----------------
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    for sheet_name, header_row in HEADER_ROWS.items():
        prefix_df, header, combined_data_df = process_sheet(sheet_name, header_row)

        start_row = 0

        # Write prefix rows
        prefix_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )

        start_row += len(prefix_df)

        # Write header
        pd.DataFrame([header]).to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )

        start_row += 1

        # Write appended data
        combined_data_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )

print("✅ Combined file created successfully")
