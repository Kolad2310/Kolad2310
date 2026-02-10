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

        # ---- Read prefix rows from FIRST file only ----
        first_raw = pd.read_excel(
            input_files[0],
            sheet_name=sheet_name,
            header=None
        )

        prefix_df = first_raw.iloc[:header_row]
        header = first_raw.iloc[header_row].tolist()

        # ---- Read & append data from ALL files ----
        for file in input_files:
            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=header_row
            )

            # Skip rows where all values are NA
            df = df.dropna(how="all")

            all_data.append(df)

        combined_df = pd.concat(all_data, ignore_index=True)

        # ---- Write output ----
        start_row = 0

        # Prefix rows
        prefix_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )
        start_row += len(prefix_df)

        # Header row
        pd.DataFrame([header]).to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )
        start_row += 1

        # Data
        combined_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )

print("✅ Prefix rows kept, data appended, blank rows skipped")
