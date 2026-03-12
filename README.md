```
import pandas as pd

input_files = [
    r"C:\data\file1.xlsx",
    r"C:\data\file2.xlsx",
    r"C:\data\file3.xlsx",
    r"C:\data\file4.xlsx",
]

output_file_sheets = r"C:\data\combined_output.xlsx"
output_file_all = r"C:\data\all_sheets_combined.xlsx"

SHEETS = {
    "P&L": 22,   # Excel row 23
    "BS": 20,    # Excel row 21
    "SD": 20
}

all_sheets_data = []
header_written = False
final_header = None

with pd.ExcelWriter(output_file_sheets, engine="xlsxwriter") as writer:

    for sheet_name, header_row in SHEETS.items():

        print(f"\nProcessing sheet: {sheet_name}")

        all_data = []

        # ---- Read prefix rows from first file ----
        first_raw = pd.read_excel(
            input_files[0],
            sheet_name=sheet_name,
            header=None
        )

        prefix_df = first_raw.iloc[:header_row]
        header = first_raw.iloc[header_row].tolist()

        # Save header for second output
        if not header_written:
            final_header = header
            header_written = True

        # ---- Read data from all files ----
        for file in input_files:
            print(f"Reading {file}")

            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                header=header_row
            )

            df = df.dropna(how="all")

            all_data.append(df)

        combined_df = pd.concat(all_data, ignore_index=True)

        # Save for second file
        all_sheets_data.append(combined_df)

        # ---- Write sheet-wise output ----
        start_row = 0

        prefix_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )

        start_row += len(prefix_df)

        pd.DataFrame([header]).to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )

        start_row += 1

        combined_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=start_row
        )

# ---- Second file (all sheets appended) ----
final_df = pd.concat(all_sheets_data, ignore_index=True)

with pd.ExcelWriter(output_file_all, engine="xlsxwriter") as writer:
    final_df.to_excel(
        writer,
        sheet_name="Combined",
        index=False,
        header=True
    )

print("\n✅ Sheet-wise consolidated file created:", output_file_sheets)
print("✅ All sheets appended into one sheet:", output_file_all)
