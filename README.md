```
import pandas as pd
import os

# ===============================
# CONFIG
# ===============================

INPUT_FILE = "input.xlsx"   # change path
OUTPUT_FOLDER = "output_files"
COMPONENT_COLUMN = "component"   # column used to split (change if needed)
MAX_ROWS = 1048576  # Excel limit

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ===============================
# READ FILE
# ===============================

df = pd.read_excel(INPUT_FILE)

# Clean column names (safe practice)
df.columns = df.columns.str.strip()

# Standardize component column
df[COMPONENT_COLUMN] = df[COMPONENT_COLUMN].astype(str).str.strip().str.upper()

# ===============================
# SPLIT DATA
# ===============================

rwa_df = df[df[COMPONENT_COLUMN] == "RWA"].copy()
avbs_df = df[df[COMPONENT_COLUMN] == "AVBS_SD"].copy()
pbt_bs_df = df[df[COMPONENT_COLUMN] == "PBT_BS"].copy()

print("RWA rows:", len(rwa_df))
print("AVBS rows:", len(avbs_df))
print("PBT_BS rows:", len(pbt_bs_df))

# ===============================
# ALIGN SCHEMA (KEEP ALL COLUMNS)
# ===============================

all_columns = []

for temp_df in [rwa_df, avbs_df, pbt_bs_df]:
    for col in temp_df.columns:
        if col not in all_columns:
            all_columns.append(col)

rwa_df = rwa_df.reindex(columns=all_columns)
avbs_df = avbs_df.reindex(columns=all_columns)
pbt_bs_df = pbt_bs_df.reindex(columns=all_columns)

# ===============================
# FUNCTION TO WRITE WITH SHEET SPLIT
# ===============================

def write_with_sheet_split(df, file_path, base_sheet_name):
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        total_rows = len(df)

        if total_rows == 0:
            # Still create empty sheet
            df.to_excel(writer, sheet_name=base_sheet_name, index=False)
            return

        sheet_count = 1
        start_row = 0

        while start_row < total_rows:
            end_row = start_row + MAX_ROWS
            chunk = df.iloc[start_row:end_row]

            if sheet_count == 1:
                sheet_name = base_sheet_name
            else:
                sheet_name = f"{base_sheet_name}_{sheet_count}"

            chunk.to_excel(writer, sheet_name=sheet_name, index=False)

            start_row += MAX_ROWS
            sheet_count += 1


# ===============================
# WRITE 3 SEPARATE FILES
# ===============================

write_with_sheet_split(
    rwa_df,
    os.path.join(OUTPUT_FOLDER, "RWA.xlsx"),
    "RWA"
)

write_with_sheet_split(
    avbs_df,
    os.path.join(OUTPUT_FOLDER, "AVBS_SD.xlsx"),
    "AVBS_SD"
)

write_with_sheet_split(
    pbt_bs_df,
    os.path.join(OUTPUT_FOLDER, "PBT_BS.xlsx"),
    "PBT_BS"
)

print("✅ All files created successfully.")
