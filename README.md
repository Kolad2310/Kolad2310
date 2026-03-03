```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import traceback
import xlsxwriter

LOG_FILE = "Processing_Log.txt"
HEADER_FILE = "Header_Diagnostics.xlsx"

EXCEL_MAX_ROWS = 1048576
DATA_ROWS_PER_SHEET = EXCEL_MAX_ROWS - 1

file_store = {
    "RWA_Actuals": [],
    "RWA_Plan": [],
    "SD_Actuals": [],
    "SD_Plan": [],
    "AVBS_Actuals": [],
    "AVBS_Plan": [],
    "PBT_Actuals": [],
    "PBT_Plan": [],
    "BS_Actuals": [],
    "BS_Plan": []
}

# =====================================================
# LOGGING
# =====================================================
def log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# =====================================================
# HEADER DETECTION (ROBUST)
# =====================================================
def detect_header(df):

    for i in range(0, min(60, len(df))):

        row = [
            str(v).strip().replace("\xa0", "").lower()
            for v in df.iloc[i]
        ]

        has_year = any("year" in r for r in row)
        has_entity = any("entity" in r for r in row)
        has_currency = any("currency" in r for r in row)

        if has_year and has_entity and has_currency:
            return i

    return None

# =====================================================
# CONSOLIDATE ONE CATEGORY
# =====================================================
def consolidate_category(category):

    log(f"--- Consolidating {category} ---")

    collected = []
    header_records = []

    for file in file_store[category]:

        log(f"Reading file: {file}")
        xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:

            preview = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None,
                nrows=60,
                dtype=object
            )

            header_row = detect_header(preview)

            if header_row is None:
                log(f"Header NOT detected → {file} | {sheet}")
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row,
                dtype=object
            )

            df.columns = df.columns.str.strip().str.lower()
            df = df.dropna(how="all")

            # Convert numerics where possible
            for col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="ignore")

            df["Source_File"] = os.path.basename(file)
            df["Source_Sheet"] = sheet

            collected.append(df)

            header_records.append({
                "Category": category,
                "File": os.path.basename(file),
                "Sheet": sheet,
                "Header_Row": header_row + 1,
                "Columns_Detected": ", ".join(df.columns)
            })

            log(f"{category} | {sheet} rows: {len(df)}")

    if collected:
        final_df = pd.concat(collected, ignore_index=True)
        log(f"{category} TOTAL rows: {len(final_df)}")
    else:
        final_df = pd.DataFrame()
        log(f"{category} TOTAL rows: 0")

    # ===============================
    # APPLY BUSINESS LOGIC AFTER CONSOLIDATION
    # ===============================
    if category == "PBT_Actuals" and not final_df.empty:

        log("Applying PBT ÷1000 (post consolidation)")

        if "year" in final_df.columns:

            year_index = final_df.columns.get_loc("year")
            cols_after_year = final_df.columns[year_index + 1:]

            for col in cols_after_year:
                if pd.api.types.is_numeric_dtype(final_df[col]):
                    final_df[col] = final_df[col] / 1000

    return final_df, header_records

# =====================================================
# WRITE EXCEL WITH SPLIT
# =====================================================
def write_excel_file(file_name, sheet_dict):

    workbook = xlsxwriter.Workbook(
        file_name,
        {'constant_memory': True}
    )

    for base_sheet_name, df in sheet_dict.items():

        log(f"Writing {base_sheet_name} ({len(df)} rows)")

        if df.empty:
            workbook.add_worksheet(base_sheet_name[:31])
            continue

        df = df.where(pd.notnull(df), None)

        total_rows = len(df)
        splits = (total_rows // DATA_ROWS_PER_SHEET) + 1

        for split_index in range(splits):

            start = split_index * DATA_ROWS_PER_SHEET
            end = min(start + DATA_ROWS_PER_SHEET, total_rows)

            if start >= total_rows:
                break

            sheet_name = (
                base_sheet_name
                if split_index == 0
                else f"{base_sheet_name}_{split_index+1}"
            )

            worksheet = workbook.add_worksheet(sheet_name[:31])

            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name)

            chunk = df.iloc[start:end]

            for row_num, row in enumerate(
                    chunk.itertuples(index=False),
                    start=1):
                worksheet.write_row(row_num, 0, list(row))

    workbook.close()
    log(f"{file_name} completed")

# =====================================================
# MAIN PROCESS
# =====================================================
def start_processing():

    try:
        open(LOG_FILE, "w", encoding="utf-8").close()
        log("Processing started")

        tables = {}
        all_headers = []

        # Consolidate each category fully first
        for key in file_store:
            df, headers = consolidate_category(key)
            tables[key] = df
            all_headers.extend(headers)

        # Write diagnostics with full column list
        pd.DataFrame(all_headers).to_excel(
            HEADER_FILE,
            index=False
        )
        log("Header diagnostics written with full column list")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # RWA
        write_excel_file(
            f"RWA_Output_{timestamp}.xlsx",
            {
                "RWA Actuals": tables["RWA_Actuals"],
                "RWA Plan": tables["RWA_Plan"]
            }
        )

        # AVBS_SD
        write_excel_file(
            f"AVBS_SD_Output_{timestamp}.xlsx",
            {
                "AVBS_SD Actuals":
                    pd.concat([tables["AVBS_Actuals"],
                               tables["SD_Actuals"]],
                              ignore_index=True),
                "AVBS_SD Plan":
                    pd.concat([tables["AVBS_Plan"],
                               tables["SD_Plan"]],
                              ignore_index=True)
            }
        )

        # PBT_BS
        write_excel_file(
            f"PBT_BS_Output_{timestamp}.xlsx",
            {
                "PBT_BS Actuals":
                    pd.concat([tables["PBT_Actuals"],
                               tables["BS_Actuals"]],
                              ignore_index=True),
                "PBT_BS Plan":
                    pd.concat([tables["PBT_Plan"],
                               tables["BS_Plan"]],
                              ignore_index=True)
            }
        )

        log("Processing completed successfully")

        messagebox.showinfo(
            "Success",
            "Completed.\n3 files created."
        )

    except Exception:
        log("ERROR OCCURRED")
        log(traceback.format_exc())
        messagebox.showerror(
            "Error",
            "Processing failed.\nCheck Processing_Log.txt"
        )

# =====================================================
# GUI
# =====================================================
root = tk.Tk()
root.title("3 File Consolidation Tool")
root.geometry("800x600")

tk.Label(root,
         text="Select Files",
         font=("Arial", 14, "bold")).pack(pady=15)

frame = tk.Frame(root)
frame.pack()

labels = {}

def select_files_gui(key):
    files = filedialog.askopenfilenames(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if files:
        file_store[key] = list(files)
        labels[key].config(text=f"{len(files)} files selected")
        log(f"{key}: {len(files)} files selected")

for i, key in enumerate(file_store.keys()):
    tk.Label(frame, text=key,
             width=22).grid(row=i, column=0)
    tk.Button(frame,
              text="Select Files",
              command=lambda k=key:
              select_files_gui(k)).grid(row=i, column=1)
    lbl = tk.Label(frame,
                   text="No files selected",
                   width=25)
    lbl.grid(row=i, column=2)
    labels[key] = lbl

tk.Button(root,
          text="Submit & Process",
          command=start_processing,
          bg="green",
          fg="white",
          font=("Arial", 12, "bold")).pack(pady=20)

root.mainloop()
