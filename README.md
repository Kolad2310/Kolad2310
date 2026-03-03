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
# HEADER DETECTION
# =====================================================
def detect_header(df):

    for i in range(0, min(60, len(df))):
        row = [str(v).strip().lower() for v in df.iloc[i]]

        has_year = any("year" == v for v in row)
        has_entity = any("entity" == v for v in row)
        has_currency = any("currency" == v for v in row)

        if has_year and has_entity and has_currency:
            return i

    return None

# =====================================================
# LOAD CATEGORY
# =====================================================
def load_category(category):

    combined = []
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
                log(f"Header NOT found → {file} | {sheet}")
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row,
                dtype=object
            )

            df.columns = df.columns.str.strip().str.lower()
            df = df.dropna(how="all")

            log(f"{category} | {sheet} rows loaded: {len(df)}")

            df["Source_File"] = os.path.basename(file)
            df["Source_Sheet"] = sheet

            # ===============================
            # FIXED PBT ÷1000 LOGIC
            # ===============================
            if category == "PBT_Actuals":

                if "year" in df.columns:

                    year_index = df.columns.get_loc("year")
                    cols_after_year = df.columns[year_index + 1:]

                    for col in cols_after_year:
                        df[col] = pd.to_numeric(df[col], errors="coerce")
                        if pd.api.types.is_numeric_dtype(df[col]):
                            df[col] = df[col] / 1000

                    log("PBT ÷1000 applied correctly")

            combined.append(df)

            header_records.append({
                "Category": category,
                "File": os.path.basename(file),
                "Sheet": sheet,
                "Header_Row": header_row + 1
            })

    if combined:
        final_df = pd.concat(combined, ignore_index=True)
        log(f"{category} TOTAL rows after concat: {len(final_df)}")
        return final_df, header_records

    log(f"{category} TOTAL rows after concat: 0")
    return pd.DataFrame(), header_records

# =====================================================
# WRITE WITH AUTO SPLIT
# =====================================================
def write_excel_file(file_name, sheet_dict):

    workbook = xlsxwriter.Workbook(
        file_name,
        {'constant_memory': True}
    )

    for base_sheet_name, df in sheet_dict.items():

        log(f"Preparing {base_sheet_name} ({len(df)} rows)")

        if df.empty:
            workbook.add_worksheet(base_sheet_name[:31])
            continue

        # Safe conversion
        df = df.astype(object)
        df = df.where(pd.notnull(df), None)

        for col in df.columns:
            df[col] = df[col].apply(lambda x: "" if x is None else str(x))

        total_rows = len(df)
        split_count = (total_rows // DATA_ROWS_PER_SHEET) + 1

        for split_index in range(split_count):

            start = split_index * DATA_ROWS_PER_SHEET
            end = min(start + DATA_ROWS_PER_SHEET, total_rows)

            if start >= total_rows:
                break

            sheet_name = (
                base_sheet_name
                if split_index == 0
                else f"{base_sheet_name}_{split_index+1}"
            )

            log(f"Writing {sheet_name} rows {start} to {end}")

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

        for key in file_store:
            df, headers = load_category(key)
            tables[key] = df
            all_headers.extend(headers)

        pd.DataFrame(all_headers).to_excel(
            HEADER_FILE,
            index=False
        )
        log("Header diagnostics saved")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # DEBUG COUNTS
        log(f"PBT_Actuals rows: {len(tables['PBT_Actuals'])}")
        log(f"BS_Actuals rows: {len(tables['BS_Actuals'])}")

        # RWA FILE
        write_excel_file(
            f"RWA_Output_{timestamp}.xlsx",
            {
                "RWA Actuals": tables["RWA_Actuals"],
                "RWA Plan": tables["RWA_Plan"]
            }
        )

        # AVBS_SD FILE
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

        # PBT_BS FILE
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
            "Completed.\n\n3 files created successfully."
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
