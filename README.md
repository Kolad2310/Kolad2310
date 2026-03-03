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
# FILE SELECTOR
# =====================================================
def select_files(key):
    files = filedialog.askopenfilenames(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if files:
        file_store[key] = list(files)
        labels[key].config(text=f"{len(files)} files selected")
        log(f"{key}: {len(files)} files selected")

# =====================================================
# HEADER DETECTION
# =====================================================
def detect_header(df):

    def norm(x):
        return str(x).lower().replace("_", "").replace(" ", "").strip()

    for i in range(0, min(60, len(df))):
        row = [norm(v) for v in df.iloc[i]]

        if (
            any("year" in v for v in row)
            and any("entity" in v for v in row)
            and any(x in v for v in row for x in ["currency", "curr", "ccy"])
        ):
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
                nrows=60
            )

            header_row = detect_header(preview)

            if header_row is None:
                log(f"Header not found → {file} | {sheet}")
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row
            )

            df.columns = df.columns.str.strip().str.lower()

            df["Source_File"] = os.path.basename(file)
            df["Source_Sheet"] = sheet

            # ONE TIME PBT FIX
            if category == "PBT_Actuals":
                numeric_cols = df.select_dtypes(include=["number"]).columns
                numeric_cols = [c for c in numeric_cols if c != "year"]
                df[numeric_cols] = df[numeric_cols] / 1000

            combined.append(df)

            header_records.append({
                "Category": category,
                "File": os.path.basename(file),
                "Sheet": sheet,
                "Header_Row": header_row + 1
            })

    if combined:
        return pd.concat(combined, ignore_index=True), header_records

    return pd.DataFrame(), header_records

# =====================================================
# MAIN PROCESS
# =====================================================
def start_processing():

    try:

        open(LOG_FILE, "w", encoding="utf-8").close()
        log("Processing started")

        status_window = tk.Toplevel(root)
        status_window.title("Processing")
        status_window.geometry("400x120")

        status_label = tk.Label(status_window, text="Loading...")
        status_label.pack(pady=20)

        tables = {}
        all_headers = []

        for key in file_store:
            status_label.config(text=f"Loading {key}...")
            status_label.update()

            df, headers = load_category(key)
            tables[key] = df
            all_headers.extend(headers)

            log(f"{key} rows: {len(df)}")

        # Save header diagnostics
        pd.DataFrame(all_headers).to_excel(
            HEADER_FILE,
            index=False
        )
        log("Header diagnostics saved")

        # =====================================================
        # CREATE 3 FINAL COMPONENTS
        # =====================================================
        rwa_df = pd.concat(
            [tables["RWA_Actuals"], tables["RWA_Plan"]],
            ignore_index=True
        )

        pbt_bs_df = pd.concat(
            [
                tables["PBT_Actuals"],
                tables["PBT_Plan"],
                tables["BS_Actuals"],
                tables["BS_Plan"]
            ],
            ignore_index=True
        )

        avbs_sd_df = pd.concat(
            [
                tables["AVBS_Actuals"],
                tables["AVBS_Plan"],
                tables["SD_Actuals"],
                tables["SD_Plan"]
            ],
            ignore_index=True
        )

        output_name = (
            f"3_Component_Output_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        workbook = xlsxwriter.Workbook(
            output_name,
            {'constant_memory': True}
        )

        final_sheets = {
            "RWA": rwa_df,
            "PBT_BS": pbt_bs_df,
            "AVBS_SD": avbs_sd_df
        }

        for sheet, df in final_sheets.items():

            log(f"Writing {sheet} → {len(df)} rows")

            # SAFE conversion
            df = df.astype(object)
            df = df.where(pd.notnull(df), None)

            for col in df.columns:
                df[col] = df[col].apply(lambda x: "" if x is None else str(x))

            worksheet = workbook.add_worksheet(sheet[:31])

            # headers
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name)

            # rows
            for row_num, row in enumerate(
                    df.itertuples(index=False),
                    start=1):
                worksheet.write_row(row_num, 0, list(row))

        workbook.close()

        log("Processing completed successfully")

        messagebox.showinfo(
            "Success",
            f"Completed.\n\nCreated:\n{output_name}\n{HEADER_FILE}\n{LOG_FILE}"
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
root.title("3 Component Consolidation Tool")
root.geometry("800x600")

tk.Label(root,
         text="Select Files",
         font=("Arial", 14, "bold")).pack(pady=15)

frame = tk.Frame(root)
frame.pack()

labels = {}

for i, key in enumerate(file_store.keys()):
    tk.Label(frame, text=key,
             width=22).grid(row=i, column=0)
    tk.Button(frame,
              text="Select Files",
              command=lambda k=key:
              select_files(k)).grid(row=i, column=1)
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
