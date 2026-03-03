```
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import duckdb
import os
from datetime import datetime
import traceback
import xlsxwriter

# =====================================================
# GLOBALS
# =====================================================
LOG_FILE = "Processing_Log.txt"

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
# LOG FUNCTION
# =====================================================
def log(message):
    timestamp = datetime.now().strftime("%H:%M:%S")
    full_msg = f"[{timestamp}] {message}"
    print(full_msg)

    with open(LOG_FILE, "a") as f:
        f.write(full_msg + "\n")

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

        has_year = any("year" in v for v in row)
        has_entity = any("entity" in v for v in row)
        has_currency = any(x in v for v in row for x in ["currency", "curr", "ccy"])

        if has_year and has_entity and has_currency:
            return i

    return None

# =====================================================
# LOAD CATEGORY
# =====================================================
def load_category(category, status_label):

    log(f"Starting load for {category}")
    status_label.config(text=f"Loading {category}...")
    status_label.update()

    all_dfs = []

    for file in file_store[category]:

        log(f"Reading file: {file}")
        xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:

            log(f"Checking sheet: {sheet}")
            preview = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None,
                nrows=60
            )

            header_row = detect_header(preview)

            if header_row is None:
                log(f"Header NOT found in {file} | {sheet}")
                continue

            log(f"Header detected at row {header_row+1}")

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row
            )

            df.columns = df.columns.str.strip().str.lower()

            df["Source_File"] = os.path.basename(file)
            df["Source_Sheet"] = sheet

            # ==========================================
            # ONE TIME PBT FIX (DELETE LATER)
            # ==========================================
            if category == "PBT_Actuals":
                log("Applying PBT ÷1000 adjustment")
                numeric_cols = df.select_dtypes(include=["number"]).columns
                numeric_cols = [c for c in numeric_cols if c != "year"]
                df[numeric_cols] = df[numeric_cols] / 1000
            # ==========================================

            all_dfs.append(df)

    if all_dfs:
        combined = pd.concat(all_dfs, ignore_index=True)
        log(f"{category} loaded: {len(combined)} rows")
        return combined

    log(f"{category} loaded: 0 rows")
    return pd.DataFrame()

# =====================================================
# WRITE EXCEL (STREAMING)
# =====================================================
def write_excel_streaming(output_name, con, status_label):

    log("Starting Excel writing phase")
    status_label.config(text="Writing Excel file...")
    status_label.update()

    workbook = xlsxwriter.Workbook(
        output_name,
        {'constant_memory': True}
    )

    queries = {
        "RWA Actuals":
            "SELECT * FROM RWA_Actuals",
        "RWA Plan":
            "SELECT * FROM RWA_Plan",
        "PBT_BS Actuals":
            "SELECT * FROM PBT_Actuals UNION ALL SELECT * FROM BS_Actuals",
        "PBT_BS Plan":
            "SELECT * FROM PBT_Plan UNION ALL SELECT * FROM BS_Plan",
        "AVBS_SD Actuals":
            "SELECT * FROM AVBS_Actuals UNION ALL SELECT * FROM SD_Actuals",
        "AVBS_SD Plan":
            "SELECT * FROM AVBS_Plan UNION ALL SELECT * FROM SD_Plan"
    }

    for sheet, query in queries.items():

        log(f"Executing query for {sheet}")
        df = con.execute(query).df()
        log(f"{sheet}: {len(df)} rows to write")

        worksheet = workbook.add_worksheet(sheet[:31])

        # Write headers
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name)

        # Write rows
        for row_num, row in enumerate(
                df.itertuples(index=False),
                start=1):
            worksheet.write_row(row_num, 0, row)

        log(f"{sheet} writing completed")

    workbook.close()
    log("Excel file writing completed")

# =====================================================
# MAIN PROCESS
# =====================================================
def start_processing():

    try:

        # Clear old log
        open(LOG_FILE, "w").close()
        log("Processing started")

        progress_window = tk.Toplevel(root)
        progress_window.title("Processing")
        progress_window.geometry("600x150")

        status_label = tk.Label(progress_window, text="")
        status_label.pack(pady=20)

        tables = {}

        for key in file_store:
            tables[key] = load_category(key, status_label)

        log("All categories loaded")

        # Align schema
        log("Aligning schemas")
        all_columns = set()
        for df in tables.values():
            if not df.empty:
                all_columns.update(df.columns)

        all_columns = list(all_columns)

        for key in tables:
            if tables[key].empty:
                tables[key] = pd.DataFrame(columns=all_columns)
            else:
                for col in all_columns:
                    if col not in tables[key].columns:
                        tables[key][col] = None
                tables[key] = tables[key][all_columns]

        log("Schema alignment completed")

        # Register in DuckDB
        log("Registering tables in DuckDB")
        con = duckdb.connect(database=":memory:")
        for name, df in tables.items():
            con.register(name, df)

        log("DuckDB registration completed")

        output_name = (
            f"Final_Output_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        write_excel_streaming(output_name, con, status_label)

        log(f"File created: {output_name}")

        messagebox.showinfo(
            "Success",
            f"Completed.\n\nCreated:\n{output_name}\n\nSee Processing_Log.txt for steps."
        )

    except Exception as e:

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
root.title("0.5GB Optimized Consolidation Tool")
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
