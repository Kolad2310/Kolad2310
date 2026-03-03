```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import duckdb
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
# LOGGING (UTF-8 SAFE)
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
# MAIN PROCESS
# =====================================================
def start_processing():

    try:

        open(LOG_FILE, "w", encoding="utf-8").close()
        log("Processing started")

        status_window = tk.Toplevel(root)
        status_window.title("Processing Status")
        status_window.geometry("400x120")

        status_label = tk.Label(status_window, text="Starting...")
        status_label.pack(pady=20)

        tables = {k: [] for k in file_store}
        header_records = []

        # =====================================================
        # READ EACH FILE ONLY ONCE
        # =====================================================
        for category, files in file_store.items():

            log(f"Loading category: {category}")
            status_label.config(text=f"Loading {category}...")
            status_label.update()

            for file in files:

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
                        header_records.append({
                            "Category": category,
                            "File": os.path.basename(file),
                            "Sheet": sheet,
                            "Header_Row": "NOT FOUND"
                        })
                        continue

                    df = pd.read_excel(
                        file,
                        sheet_name=sheet,
                        header=header_row
                    )

                    df.columns = df.columns.str.strip().str.lower()

                    df["Source_File"] = os.path.basename(file)
                    df["Source_Sheet"] = sheet

                    # =====================================================
                    # ONE TIME PBT FIX (DELETE LATER)
                    # =====================================================
                    if category == "PBT_Actuals":
                        log("Applying PBT ÷1000 adjustment")
                        numeric_cols = df.select_dtypes(include=["number"]).columns
                        numeric_cols = [c for c in numeric_cols if c != "year"]
                        df[numeric_cols] = df[numeric_cols] / 1000
                    # =====================================================

                    tables[category].append(df)

                    header_records.append({
                        "Category": category,
                        "File": os.path.basename(file),
                        "Sheet": sheet,
                        "Header_Row": header_row + 1
                    })

        # =====================================================
        # SAVE HEADER DIAGNOSTICS
        # =====================================================
        pd.DataFrame(header_records).to_excel(
            HEADER_FILE,
            index=False
        )
        log("Header diagnostics created")

        # =====================================================
        # CONCAT ONCE PER CATEGORY
        # =====================================================
        for key in tables:
            if tables[key]:
                tables[key] = pd.concat(tables[key], ignore_index=True)
                log(f"{key} rows: {len(tables[key])}")
            else:
                tables[key] = pd.DataFrame()
                log(f"{key} rows: 0")

        # =====================================================
        # ALIGN SCHEMA
        # =====================================================
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

        # =====================================================
        # REGISTER IN DUCKDB
        # =====================================================
        con = duckdb.connect(database=":memory:")
        for name, df in tables.items():
            con.register(name, df)

        log("DuckDB registration completed")

        output_name = (
            f"Consolidated_Output_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        status_label.config(text="Writing Excel...")
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

            df = con.execute(query).df()
            log(f"Writing {sheet} → {len(df)} rows")

            # =====================================================
            # SAFE CONVERSION FOR EXCEL
            # =====================================================
            df = df.copy()
            df = df.astype(object)
            df = df.where(pd.notnull(df), None)

            for col in df.columns:
                df[col] = df[col].apply(lambda x: "" if x is None else str(x))
            # =====================================================

            worksheet = workbook.add_worksheet(sheet[:31])

            # Write headers
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name)

            # Write rows
            for row_num, row in enumerate(
                    df.itertuples(index=False),
                    start=1):
                worksheet.write_row(row_num, 0, list(row))

        workbook.close()

        log("Excel writing completed")
        log("Processing finished successfully")

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
root.title("Optimized Consolidation Tool")
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
