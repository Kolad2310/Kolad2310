```
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import duckdb
import os
from datetime import datetime
import traceback

# =====================================================
# FILE STORAGE
# =====================================================
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
# FILE SELECTOR
# =====================================================
def select_files(key):
    files = filedialog.askopenfilenames(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if files:
        file_store[key] = list(files)
        labels[key].config(text=f"{len(files)} files selected")

# =====================================================
# ROBUST HEADER DETECTION
# =====================================================
def detect_header_robust(df):

    def normalize(val):
        return (
            str(val)
            .lower()
            .replace("_", "")
            .replace(" ", "")
            .strip()
        )

    def is_header(row_values):
        normalized = [normalize(v) for v in row_values]

        has_year = any("year" in v for v in normalized)
        has_entity = any("entity" in v for v in normalized)
        has_currency = any(
            x in v for v in normalized
            for x in ["currency", "curr", "ccy"]
        )

        return has_year and has_entity and has_currency

    for i in range(0, min(40, len(df))):
        if is_header(df.iloc[i].tolist()):
            return i

    return None

# =====================================================
# EXPORT HEADER DIAGNOSTICS
# =====================================================
def export_all_headers():

    header_records = []

    for category, files in file_store.items():

        for file in files:

            try:
                xls = pd.ExcelFile(file)

                for sheet in xls.sheet_names:

                    preview = pd.read_excel(
                        file,
                        sheet_name=sheet,
                        header=None,
                        nrows=50
                    )

                    header_row = detect_header_robust(preview)

                    if header_row is None:
                        header_records.append({
                            "Category": category,
                            "File": os.path.basename(file),
                            "Sheet": sheet,
                            "Header_Row": "NOT FOUND",
                            "Columns": "",
                            "Error": "Header not detected"
                        })
                        continue

                    df = pd.read_excel(
                        file,
                        sheet_name=sheet,
                        header=header_row,
                        nrows=1
                    )

                    header_records.append({
                        "Category": category,
                        "File": os.path.basename(file),
                        "Sheet": sheet,
                        "Header_Row": header_row + 1,
                        "Columns": " | ".join(df.columns.astype(str)),
                        "Error": ""
                    })

            except Exception as e:
                header_records.append({
                    "Category": category,
                    "File": os.path.basename(file),
                    "Sheet": "ERROR",
                    "Header_Row": "",
                    "Columns": "",
                    "Error": str(e)
                })

    if header_records:
        pd.DataFrame(header_records).to_excel(
            "Header_Diagnostics.xlsx",
            index=False
        )

# =====================================================
# LOAD CATEGORY SAFELY
# =====================================================
def load_category(file_list, table_name,
                  progress_bar, step, status_label):

    all_dfs = []

    for file in file_list:

        status_label.config(
            text=f"Processing {table_name} → {os.path.basename(file)}"
        )
        progress_bar.update()

        xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:

            preview = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None,
                nrows=50
            )

            header_row = detect_header_robust(preview)

            if header_row is None:
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row
            )

            df.columns = df.columns.str.strip().str.lower()

            df.rename(columns={
                "mi_year": "year",
                "mi_entity": "entity",
                "mi_currency": "currency"
            }, inplace=True)

            df["Source_File"] = os.path.basename(file)
            df["Source_Sheet"] = sheet

            # =============================================
            # ONE TIME PBT ADJUSTMENT (DELETE LATER)
            # =============================================
            if table_name == "PBT_Actuals":
                numeric_cols = df.select_dtypes(
                    include=["number"]
                ).columns
                numeric_cols = [
                    c for c in numeric_cols
                    if c.lower() != "year"
                ]
                df[numeric_cols] = df[numeric_cols] / 1000
            # =============================================

            all_dfs.append(df)

        progress_bar["value"] += step
        progress_bar.update()

    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)

    return pd.DataFrame()

# =====================================================
# MAIN PROCESS
# =====================================================
def start_processing():

    try:

        export_all_headers()

        root.destroy()

        progress_window = tk.Tk()
        progress_window.title("Processing")
        progress_window.geometry("650x220")

        progress_bar = ttk.Progressbar(progress_window, length=600)
        progress_bar.pack(pady=10)

        status_label = tk.Label(progress_window, text="")
        status_label.pack()

        total_files = sum(len(v) for v in file_store.values())

        if total_files == 0:
            messagebox.showerror("Error", "No files selected!")
            return

        step = 100 / total_files

        tables = {}

        for key in file_store:
            tables[key] = load_category(
                file_store[key],
                key,
                progress_bar,
                step,
                status_label
            )

        # Align schemas
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

        con = duckdb.connect(database=":memory:")
        con.execute("PRAGMA threads=8")

        for name, df in tables.items():
            con.register(name, df)

        output_name = (
            f"Final_Output_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        with pd.ExcelWriter(output_name, engine="openpyxl") as writer:

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
                df.to_excel(writer, sheet_name=sheet, index=False)

        progress_bar["value"] = 100
        status_label.config(
            text=f"Completed! File Created: {output_name}"
        )

    except Exception as e:

        error_message = traceback.format_exc()

        with open("Error_Log.txt", "w") as f:
            f.write(error_message)

        messagebox.showerror(
            "Processing Failed",
            f"An error occurred.\nCheck Error_Log.txt"
        )

# =====================================================
# GUI
# =====================================================
root = tk.Tk()
root.title("1GB Consolidation Tool (Diagnostic Mode)")
root.geometry("800x600")

tk.Label(root,
         text="Select Files for Consolidation",
         font=("Arial", 14, "bold")).pack(pady=15)

frame = tk.Frame(root)
frame.pack()

labels = {}

for i, key in enumerate(file_store.keys()):
    tk.Label(frame, text=key, width=22,
             anchor="w").grid(row=i, column=0)
    tk.Button(frame,
              text="Select Files",
              command=lambda k=key:
              select_files(k)).grid(row=i, column=1)
    lbl = tk.Label(frame,
                   text="No files selected",
                   width=25)
    lbl.grid(row=i, column=2)
    labels[key] = lbl

tk.Button(
    root,
    text="Submit & Process",
    command=start_processing,
    bg="green",
    fg="white",
    font=("Arial", 12, "bold")
).pack(pady=20)

root.mainloop()
