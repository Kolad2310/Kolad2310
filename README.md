```
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import duckdb
import os
from datetime import datetime

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
# FAST HEADER DETECTION
# =====================================================
def detect_header_fast(df):
    for i in range(5, min(13, len(df))):
        row = df.iloc[i].astype(str).str.lower().str.strip().tolist()
        if (
            ("year" in row and "entity" in row and "currency" in row)
            or
            ("mi_year" in row and "mi_entity" in row and "mi_currency" in row)
        ):
            return i
    return None

# =====================================================
# LOAD CATEGORY (FIXED SAFE VERSION)
# =====================================================
def load_category_fast(file_list, table_name,
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
                nrows=15
            )

            header_row = detect_header_fast(preview)

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

            # =====================================================
            # ONE TIME ADJUSTMENT (DELETE LATER)
            # Divide PBT Actuals by 1000
            # =====================================================
            if table_name == "PBT_Actuals":
                numeric_cols = df.select_dtypes(
                    include=["number"]
                ).columns
                numeric_cols = [
                    c for c in numeric_cols
                    if c.lower() != "year"
                ]
                df[numeric_cols] = df[numeric_cols] / 1000
            # =====================================================

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

    root.destroy()

    progress_window = tk.Tk()
    progress_window.title("Ultra Fast Processing")
    progress_window.geometry("650x220")

    tk.Label(
        progress_window,
        text="DuckDB Processing Running...",
        font=("Arial", 11, "bold")
    ).pack(pady=10)

    progress_bar = ttk.Progressbar(progress_window, length=600)
    progress_bar.pack(pady=10)

    status_label = tk.Label(progress_window, text="")
    status_label.pack()

    total_files = sum(len(v) for v in file_store.values())
    if total_files == 0:
        messagebox.showerror("Error", "No files selected!")
        return

    step = 100 / total_files

    # LOAD ALL CATEGORIES SAFELY
    tables = {}

    for key in file_store:
        tables[key] = load_category_fast(
            file_store[key],
            key,
            progress_bar,
            step,
            status_label
        )

    status_label.config(text="Registering Tables in DuckDB...")
    progress_bar.update()

    con = duckdb.connect(database=":memory:")
    con.execute("PRAGMA threads=8")

    # Register only non-empty tables
    for name, df in tables.items():
        if not df.empty:
            con.register(name, df)

    # CONSOLIDATION
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

    output_name = (
        f"Fixed_DuckDB_Output_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    with pd.ExcelWriter(output_name, engine="openpyxl") as writer:

        for sheet, query in queries.items():
            try:
                df = con.execute(query).df()
            except:
                df = pd.DataFrame()

            df.to_excel(writer, sheet_name=sheet, index=False)

        # Reconciliation
        recon_query = """
        SELECT Source_File, Source_Sheet, year, COUNT(*) AS rows_loaded
        FROM (
            SELECT * FROM RWA_Actuals
            UNION ALL SELECT * FROM RWA_Plan
            UNION ALL SELECT * FROM PBT_Actuals
            UNION ALL SELECT * FROM BS_Actuals
            UNION ALL SELECT * FROM AVBS_Actuals
            UNION ALL SELECT * FROM SD_Actuals
            UNION ALL SELECT * FROM AVBS_Plan
            UNION ALL SELECT * FROM SD_Plan
        )
        GROUP BY Source_File, Source_Sheet, year
        """

        try:
            recon_df = con.execute(recon_query).df()
        except:
            recon_df = pd.DataFrame()

        recon_df.to_excel(writer,
                          sheet_name="Reconciliation",
                          index=False)

    progress_bar["value"] = 100
    status_label.config(
        text=f"Completed! File Created: {output_name}"
    )

# =====================================================
# GUI
# =====================================================
root = tk.Tk()
root.title("Ultra Fast 1GB Excel Consolidation Tool")
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
