```
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import sqlite3
from datetime import datetime

# -----------------------------
# GLOBAL FILE STORE
# -----------------------------
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

# -----------------------------
# FILE SELECTOR
# -----------------------------
def select_files(key):
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if files:
        file_store[key] = list(files)
        labels[key].config(text=f"{len(files)} files selected")

# -----------------------------
# HEADER DETECTION (FAST)
# -----------------------------
def detect_header(df):
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.lower().str.strip().tolist()
        if (
            ("year" in row and "entity" in row and "currency" in row)
            or
            ("mi_year" in row and "mi_entity" in row and "mi_currency" in row)
        ):
            return i
    return None

# -----------------------------
# LOAD TO SQLITE
# -----------------------------
def load_category_to_sql(conn, file_list, table_name, progress_bar, step):
    cursor = conn.cursor()

    for file in file_list:
        xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:
            raw = pd.read_excel(file, sheet_name=sheet, header=None)
            header_row = detect_header(raw)

            if header_row is None:
                continue

            df = pd.read_excel(file, sheet_name=sheet, header=header_row)
            df.columns = df.columns.str.strip().str.lower()

            df.rename(columns={
                "mi_year": "year",
                "mi_entity": "entity",
                "mi_currency": "currency"
            }, inplace=True)

            df["Source_File"] = os.path.basename(file)
            df["Source_Sheet"] = sheet

            df.to_sql(table_name, conn, if_exists="append", index=False)

        progress_bar["value"] += step
        progress_bar.update()

# -----------------------------
# MAIN PROCESS
# -----------------------------
def start_processing():
    root.destroy()

    progress_window = tk.Tk()
    progress_window.title("Processing")
    progress_window.geometry("500x150")

    tk.Label(progress_window, text="Processing with SQL Engine...").pack(pady=10)

    progress_bar = ttk.Progressbar(progress_window, length=400, mode="determinate")
    progress_bar.pack(pady=10)

    total_files = sum(len(v) for v in file_store.values())
    if total_files == 0:
        messagebox.showerror("Error", "No files selected!")
        return

    step = 100 / total_files

    conn = sqlite3.connect(":memory:")

    # Load all categories
    for key in file_store:
        load_category_to_sql(conn, file_store[key], key, progress_bar, step)

    # -----------------------------
    # SQL CONSOLIDATION
    # -----------------------------
    queries = {
        "RWA Actuals": "SELECT * FROM RWA_Actuals",
        "RWA Plan": "SELECT * FROM RWA_Plan",
        "PBT_BS Actuals": """
            SELECT * FROM PBT_Actuals
            UNION ALL
            SELECT * FROM BS_Actuals
        """,
        "PBT_BS Plan": """
            SELECT * FROM PBT_Plan
            UNION ALL
            SELECT * FROM BS_Plan
        """,
        "AVBS_SD Actuals": """
            SELECT * FROM AVBS_Actuals
            UNION ALL
            SELECT * FROM SD_Actuals
        """,
        "AVBS_SD Plan": """
            SELECT * FROM AVBS_Plan
            UNION ALL
            SELECT * FROM SD_Plan
        """
    }

    output_name = f"Consolidated_SQL_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    with pd.ExcelWriter(output_name, engine="openpyxl") as writer:

        consolidated_tables = {}

        for sheet_name, query in queries.items():
            try:
                df = pd.read_sql_query(query, conn)
            except:
                df = pd.DataFrame()

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            consolidated_tables[sheet_name] = df

        # -----------------------------
        # RECONCILIATION (SQL GROUP BY)
        # -----------------------------
        recon_query = """
        SELECT 
            Source_File,
            Source_Sheet,
            year,
            COUNT(*) as Row_Count
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
            recon_df = pd.read_sql_query(recon_query, conn)
        except:
            recon_df = pd.DataFrame()

        recon_df.to_excel(writer, sheet_name="Reconciliation", index=False)

    progress_bar["value"] = 100
    tk.Label(progress_window, text="Completed Successfully!", fg="green").pack(pady=10)

# -----------------------------
# GUI
# -----------------------------
root = tk.Tk()
root.title("SQL-Based Financial Consolidation")
root.geometry("750x550")

tk.Label(root, text="Select Files",
         font=("Arial", 14, "bold")).pack(pady=15)

frame = tk.Frame(root)
frame.pack()

labels = {}

categories = [
    ("RWA - Actuals", "RWA_Actuals"),
    ("RWA - Plan", "RWA_Plan"),
    ("SD - Actuals", "SD_Actuals"),
    ("SD - Plan", "SD_Plan"),
    ("AVBS - Actuals", "AVBS_Actuals"),
    ("AVBS - Plan", "AVBS_Plan"),
    ("PBT - Actuals", "PBT_Actuals"),
    ("PBT - Plan", "PBT_Plan"),
    ("BS - Actuals", "BS_Actuals"),
    ("BS - Plan", "BS_Plan"),
]

for i, (text, key) in enumerate(categories):
    tk.Label(frame, text=text, width=20, anchor="w").grid(row=i, column=0)
    tk.Button(frame, text="Select Files",
              command=lambda k=key: select_files(k)).grid(row=i, column=1)
    lbl = tk.Label(frame, text="No files selected", width=25)
    lbl.grid(row=i, column=2)
    labels[key] = lbl

tk.Button(root, text="Submit & Process",
          command=start_processing,
          bg="green", fg="white",
          font=("Arial", 12, "bold")).pack(pady=20)

root.mainloop()
