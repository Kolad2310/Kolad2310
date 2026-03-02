```
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import sqlite3
from datetime import datetime

# -----------------------------
# FILE STORAGE
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
# FAST HEADER DETECTION (Rows 6–13 only)
# -----------------------------
def detect_header_fast(df):
    for i in range(5, min(13, len(df))):  # rows 6–13
        row = df.iloc[i].astype(str).str.lower().str.strip().tolist()
        if (
            ("year" in row and "entity" in row and "currency" in row)
            or
            ("mi_year" in row and "mi_entity" in row and "mi_currency" in row)
        ):
            return i
    return None

# -----------------------------
# LOAD DATA INTO SQLITE
# -----------------------------
def load_category(conn, file_list, table_name, progress_bar, step):
    for file in file_list:
        xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:
            preview = pd.read_excel(file, sheet_name=sheet, header=None, nrows=15)
            header_row = detect_header_fast(preview)

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
# SAFE UNION
# -----------------------------
def safe_union(conn, tables):
    existing = []
    cursor = conn.cursor()

    for t in tables:
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (t,))
        if cursor.fetchone():
            existing.append(t)

    if not existing:
        return pd.DataFrame()

    union_query = " UNION ALL ".join([f"SELECT * FROM {t}" for t in existing])
    return pd.read_sql_query(union_query, conn)

# -----------------------------
# RECONCILIATION
# -----------------------------
def create_reconciliation(df, category, dtype):
    if df.empty:
        return pd.DataFrame()

    month_cols = [c for c in df.columns if c.upper().startswith(("M", "YTD"))]

    recon = []
    for col in month_cols:
        temp = df.groupby(
            ["Source_File", "Source_Sheet", "year"],
            dropna=False
        )[col].sum().reset_index()

        temp["Month_Column"] = col
        temp["Category"] = category
        temp["Type"] = dtype
        temp.rename(columns={col: "Value"}, inplace=True)
        recon.append(temp)

    return pd.concat(recon, ignore_index=True) if recon else pd.DataFrame()

# -----------------------------
# MAIN PROCESS
# -----------------------------
def start_processing():
    root.destroy()

    progress_window = tk.Tk()
    progress_window.title("Processing")
    progress_window.geometry("500x150")

    tk.Label(progress_window, text="SQL Fast Consolidation Running...").pack(pady=10)

    progress_bar = ttk.Progressbar(progress_window, length=400, mode="determinate")
    progress_bar.pack(pady=10)

    total_files = sum(len(v) for v in file_store.values())
    if total_files == 0:
        messagebox.showerror("Error", "No files selected!")
        return

    step = 100 / total_files
    conn = sqlite3.connect(":memory:")

    # Load All
    for key in file_store:
        load_category(conn, file_store[key], key, progress_bar, step)

    # Consolidations
    rwa_actual = safe_union(conn, ["RWA_Actuals"])
    rwa_plan = safe_union(conn, ["RWA_Plan"])

    pbt_bs_actual = safe_union(conn, ["PBT_Actuals", "BS_Actuals"])
    pbt_bs_plan = safe_union(conn, ["PBT_Plan", "BS_Plan"])

    avbs_sd_actual = safe_union(conn, ["AVBS_Actuals", "SD_Actuals"])
    avbs_sd_plan = safe_union(conn, ["AVBS_Plan", "SD_Plan"])

    # Reconciliation
    recon_frames = [
        create_reconciliation(rwa_actual, "RWA", "Actual"),
        create_reconciliation(rwa_plan, "RWA", "Plan"),
        create_reconciliation(pbt_bs_actual, "PBT_BS", "Actual"),
        create_reconciliation(pbt_bs_plan, "PBT_BS", "Plan"),
        create_reconciliation(avbs_sd_actual, "AVBS_SD", "Actual"),
        create_reconciliation(avbs_sd_plan, "AVBS_SD", "Plan"),
    ]

    reconciliation_df = pd.concat(recon_frames, ignore_index=True)

    # Write Output
    output_name = f"Fast_SQL_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    with pd.ExcelWriter(output_name, engine="openpyxl") as writer:
        rwa_actual.to_excel(writer, "RWA Actuals", index=False)
        rwa_plan.to_excel(writer, "RWA Plan", index=False)
        pbt_bs_actual.to_excel(writer, "PBT_BS Actuals", index=False)
        pbt_bs_plan.to_excel(writer, "PBT_BS Plan", index=False)
        avbs_sd_actual.to_excel(writer, "AVBS_SD Actuals", index=False)
        avbs_sd_plan.to_excel(writer, "AVBS_SD Plan", index=False)
        reconciliation_df.to_excel(writer, "Reconciliation", index=False)

    progress_bar["value"] = 100
    tk.Label(progress_window, text="Completed Successfully!", fg="green").pack(pady=10)

# -----------------------------
# GUI
# -----------------------------
root = tk.Tk()
root.title("High-Speed SQL Consolidation Tool")
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
