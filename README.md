```
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime

# -----------------------------
# GLOBAL FILE STORAGE
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
# HEADER DETECTION
# -----------------------------
def find_header_row(file_path, sheet):
    temp_df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    for i in range(len(temp_df)):
        row_values = (
            temp_df.iloc[i]
            .astype(str)
            .str.strip()
            .str.lower()
            .tolist()
        )

        standard_match = (
            "year" in row_values and
            "entity" in row_values and
            "currency" in row_values
        )

        mi_match = (
            "mi_year" in row_values and
            "mi_entity" in row_values and
            "mi_currency" in row_values
        )

        if standard_match or mi_match:
            return i

    return None

# -----------------------------
# MONTH MAPPER
# -----------------------------
def month_mapper(col_name):
    col_name = str(col_name).upper().strip()

    month_map = {
        "01": "Jan", "02": "Feb", "03": "Mar",
        "04": "Apr", "05": "May", "06": "Jun",
        "07": "Jul", "08": "Aug", "09": "Sep",
        "10": "Oct", "11": "Nov", "12": "Dec"
    }

    if col_name.startswith("M") and len(col_name) == 3:
        return month_map.get(col_name[1:], None)

    if col_name.startswith("YTD") and len(col_name) == 5:
        return month_map.get(col_name[3:], None)

    return None

# -----------------------------
# CONSOLIDATION FUNCTION
# -----------------------------
def consolidate_files(file_list, progress_bar, status_label, progress_step):
    final_df = pd.DataFrame()

    for file in file_list:
        wb = load_workbook(file, data_only=True)

        for sheet in wb.sheetnames:
            header_row = find_header_row(file, sheet)

            if header_row is not None:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row)

                df.columns = df.columns.str.strip().str.lower()

                df.rename(columns={
                    "mi_year": "year",
                    "mi_entity": "entity",
                    "mi_currency": "currency"
                }, inplace=True)

                df["Source_File"] = os.path.basename(file)
                df["Source_Sheet"] = sheet

                final_df = pd.concat([final_df, df], ignore_index=True)

        progress_bar["value"] += progress_step
        status_label.config(text=f"Processing {os.path.basename(file)}")
        progress_bar.update()

    return final_df

# -----------------------------
# RECONCILIATION
# -----------------------------
def create_reconciliation(df, category, data_type):
    recon_rows = []

    if df.empty:
        return pd.DataFrame()

    for col in df.columns:
        month_name = month_mapper(col)

        if month_name:
            temp = df.groupby(
                ["Source_File", "Source_Sheet", "year"],
                dropna=False
            )[col].sum().reset_index()

            temp["Month"] = month_name
            temp["Category"] = category
            temp["Type"] = data_type
            temp.rename(columns={col: "Value"}, inplace=True)

            recon_rows.append(temp)

    if recon_rows:
        return pd.concat(recon_rows, ignore_index=True)

    return pd.DataFrame()

# -----------------------------
# MAIN PROCESS
# -----------------------------
def start_processing():
    root.destroy()

    progress_window = tk.Tk()
    progress_window.title("Processing Files")
    progress_window.geometry("500x150")

    tk.Label(progress_window, text="Creating Consolidated Excel File...").pack(pady=10)

    progress_bar = ttk.Progressbar(progress_window, length=400, mode="determinate")
    progress_bar.pack(pady=10)

    status_label = tk.Label(progress_window, text="")
    status_label.pack()

    total_files = sum(len(v) for v in file_store.values())

    if total_files == 0:
        messagebox.showerror("Error", "No files selected!")
        progress_window.destroy()
        return

    progress_step = 100 / total_files

    # Individual Consolidations
    rwa_actual = consolidate_files(file_store["RWA_Actuals"], progress_bar, status_label, progress_step)
    rwa_plan = consolidate_files(file_store["RWA_Plan"], progress_bar, status_label, progress_step)

    pbt_actual = consolidate_files(file_store["PBT_Actuals"], progress_bar, status_label, progress_step)
    bs_actual = consolidate_files(file_store["BS_Actuals"], progress_bar, status_label, progress_step)
    pbt_bs_actual = pd.concat([pbt_actual, bs_actual], ignore_index=True)

    pbt_plan = consolidate_files(file_store["PBT_Plan"], progress_bar, status_label, progress_step)
    bs_plan = consolidate_files(file_store["BS_Plan"], progress_bar, status_label, progress_step)
    pbt_bs_plan = pd.concat([pbt_plan, bs_plan], ignore_index=True)

    sd_actual = consolidate_files(file_store["SD_Actuals"], progress_bar, status_label, progress_step)
    avbs_actual = consolidate_files(file_store["AVBS_Actuals"], progress_bar, status_label, progress_step)
    avbs_sd_actual = pd.concat([sd_actual, avbs_actual], ignore_index=True)

    sd_plan = consolidate_files(file_store["SD_Plan"], progress_bar, status_label, progress_step)
    avbs_plan = consolidate_files(file_store["AVBS_Plan"], progress_bar, status_label, progress_step)
    avbs_sd_plan = pd.concat([sd_plan, avbs_plan], ignore_index=True)

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

    # Output
    output_name = f"Consolidated_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    with pd.ExcelWriter(output_name, engine="openpyxl") as writer:
        rwa_actual.to_excel(writer, sheet_name="RWA Actuals", index=False)
        rwa_plan.to_excel(writer, sheet_name="RWA Plan", index=False)
        pbt_bs_actual.to_excel(writer, sheet_name="PBT_BS Actuals", index=False)
        pbt_bs_plan.to_excel(writer, sheet_name="PBT_BS Plan", index=False)
        avbs_sd_actual.to_excel(writer, sheet_name="AVBS_SD Actuals", index=False)
        avbs_sd_plan.to_excel(writer, sheet_name="AVBS_SD Plan", index=False)
        reconciliation_df.to_excel(writer, sheet_name="Reconciliation", index=False)

    progress_bar["value"] = 100
    status_label.config(text="Completed Successfully!")

    tk.Label(progress_window, text=f"File Created: {output_name}", fg="green").pack(pady=10)

# -----------------------------
# GUI
# -----------------------------
root = tk.Tk()
root.title("Financial Consolidation Tool")
root.geometry("750x550")

tk.Label(root, text="Select Files for Consolidation",
         font=("Arial", 14, "bold")).pack(pady=15)

frame = tk.Frame(root)
frame.pack(pady=10)

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
    tk.Label(frame, text=text, width=20, anchor="w").grid(row=i, column=0, padx=10, pady=5)
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
