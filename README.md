```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import traceback
from openpyxl import Workbook

LOG_FILE = "Processing_Log.txt"
HEADER_FILE = "Header_Diagnostics.xlsx"

EXCEL_MAX_ROWS = 1048576
DATA_ROWS_PER_SHEET = EXCEL_MAX_ROWS - 1

# ------------------------------------------------
# FILTER LISTS (EDIT THESE)
# ------------------------------------------------

entity_list = [
    "Entity1","Entity2","Entity3"
]

globalbusiness_list = [
    "Business1","Business2","Business3"
]

# ------------------------------------------------

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

# ------------------------------------------------
# Logging
# ------------------------------------------------

def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE,"a",encoding="utf-8") as f:
        f.write(line+"\n")

# ------------------------------------------------
# Header Detection
# ------------------------------------------------

def detect_header(df):

    for i in range(min(60,len(df))):

        row = [str(v).strip().replace("\xa0","").lower() for v in df.iloc[i]]

        if (
            any("year" in r for r in row) and
            any("currency" in r for r in row) and
            any("entity" in r for r in row)
        ):
            return i

    return None

# ------------------------------------------------
# Normalize text
# ------------------------------------------------

def normalize_series(s):

    return (
        s.astype(str)
        .str.replace("\xa0","",regex=False)
        .str.strip()
        .str.upper()
    )

# ------------------------------------------------
# Proper case headers
# ------------------------------------------------

def proper_case(cols):

    return [
        str(c).replace("_"," ").title()
        for c in cols
    ]

# ------------------------------------------------
# Consolidate Category
# ------------------------------------------------

def consolidate_category(category):

    collected = []
    header_info = []

    for file in file_store[category]:

        log(f"Reading {file}")

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
                log(f"Header not found {file} | {sheet}")
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row,
                dtype=object
            )

            df.columns = df.columns.str.strip().str.lower()
            df = df.dropna(how="all")

            for col in df.columns:
                df[col] = pd.to_numeric(df[col],errors="ignore")

            df["source_file"] = os.path.basename(file)
            df["source_sheet"] = sheet

            collected.append(df)

            header_info.append({
                "Category":category,
                "File":os.path.basename(file),
                "Sheet":sheet,
                "Header_Row":header_row+1,
                "Columns":",".join(df.columns)
            })

            log(f"{category} {sheet} rows {len(df)}")

    if not collected:
        return pd.DataFrame(), header_info

    df = pd.concat(collected,ignore_index=True)

    log(f"{category} rows before filter {len(df)}")

    # ------------------------------------------------
    # Normalize columns
    # ------------------------------------------------

    if "currency" in df.columns:
        df["currency"] = normalize_series(df["currency"])

    if "entity" in df.columns:
        df["entity"] = normalize_series(df["entity"])

    if "global business" in df.columns:
        df["global business"] = normalize_series(df["global business"])

    entity_clean = [e.strip().upper() for e in entity_list]
    gb_clean = [g.strip().upper() for g in globalbusiness_list]

    # ------------------------------------------------
    # Filtering
    # ------------------------------------------------

    if "currency" in df.columns:
        df = df[df["currency"]=="USD"]

    if "entity" in df.columns:
        df = df[df["entity"].isin(entity_clean)]

    if "global business" in df.columns:
        df = df[df["global business"].isin(gb_clean)]

    log(f"{category} rows after filter {len(df)}")

    # ------------------------------------------------
    # PBT logic
    # ------------------------------------------------

    if category=="PBT_Actuals" and "year" in df.columns:

        year_index = df.columns.get_loc("year")
        numeric_cols = df.columns[year_index+1:]

        for col in numeric_cols:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col]/1000

    return df, header_info

# ------------------------------------------------
# Excel Writer (openpyxl write-only mode)
# ------------------------------------------------

def write_excel(file_name, sheet_dict):

    wb = Workbook(write_only=True)

    for base_sheet, df in sheet_dict.items():

        if df.empty:
            ws = wb.create_sheet(base_sheet)
            continue

        headers = proper_case(df.columns)

        total_rows = len(df)
        splits = (total_rows // DATA_ROWS_PER_SHEET) + 1

        for split in range(splits):

            start = split * DATA_ROWS_PER_SHEET
            end = min(start + DATA_ROWS_PER_SHEET,total_rows)

            if start >= total_rows:
                break

            sheet_name = base_sheet if split==0 else f"{base_sheet}_{split+1}"

            ws = wb.create_sheet(sheet_name[:31])

            ws.append(headers)

            chunk = df.iloc[start:end]

            for row in chunk.itertuples(index=False):

                clean_row = [
                    None if pd.isna(v) or v in [float("inf"),float("-inf")] else v
                    for v in row
                ]

                ws.append(clean_row)

    wb.save(file_name)

    log(f"{file_name} written")

# ------------------------------------------------
# Main Process
# ------------------------------------------------

def start_processing():

    try:

        open(LOG_FILE,"w",encoding="utf-8").close()

        log("Processing started")

        tables = {}
        headers = []

        for key in file_store:

            df,h = consolidate_category(key)

            tables[key]=df
            headers.extend(h)

        pd.DataFrame(headers).to_excel(HEADER_FILE,index=False)

        log("Header diagnostics written")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        # RWA

        write_excel(
            f"RWA_Output_{ts}.xlsx",
            {
                "RWA Actuals":tables["RWA_Actuals"],
                "RWA Plan":tables["RWA_Plan"]
            }
        )

        # AVBS_SD

        write_excel(
            f"AVBS_SD_Output_{ts}.xlsx",
            {
                "AVBS_SD Actuals":
                    pd.concat(
                        [tables["AVBS_Actuals"],tables["SD_Actuals"]],
                        ignore_index=True
                    ),

                "AVBS_SD Plan":
                    pd.concat(
                        [tables["AVBS_Plan"],tables["SD_Plan"]],
                        ignore_index=True
                    )
            }
        )

        # PBT_BS

        write_excel(
            f"PBT_BS_Output_{ts}.xlsx",
            {
                "PBT_BS Actuals":
                    pd.concat(
                        [tables["PBT_Actuals"],tables["BS_Actuals"]],
                        ignore_index=True
                    ),

                "PBT_BS Plan":
                    pd.concat(
                        [tables["PBT_Plan"],tables["BS_Plan"]],
                        ignore_index=True
                    )
            }
        )

        log("Processing completed")

        messagebox.showinfo(
            "Success",
            "Processing completed successfully"
        )

    except Exception:

        log("ERROR OCCURRED")
        log(traceback.format_exc())

        messagebox.showerror(
            "Error",
            "Processing failed. Check Processing_Log.txt"
        )

# ------------------------------------------------
# GUI
# ------------------------------------------------

root = tk.Tk()
root.title("Financial Consolidation Tool")
root.geometry("800x600")

tk.Label(
    root,
    text="Select Files",
    font=("Arial",14,"bold")
).pack(pady=15)

frame = tk.Frame(root)
frame.pack()

labels={}

def select_files_gui(key):

    files = filedialog.askopenfilenames(
        filetypes=[("Excel files","*.xlsx *.xls")]
    )

    if files:
        file_store[key]=list(files)
        labels[key].config(text=f"{len(files)} files selected")
        log(f"{key} files selected")

for i,key in enumerate(file_store):

    tk.Label(frame,text=key,width=22).grid(row=i,column=0)

    tk.Button(
        frame,
        text="Select Files",
        command=lambda k=key:select_files_gui(k)
    ).grid(row=i,column=1)

    lbl=tk.Label(frame,text="No files selected",width=25)
    lbl.grid(row=i,column=2)

    labels[key]=lbl

tk.Button(
    root,
    text="Submit & Process",
    command=start_processing,
    bg="green",
    fg="white",
    font=("Arial",12,"bold")
).pack(pady=20)

root.mainloop()
