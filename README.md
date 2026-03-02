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
# HEADER DETECTION (ROBUST BUT SIMPLE)
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
# HEADER DIAGNOSTICS (ALWAYS RUN FIRST)
# =====================================================
def export_header_diagnostics():

    records = []

    for category, files in file_store.items():

        for file in files:

            try:
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
                        records.append({
                            "Category": category,
                            "File": os.path.basename(file),
                            "Sheet": sheet,
                            "Header_Row": "NOT FOUND",
                            "Columns": "",
                            "Status": "HEADER NOT DETECTED"
                        })
                        continue

                    df = pd.read_excel(
                        file,
                        sheet_name=sheet,
                        header=header_row,
                        nrows=1
                    )

                    records.append({
                        "Category": category,
                        "File": os.path.basename(file),
                        "Sheet": sheet,
                        "Header_Row": header_row + 1,
                        "Columns": " | ".join(df.columns.astype(str)),
                        "Status": "OK"
                    })

            except Exception as e:
                records.append({
                    "Category": category,
                    "File": os.path.basename(file),
                    "Sheet": "ERROR",
                    "Header_Row": "",
                    "Columns": "",
                    "Status": str(e)
                })

    df_diag = pd.DataFrame(records)
    df_diag.to_excel("Header_Diagnostics.xlsx", index=False)
    print("Header_Diagnostics.xlsx created")

# =====================================================
# LOAD CATEGORY
# =====================================================
def load_category(category):

    all_dfs = []

    for file in file_store[category]:

        print("Processing:", category, file)

        xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:

            preview = pd.read_excel(file, sheet_name=sheet,
                                    header=None, nrows=60)

            header_row = detect_header(preview)

            if header_row is None:
                print("Header NOT found:", file, sheet)
                continue

            df = pd.read_excel(file,
                               sheet_name=sheet,
                               header=header_row)

            df.columns = df.columns.str.strip().str.lower()

            df["Source_File"] = os.path.basename(file)
            df["Source_Sheet"] = sheet

            # ONE TIME PBT FIX
            if category == "PBT_Actuals":
                numeric_cols = df.select_dtypes(
                    include=["number"]
                ).columns
                numeric_cols = [c for c in numeric_cols if c != "year"]
                df[numeric_cols] = df[numeric_cols] / 1000

            all_dfs.append(df)

    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)

    return pd.DataFrame()

# =====================================================
# MAIN PROCESS
# =====================================================
def start_processing():

    try:

        export_header_diagnostics()

        tables = {}

        for key in file_store:
            tables[key] = load_category(key)

        # Save load summary
        summary = []
        for key, df in tables.items():
            summary.append({
                "Category": key,
                "Rows_Loaded": len(df)
            })

        pd.DataFrame(summary).to_excel(
            "Load_Summary.xlsx",
            index=False
        )

        print("Load_Summary.xlsx created")

        # Print row counts
        for key in tables:
            print(key, "rows:", len(tables[key]))

        # Align schema
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

        for name, df in tables.items():
            con.register(name, df)

        output_name = (
            f"Final_Output_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        with pd.ExcelWriter(output_name,
                            engine="openpyxl") as writer:

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
                df.to_excel(writer,
                            sheet_name=sheet,
                            index=False)

        messagebox.showinfo(
            "Success",
            f"Completed.\nCreated:\nHeader_Diagnostics.xlsx\nLoad_Summary.xlsx\n{output_name}"
        )

    except Exception as e:

        with open("Error_Log.txt", "w") as f:
            f.write(traceback.format_exc())

        messagebox.showerror(
            "Error",
            "Processing failed.\nCheck Error_Log.txt"
        )

# =====================================================
# GUI
# =====================================================
root = tk.Tk()
root.title("Diagnostic Consolidation Tool")
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
          fg="white").pack(pady=20)

root.mainloop()
