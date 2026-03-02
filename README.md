```
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import duckdb
import os
from datetime import datetime
import pyarrow as pa
import pyarrow.parquet as pq

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

TEMP_DIR = "temp_parquet"
os.makedirs(TEMP_DIR, exist_ok=True)


def select_files(key):
    files = filedialog.askopenfilenames(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if files:
        file_store[key] = list(files)
        labels[key].config(text=f"{len(files)} files selected")


def detect_header_robust(df):

    def norm(x):
        return str(x).lower().replace("_", "").replace(" ", "").strip()

    for i in range(0, min(40, len(df))):
        row = [norm(v) for v in df.iloc[i]]
        if (
            any("year" in v for v in row)
            and any("entity" in v for v in row)
            and any(x in v for v in row for x in ["currency", "curr", "ccy"])
        ):
            return i

    return None


def stream_excel_to_parquet(file, category):

    parquet_files = []

    xls = pd.ExcelFile(file)

    for sheet in xls.sheet_names:

        preview = pd.read_excel(file, sheet_name=sheet,
                                header=None, nrows=40)

        header_row = detect_header_robust(preview)
        if header_row is None:
            continue

        df = pd.read_excel(file, sheet_name=sheet,
                           header=header_row)

        df.columns = df.columns.str.strip().str.lower()

        df["Source_File"] = os.path.basename(file)
        df["Source_Sheet"] = sheet

        # One-time PBT scaling
        if category == "PBT_Actuals":
            numeric_cols = df.select_dtypes(
                include=["number"]
            ).columns
            numeric_cols = [
                c for c in numeric_cols
                if c.lower() != "year"
            ]
            df[numeric_cols] = df[numeric_cols] / 1000

        parquet_path = os.path.join(
            TEMP_DIR,
            f"{category}_{os.path.basename(file)}_{sheet}.parquet"
        )

        table = pa.Table.from_pandas(df)
        pq.write_table(table, parquet_path)

        parquet_files.append(parquet_path)

    return parquet_files


def start_processing():

    root.destroy()

    progress_window = tk.Tk()
    progress_window.title("Ultra Fast Streaming Mode")
    progress_window.geometry("600x200")

    progress_bar = ttk.Progressbar(progress_window,
                                   length=500)
    progress_bar.pack(pady=10)

    status_label = tk.Label(progress_window, text="")
    status_label.pack()

    total_files = sum(len(v) for v in file_store.values())
    if total_files == 0:
        messagebox.showerror("Error", "No files selected!")
        return

    step = 100 / total_files

    parquet_map = {}

    # Stream all Excel to parquet
    for category, files in file_store.items():
        parquet_map[category] = []
        for file in files:

            status_label.config(
                text=f"Streaming {category} → {os.path.basename(file)}"
            )
            progress_bar.update()

            parquet_files = stream_excel_to_parquet(file, category)
            parquet_map[category].extend(parquet_files)

            progress_bar["value"] += step
            progress_bar.update()

    status_label.config(text="Running DuckDB Queries...")
    progress_bar.update()

    con = duckdb.connect()
    con.execute("PRAGMA threads=8")

    output_name = (
        f"UltraFast_Output_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    with pd.ExcelWriter(output_name,
                        engine="openpyxl") as writer:

        queries = {
            "RWA Actuals":
                f"SELECT * FROM parquet_scan({parquet_map['RWA_Actuals']})",
            "RWA Plan":
                f"SELECT * FROM parquet_scan({parquet_map['RWA_Plan']})",
            "PBT_BS Actuals":
                f"""
                SELECT * FROM parquet_scan({parquet_map['PBT_Actuals']})
                UNION ALL
                SELECT * FROM parquet_scan({parquet_map['BS_Actuals']})
                """,
            "PBT_BS Plan":
                f"""
                SELECT * FROM parquet_scan({parquet_map['PBT_Plan']})
                UNION ALL
                SELECT * FROM parquet_scan({parquet_map['BS_Plan']})
                """,
            "AVBS_SD Actuals":
                f"""
                SELECT * FROM parquet_scan({parquet_map['AVBS_Actuals']})
                UNION ALL
                SELECT * FROM parquet_scan({parquet_map['SD_Actuals']})
                """,
            "AVBS_SD Plan":
                f"""
                SELECT * FROM parquet_scan({parquet_map['AVBS_Plan']})
                UNION ALL
                SELECT * FROM parquet_scan({parquet_map['SD_Plan']})
                """
        }

        for sheet, query in queries.items():
            try:
                df = con.execute(query).df()
            except:
                df = pd.DataFrame()
            df.to_excel(writer,
                        sheet_name=sheet,
                        index=False)

    progress_bar["value"] = 100
    status_label.config(
        text=f"Completed! File: {output_name}"
    )


root = tk.Tk()
root.title("Ultra Fast 1GB Streaming Tool")
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
