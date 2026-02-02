```
import os
import psutil
import pandas as pd
from openpyxl import load_workbook
from concurrent.futures import ProcessPoolExecutor, as_completed
import pyarrow as pa
import pyarrow.parquet as pq
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog, messagebox


# ================= PERFORMANCE CONFIG =================
MAX_WORKERS = max(1, os.cpu_count() - 1)
ROWS_PER_MB = 450
MIN_CHUNK = 50_000
MAX_CHUNK = 300_000
# =====================================================


def auto_chunk_size():
    available_mb = psutil.virtual_memory().available // (1024 * 1024)
    chunk = int(available_mb * ROWS_PER_MB * 0.25)
    return max(MIN_CHUNK, min(chunk, MAX_CHUNK))


def read_excel_chunk(file_path, sheet_name, start_row, end_row):
    wb = load_workbook(
        file_path,
        read_only=True,
        data_only=True  # VALUES ONLY (no formulas)
    )
    ws = wb[sheet_name]

    return list(
        ws.iter_rows(
            min_row=start_row,
            max_row=end_row,
            values_only=True
        )
    )


def infer_schema(sample_rows, headers):
    df = pd.DataFrame(sample_rows, columns=headers)
    return pa.Schema.from_pandas(df, preserve_index=False)


def excel_to_parquet_fast(excel_file, sheet_name, output_file):
    chunk_size = auto_chunk_size()
    print(f"🔧 Auto chunk size: {chunk_size:,}")

    wb = load_workbook(excel_file, read_only=True, data_only=True)
    ws = wb[sheet_name]

    total_rows = ws.max_row
    headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))

    row_ranges = [
        (start, min(start + chunk_size - 1, total_rows))
        for start in range(2, total_rows + 1, chunk_size)
    ]

    # Infer schema
    sample_rows = read_excel_chunk(
        excel_file,
        sheet_name,
        row_ranges[0][0],
        min(row_ranges[0][0] + 5_000, row_ranges[0][1])
    )
    schema = infer_schema(sample_rows, headers)

    writer = pq.ParquetWriter(
        output_file,
        schema,
        compression="snappy"
    )

    with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(
                read_excel_chunk,
                excel_file,
                sheet_name,
                start,
                end
            )
            for start, end in row_ranges
        ]

        for future in tqdm(
            as_completed(futures),
            total=len(futures),
            desc="📊 Processing Excel"
        ):
            rows = future.result()
            if not rows:
                continue

            df = pd.DataFrame(rows, columns=headers)
            table = pa.Table.from_pandas(
                df,
                schema=schema,
                preserve_index=False
            )
            writer.write_table(table)

    writer.close()
    print("✅ Conversion complete")


# ================= GUI =================
def launch_gui():
    root = tk.Tk()
    root.title("Excel → Lightweight Parquet Converter")
    root.geometry("520x300")
    root.resizable(False, False)

    excel_path = tk.StringVar()
    sheet_name = tk.StringVar()
    output_name = tk.StringVar(value="output.parquet")

    def browse_file():
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not path:
            return

        excel_path.set(path)

        wb = load_workbook(path, read_only=True)
        sheets = wb.sheetnames

        sheet_menu["menu"].delete(0, "end")
        for s in sheets:
            sheet_menu["menu"].add_command(
                label=s,
                command=lambda value=s: sheet_name.set(value)
            )

        sheet_name.set(sheets[0])

    def submit():
        if not excel_path.get():
            messagebox.showerror("Error", "Please select an Excel file")
            return

        if not sheet_name.get():
            messagebox.showerror("Error", "Please select a sheet")
            return

        if not output_name.get():
            messagebox.showerror("Error", "Please enter output file name")
            return

        root.destroy()

        excel_to_parquet_fast(
            excel_path.get(),
            sheet_name.get(),
            output_name.get()
        )

        messagebox.showinfo(
            "Success",
            f"Parquet file created:\n{output_name.get()}"
        )

    # ---- UI Layout ----
    tk.Label(root, text="Excel File:", anchor="w").place(x=20, y=30)
    tk.Entry(root, textvariable=excel_path, width=50).place(x=120, y=30)
    tk.Button(root, text="Browse", command=browse_file).place(x=430, y=26)

    tk.Label(root, text="Sheet Name:", anchor="w").place(x=20, y=80)
    sheet_menu = tk.OptionMenu(root, sheet_name, "")
    sheet_menu.place(x=120, y=75, width=200)

    tk.Label(root, text="Output File:", anchor="w").place(x=20, y=130)
    tk.Entry(root, textvariable=output_name, width=30).place(x=120, y=130)
    tk.Label(root, text=".parquet").place(x=360, y=130)

    tk.Button(
        root,
        text="Start Conversion",
        width=20,
        height=2,
        command=submit,
        bg="#4CAF50",
        fg="white"
    ).place(x=170, y=190)

    root.mainloop()


# ================= RUN =================
if __name__ == "__main__":
    launch_gui()
