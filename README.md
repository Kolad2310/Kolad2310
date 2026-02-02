```
import os
import psutil
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl import load_workbook
from concurrent.futures import ProcessPoolExecutor
import pyarrow as pa
import pyarrow.parquet as pq


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
    """
    Reads raw Excel values (no headers, values only)
    Preserves numeric vs string types
    """
    wb = load_workbook(
        file_path,
        read_only=True,
        data_only=True
    )
    ws = wb[sheet_name]

    return list(
        ws.iter_rows(
            min_row=start_row,
            max_row=end_row,
            values_only=True
        )
    )


def build_schema_and_columns(sample_rows):
    """
    Build positional column names: col_1, col_2, ...
    Infer schema so numbers stay numbers, strings stay strings
    """
    col_count = max(len(row) for row in sample_rows)
    columns = [f"col_{i+1}" for i in range(col_count)]

    df = pd.DataFrame(sample_rows, columns=columns)
    schema = pa.Schema.from_pandas(df, preserve_index=False)

    return schema, columns


# ================= CORE CONVERSION =================
def excel_to_parquet_with_progress(
    excel_file,
    sheet_name,
    output_file,
    progress_queue
):
    try:
        chunk_size = auto_chunk_size()

        wb = load_workbook(excel_file, read_only=True, data_only=True)
        ws = wb[sheet_name]

        total_rows = ws.max_row

        # Prepare row ranges (NO HEADER SKIP)
        row_ranges = [
            (start, min(start + chunk_size - 1, total_rows))
            for start in range(1, total_rows + 1, chunk_size)
        ]

        total_chunks = len(row_ranges)

        # Sample rows for schema inference
        sample_rows = read_excel_chunk(
            excel_file,
            sheet_name,
            row_ranges[0][0],
            min(row_ranges[0][0] + 5_000, row_ranges[0][1])
        )

        schema, columns = build_schema_and_columns(sample_rows)

        writer = pq.ParquetWriter(
            output_file,
            schema,
            compression="snappy"
        )

        processed = 0

        with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
            tasks = (
                (excel_file, sheet_name, start, end)
                for start, end in row_ranges
            )

            for rows in executor.map(lambda args: read_excel_chunk(*args), tasks):
                if rows:
                    df = pd.DataFrame(rows, columns=columns)
                    table = pa.Table.from_pandas(
                        df,
                        schema=schema,
                        preserve_index=False
                    )
                    writer.write_table(table)

                processed += 1
                progress_queue.put((processed, total_chunks))

        writer.close()
        progress_queue.put(("DONE", output_file))

    except Exception as e:
        progress_queue.put(("ERROR", str(e)))


# ================= GUI =================
def launch_gui():
    root = tk.Tk()
    root.title("Excel → Parquet (Raw Positional Data)")
    root.geometry("540x300")
    root.resizable(False, False)

    excel_path = tk.StringVar()
    sheet_name = tk.StringVar()
    output_name = tk.StringVar(value="output.parquet")

    progress_queue = queue.Queue()

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

    def start_conversion():
        if not excel_path.get() or not sheet_name.get():
            messagebox.showerror("Error", "Select Excel file and sheet")
            return

        root.withdraw()

        progress_window = tk.Toplevel()
        progress_window.title("Processing Excel")
        progress_window.geometry("420x150")
        progress_window.resizable(False, False)

        tk.Label(
            progress_window,
            text="Reading Excel (values only)\nSaving raw positional data",
            pady=10
        ).pack()

        progress_bar = ttk.Progressbar(
            progress_window,
            orient="horizontal",
            length=340,
            mode="determinate"
        )
        progress_bar.pack(pady=10)

        status_label = tk.Label(progress_window, text="Starting...")
        status_label.pack()

        def process_queue():
            try:
                msg = progress_queue.get_nowait()

                if msg[0] == "DONE":
                    progress_window.destroy()
                    messagebox.showinfo(
                        "Success",
                        f"Parquet file created:\n{msg[1]}"
                    )
                    root.destroy()
                    return

                if msg[0] == "ERROR":
                    progress_window.destroy()
                    messagebox.showerror("Error", msg[1])
                    root.destroy()
                    return

                processed, total = msg
                progress_bar["maximum"] = total
                progress_bar["value"] = processed
                status_label.config(
                    text=f"Processed {processed} of {total} chunks"
                )

            except queue.Empty:
                pass

            progress_window.after(200, process_queue)

        threading.Thread(
            target=excel_to_parquet_with_progress,
            args=(
                excel_path.get(),
                sheet_name.get(),
                output_name.get(),
                progress_queue
            ),
            daemon=True
        ).start()

        process_queue()

    # ---- Layout ----
    tk.Label(root, text="Excel File:").place(x=20, y=30)
    tk.Entry(root, textvariable=excel_path, width=45).place(x=140, y=30)
    tk.Button(root, text="Browse", command=browse_file).place(x=450, y=26)

    tk.Label(root, text="Sheet:").place(x=20, y=80)
    sheet_menu = tk.OptionMenu(root, sheet_name, "")
    sheet_menu.place(x=140, y=75, width=220)

    tk.Label(root, text="Output File:").place(x=20, y=130)
    tk.Entry(root, textvariable=output_name, width=32).place(x=140, y=130)
    tk.Label(root, text=".parquet").place(x=380, y=130)

    tk.Button(
        root,
        text="Start Conversion",
        width=22,
        height=2,
        bg="#4CAF50",
        fg="white",
        command=start_conversion
    ).place(x=180, y=190)

    root.mainloop()


# ================= RUN =================
if __name__ == "__main__":
    launch_gui()
