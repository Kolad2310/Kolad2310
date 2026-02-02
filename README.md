```
import os
import psutil
import threading
import queue
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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


# ---------- WORKER FUNCTIONS (TOP-LEVEL ONLY) ----------
def read_excel_chunk_worker(args):
    """
    Worker-safe function (pickleable)
    """
    file_path, sheet_name, start_row, end_row = args

    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb[sheet_name]

    return list(
        ws.iter_rows(
            min_row=start_row,
            max_row=end_row,
            values_only=True
        )
    )


def detect_datetime_columns(rows, col_count):
    datetime_cols = set()
    for row in rows:
        for i in range(min(len(row), col_count)):
            if isinstance(row[i], datetime.datetime):
                datetime_cols.add(i)
    return datetime_cols


def build_safe_schema(sample_rows):
    col_count = max(len(r) for r in sample_rows)
    columns = [f"col_{i+1}" for i in range(col_count)]

    datetime_cols = detect_datetime_columns(sample_rows, col_count)

    fields = []
    for i in range(col_count):
        if i in datetime_cols:
            fields.append(pa.field(columns[i], pa.timestamp("ns")))
        else:
            fields.append(pa.field(columns[i], pa.string()))

    schema = pa.schema(fields)
    return schema, columns, datetime_cols, col_count


def normalize_rows(rows, col_count, datetime_cols):
    normalized = []

    for row in rows:
        r = list(row) + [None] * (col_count - len(row))

        for i in range(col_count):
            val = r[i]
            if i in datetime_cols:
                r[i] = val if isinstance(val, datetime.datetime) else None
            else:
                r[i] = None if val is None else str(val)

        normalized.append(r)

    return normalized


# ================= CORE CONVERSION =================
def excel_to_parquet_with_progress(excel_file, sheet_name, output_file, progress_queue):
    try:
        chunk_size = auto_chunk_size()

        wb = load_workbook(excel_file, read_only=True, data_only=True)
        ws = wb[sheet_name]
        total_rows = ws.max_row

        row_ranges = [
            (start, min(start + chunk_size - 1, total_rows))
            for start in range(1, total_rows + 1, chunk_size)
        ]

        # Schema inference
        sample_rows = read_excel_chunk_worker(
            (excel_file, sheet_name, row_ranges[0][0], min(row_ranges[0][0] + 5000, row_ranges[0][1]))
        )

        schema, columns, datetime_cols, col_count = build_safe_schema(sample_rows)

        writer = pq.ParquetWriter(output_file, schema, compression="snappy")

        processed = 0
        total_chunks = len(row_ranges)

        with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
            tasks = [
                (excel_file, sheet_name, start, end)
                for start, end in row_ranges
            ]

            for rows in executor.map(read_excel_chunk_worker, tasks):
                if rows:
                    rows = normalize_rows(rows, col_count, datetime_cols)
                    table = pa.Table.from_pylist(
                        [dict(zip(columns, r)) for r in rows],
                        schema=schema
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
    root.title("Excel → Parquet (Multiprocessing Safe)")
    root.geometry("540x300")
    root.resizable(False, False)

    excel_path = tk.StringVar()
    sheet_name = tk.StringVar()
    output_name = tk.StringVar(value="output.parquet")
    q = queue.Queue()

    def browse_file():
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return

        excel_path.set(path)
        wb = load_workbook(path, read_only=True)

        sheet_menu["menu"].delete(0, "end")
        for s in wb.sheetnames:
            sheet_menu["menu"].add_command(
                label=s,
                command=lambda v=s: sheet_name.set(v)
            )
        sheet_name.set(wb.sheetnames[0])

    def start_conversion():
        if not excel_path.get() or not sheet_name.get():
            messagebox.showerror("Error", "Select Excel file and sheet")
            return

        root.withdraw()

        win = tk.Toplevel()
        win.title("Processing")
        win.geometry("420x150")
        win.resizable(False, False)

        ttk.Label(
            win,
            text="Processing Excel safely (parallel, values-only)",
            padding=10
        ).pack()

        bar = ttk.Progressbar(win, length=340, mode="determinate")
        bar.pack(pady=10)

        status = ttk.Label(win, text="Starting...")
        status.pack()

        def poll():
            try:
                msg = q.get_nowait()

                if msg[0] == "DONE":
                    win.destroy()
                    messagebox.showinfo("Success", f"Created:\n{msg[1]}")
                    root.destroy()
                    return

                if msg[0] == "ERROR":
                    win.destroy()
                    messagebox.showerror("Error", msg[1])
                    root.destroy()
                    return

                bar["maximum"] = msg[1]
                bar["value"] = msg[0]
                status.config(text=f"Processed {msg[0]} / {msg[1]} chunks")

            except queue.Empty:
                pass

            win.after(200, poll)

        threading.Thread(
            target=excel_to_parquet_with_progress,
            args=(excel_path.get(), sheet_name.get(), output_name.get(), q),
            daemon=True
        ).start()

        poll()

    # ---- Layout ----
    tk.Label(root, text="Excel File:").place(x=20, y=30)
    tk.Entry(root, textvariable=excel_path, width=45).place(x=140, y=30)
    tk.Button(root, text="Browse", command=browse_file).place(x=450, y=26)

    tk.Label(root, text="Sheet:").place(x=20, y=80)
    sheet_menu = tk.OptionMenu(root, sheet_name, "")
    sheet_menu.place(x=140, y=75, width=220)

    tk.Label(root, text="Output File:").place(x=20, y=130)
    tk.Entry(root, textvariable=output_name, width=32).place(x=140, y=130)

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


# ================= ENTRY POINT =================
if __name__ == "__main__":
    launch_gui()
