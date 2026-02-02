```
import os
import threading
import queue
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
import pyarrow as pa
import pyarrow.parquet as pq


# ================= SAFE CONFIG =================
CHUNK_SIZE = 50_000   # rows per write (safe)
# ===============================================


def detect_datetime_columns(rows, col_count):
    datetime_cols = set()
    for row in rows:
        for i in range(min(len(row), col_count)):
            if isinstance(row[i], datetime.datetime):
                datetime_cols.add(i)
    return datetime_cols


def build_schema(sample_rows):
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
    out = []
    for row in rows:
        r = list(row) + [None] * (col_count - len(row))
        for i in range(col_count):
            v = r[i]
            if i in datetime_cols:
                r[i] = v if isinstance(v, datetime.datetime) else None
            else:
                r[i] = None if v is None else str(v)
        out.append(r)
    return out


# ================= CORE (NO CONCURRENCY) =================
def excel_to_parquet_sequential(excel_file, sheet_name, output_file, q):
    try:
        wb = load_workbook(excel_file, read_only=True, data_only=True)
        ws = wb[sheet_name]

        total_rows = ws.max_row
        rows_iter = ws.iter_rows(values_only=True)

        # ---- Read first chunk for schema ----
        first_chunk = []
        for _ in range(min(CHUNK_SIZE, total_rows)):
            first_chunk.append(next(rows_iter))

        schema, columns, datetime_cols, col_count = build_schema(first_chunk)

        writer = pq.ParquetWriter(
            output_file,
            schema,
            compression="snappy"
        )

        processed_rows = 0

        # ---- Write first chunk ----
        first_chunk = normalize_rows(first_chunk, col_count, datetime_cols)
        table = pa.Table.from_pylist(
            [dict(zip(columns, r)) for r in first_chunk],
            schema=schema
        )
        writer.write_table(table)

        processed_rows += len(first_chunk)
        q.put((processed_rows, total_rows))

        # ---- Stream remaining rows ----
        buffer = []
        for row in rows_iter:
            buffer.append(row)

            if len(buffer) >= CHUNK_SIZE:
                buffer = normalize_rows(buffer, col_count, datetime_cols)
                table = pa.Table.from_pylist(
                    [dict(zip(columns, r)) for r in buffer],
                    schema=schema
                )
                writer.write_table(table)

                processed_rows += len(buffer)
                q.put((processed_rows, total_rows))
                buffer.clear()

        # ---- Write remainder ----
        if buffer:
            buffer = normalize_rows(buffer, col_count, datetime_cols)
            table = pa.Table.from_pylist(
                [dict(zip(columns, r)) for r in buffer],
                schema=schema
            )
            writer.write_table(table)
            processed_rows += len(buffer)
            q.put((processed_rows, total_rows))

        writer.close()
        q.put(("DONE", output_file))

    except Exception as e:
        q.put(("ERROR", str(e)))


# ================= GUI =================
def launch_gui():
    root = tk.Tk()
    root.title("Excel → Parquet (Stable, No Concurrency)")
    root.geometry("560x320")
    root.resizable(False, False)

    excel_path = tk.StringVar()
    sheet_name = tk.StringVar()
    output_name = tk.StringVar(value="output.parquet")
    q = queue.Queue()

    def browse_file():
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
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

    def start():
        if not excel_path.get() or not sheet_name.get():
            messagebox.showerror("Error", "Select Excel file and sheet")
            return

        root.withdraw()

        win = tk.Toplevel()
        win.title("Processing")
        win.geometry("440x160")
        win.resizable(False, False)

        ttk.Label(
            win,
            text="Reading Excel sequentially\n(values only, no headers)",
            padding=10
        ).pack()

        bar = ttk.Progressbar(win, length=380, mode="determinate")
        bar.pack(pady=10)

        lbl = ttk.Label(win, text="Starting...")
        lbl.pack()

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
                lbl.config(text=f"Processed {msg[0]:,} / {msg[1]:,} rows")

            except queue.Empty:
                pass

            win.after(200, poll)

        threading.Thread(
            target=excel_to_parquet_sequential,
            args=(excel_path.get(), sheet_name.get(), output_name.get(), q),
            daemon=True
        ).start()

        poll()

    tk.Label(root, text="Excel File:").place(x=20, y=30)
    tk.Entry(root, textvariable=excel_path, width=45).place(x=150, y=30)
    tk.Button(root, text="Browse", command=browse_file).place(x=470, y=26)

    tk.Label(root, text="Sheet:").place(x=20, y=90)
    sheet_menu = tk.OptionMenu(root, sheet_name, "")
    sheet_menu.place(x=150, y=85, width=220)

    tk.Label(root, text="Output File:").place(x=20, y=150)
    tk.Entry(root, textvariable=output_name, width=32).place(x=150, y=150)

    tk.Button(
        root,
        text="Start Conversion",
        width=24,
        height=2,
        bg="#4CAF50",
        fg="white",
        command=start
    ).place(x=200, y=220)

    root.mainloop()


# ================= ENTRY =================
if __name__ == "__main__":
    launch_gui()
