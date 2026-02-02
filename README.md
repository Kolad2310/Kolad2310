```
pip install pandas openpyxl pyarrow psutil tqdm
import os
import psutil
import pandas as pd
from openpyxl import load_workbook
from concurrent.futures import ProcessPoolExecutor, as_completed
import pyarrow as pa
import pyarrow.parquet as pq
from tqdm import tqdm


# ================= CONFIG =================
EXCEL_FILE = "big_file.xlsx"
SHEET_NAME = "Sheet1"
OUTPUT_PARQUET = "sheet1_values.parquet"

MAX_WORKERS = max(1, os.cpu_count() - 1)
ROWS_PER_MB = 450      # empirical avg for Excel rows
MIN_CHUNK = 50_000
MAX_CHUNK = 300_000
# ==========================================


def auto_chunk_size():
    """
    Auto-tune chunk size based on available RAM
    """
    available_mb = psutil.virtual_memory().available // (1024 * 1024)
    chunk = int(available_mb * ROWS_PER_MB * 0.25)
    return max(MIN_CHUNK, min(chunk, MAX_CHUNK))


def infer_schema(sample_rows, headers):
    """
    Enforce schema to avoid dtype drift & reduce memory
    """
    df = pd.DataFrame(sample_rows, columns=headers)
    return pa.Schema.from_pandas(df, preserve_index=False)


def read_excel_chunk(file_path, sheet_name, start_row, end_row):
    """
    Read a chunk of rows (VALUES ONLY)
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


def excel_to_parquet_optimized():
    # ---------- Auto chunk size ----------
    chunk_size = auto_chunk_size()
    print(f"🔧 Auto-selected chunk size: {chunk_size:,}")

    # ---------- Read header & total rows ----------
    wb = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    total_rows = ws.max_row
    headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))

    row_ranges = [
        (start, min(start + chunk_size - 1, total_rows))
        for start in range(2, total_rows + 1, chunk_size)
    ]

    # ---------- Infer schema from first chunk ----------
    first_start, first_end = row_ranges[0]
    sample_rows = read_excel_chunk(
        EXCEL_FILE,
        SHEET_NAME,
        first_start,
        min(first_start + 5_000, first_end)
    )
    schema = infer_schema(sample_rows, headers)

    writer = pq.ParquetWriter(
        OUTPUT_PARQUET,
        schema,
        compression="snappy"
    )

    # ---------- Parallel execution ----------
    with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(
                read_excel_chunk,
                EXCEL_FILE,
                SHEET_NAME,
                start,
                end
            )
            for start, end in row_ranges
        ]

        for future in tqdm(
            as_completed(futures),
            total=len(futures),
            desc="📊 Processing Excel chunks"
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
    print("✅ Done. Parquet file ready.")


# ================= RUN =================
if __name__ == "__main__":
    excel_to_parquet_optimized()
