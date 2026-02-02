```
import datetime

def detect_datetime_columns(rows, col_count):
    datetime_cols = set()

    for row in rows:
        for idx in range(col_count):
            if idx < len(row) and isinstance(row[idx], datetime.datetime):
                datetime_cols.add(idx)

    return datetime_cols
