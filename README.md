```
import pandas as pd
import calendar

def compute_ytd(df, month_num):
    months = [calendar.month_abbr[i] for i in range(1, month_num + 1)]

    # ---------- 2025 ----------
    actual_25_cols = [
        f"Actuals_{m}_25" for m in months
        if f"Actuals_{m}_25" in df.columns
    ]

    if actual_25_cols:
        df["YTD_2025"] = (
            df[actual_25_cols]
            .apply(pd.to_numeric, errors='coerce')
            .sum(axis=1)
        )

    # ---------- 2026 Actuals ----------
    actual_26_cols = [
        f"Actuals_{m}_26" for m in months
        if f"Actuals_{m}_26" in df.columns
    ]

    if actual_26_cols:
        df["YTD Actuals 2026"] = (
            df[actual_26_cols]
            .apply(pd.to_numeric, errors='coerce')
            .sum(axis=1)
        )

    # ---------- 2026 Targets ----------
    target_26_cols = [
        f"Monthly Target_{m}_26" for m in months
        if f"Monthly Target_{m}_26" in df.columns
    ]

    if target_26_cols:
        df["YTD Monthly Target_2026"] = (
            df[target_26_cols]
            .apply(pd.to_numeric, errors='coerce')
            .sum(axis=1)
        )

    return df
