```
import pandas as pd
import numpy as np
import re
from docx import Document
from docx.shared import RGBColor

# =========================================================
# UNIT + FORMATTING
# =========================================================

def detect_unit(metric_col):
    t = metric_col.upper()
    if "$K" in t:
        return "K"
    if "$M" in t:
        return "M"
    return None

def to_mn(value, unit):
    if pd.isna(value):
        return value
    return value / 1000 if unit == "K" else value

def fmt_mn_bn(mn):
    if pd.isna(mn):
        return "N/A"
    sign = "+" if mn > 0 else "-" if mn < 0 else ""
    mn = abs(mn)
    if mn >= 1000:
        return f"{sign}{mn/1000:.1f}bn"
    return f"{sign}{mn:.0f}m"

def fmt_change_yoy(change, yoy, metric_col):
    unit = detect_unit(metric_col)
    mn = to_mn(change, unit)
    ch = fmt_mn_bn(mn)
    if pd.isna(yoy):
        return f"{ch} / N/A"
    return f"{ch} / {yoy:.1f}%"

# =========================================================
# AGGREGATION + FILTERING
# =========================================================

def compute_agg(df_cy, df_py, group_cols, metric_col):
    cy = df_cy.groupby(group_cols)[metric_col].sum().reset_index(name="CY")
    py = df_py.groupby(group_cols)[metric_col].sum().reset_index(name="PY")
    m = cy.merge(py, on=group_cols, how="left").fillna(0)
    m["Change"] = m["CY"] - m["PY"]
    m["YoY%"] = np.where(m["PY"] != 0, m["Change"] / m["PY"] * 100, np.nan)
    return m

def drop_noise(df):
    return df[~((df["Change"] == 0) & ((df["YoY%"] == 0) | df["YoY%"].isna()))]

# =========================================================
# 90% COVERAGE SELECTION
# =========================================================

def select_by_coverage(df, metric_col, pct=0.9, min_mn=5):
    if df.empty:
        return df

    unit = detect_unit(metric_col)
    df = df.copy()
    df["_abs_mn"] = df["Change"].abs().apply(lambda x: to_mn(x, unit))
    df = df[df["_abs_mn"] >= min_mn]

    if df.empty:
        return df

    total = df["_abs_mn"].sum()
    cutoff = total * pct

    rows, run = [], 0
    for _, r in df.iterrows():
        rows.append(r)
        run += r["_abs_mn"]
        if run >= cutoff:
            break

    return pd.DataFrame(rows).drop(columns="_abs_mn")

def join_items(df, name_col, metric_col):
    return " and ".join(
        f"{r[name_col]} ({fmt_change_yoy(r['Change'], r['YoY%'], metric_col)})"
        for _, r in df.iterrows()
    )

# =========================================================
# GENERIC HIERARCHICAL COMMENTARY
# =========================================================

def commentary_by_hierarchy(
    df_cy,
    df_py,
    metric_col,
    level1_col,
    level2_col,
    title
):
    agg_lvl1 = drop_noise(
        compute_agg(df_cy, df_py, [level1_col], metric_col)
    )

    pos = agg_lvl1[agg_lvl1["Change"] > 0].sort_values("Change", ascending=False)
    neg = agg_lvl1[agg_lvl1["Change"] < 0].sort_values("Change")

    top = select_by_coverage(pos, metric_col)
    bottom = select_by_coverage(neg, metric_col)

    def describe_lvl1(r):
        lvl1 = r[level1_col]
        base = f"{lvl1} ({fmt_change_yoy(r['Change'], r['YoY%'], metric_col)})"

        agg_lvl2 = drop_noise(
            compute_agg(
                df_cy[df_cy[level1_col] == lvl1],
                df_py[df_py[level1_col] == lvl1],
                [level2_col],
                metric_col
            )
        )

        if r["Change"] > 0:
            drivers = select_by_coverage(
                agg_lvl2[agg_lvl2["Change"] > 0]
                .sort_values("Change", ascending=False),
                metric_col
            )
        else:
            drivers = select_by_coverage(
                agg_lvl2[agg_lvl2["Change"] < 0]
                .sort_values("Change"),
                metric_col
            )

        return (
            f"{base} driven by {join_items(drivers, level2_col, metric_col)}"
            if not drivers.empty else base
        )

    text = f"{title}: Growth was led by "
    text += " and ".join(describe_lvl1(r) for _, r in top.iterrows())

    if not bottom.empty:
        offset = " and ".join(describe_lvl1(r) for _, r in bottom.iterrows())
        text += f", partially offset by {offset}"

    return text

# =========================================================
# WORD WRITER (COLORED NUMBERS)
# =========================================================

def write_commentary_to_word_colored(commentary_dict, output_file):
    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

    pattern = re.compile(
        r"(\+|-)?\d+(\.\d+)?(m|bn)(\s*/\s*\+?\-?\d+(\.\d+)?%)?"
    )

    for title, text in commentary_dict.items():
        doc.add_heading(title, level=2)
        p = doc.add_paragraph()
        idx = 0
        matches = list(pattern.finditer(text))

        if not matches:
            p.add_run(text)
            continue

        for m in matches:
            start, end = m.span()
            if start > idx:
                p.add_run(text[idx:start])

            run = p.add_run(text[start:end])
            run.font.color.rgb = (
                RGBColor(0, 176, 80) if text[start] == "+" else RGBColor(192, 0, 0)
            )
            idx = end

        if idx < len(text):
            p.add_run(text[idx:])

    doc.save(output_file)

# =========================================================
# FINAL DRIVER
# =========================================================

def build_commentary(df_cy, df_py, metric_col):
    return {
        "By Segment": commentary_by_hierarchy(
            df_cy, df_py, metric_col,
            level1_col="Segment",
            level2_col="Business Line",
            title="By Segment"
        ),
        "By Product": commentary_by_hierarchy(
            df_cy, df_py, metric_col,
            level1_col="Business Line",
            level2_col="Region2",
            title="By Product"
        ),
        "By Region": commentary_by_hierarchy(
            df_cy, df_py, metric_col,
            level1_col="Region2",
            level2_col="Business Line",
            title="By Region"
        )
    }

# =======================
# USAGE
# =======================
# commentary = build_commentary(df_cy, df_py, "Total Relationship Income ($M)")
# write_commentary_to_word_colored(commentary, "TRI_Commentary_Final.docx")
