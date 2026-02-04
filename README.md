```
import pandas as pd
import numpy as np
import re
from docx import Document
from docx.shared import RGBColor

# =========================================================
# FORMATTING HELPERS
# =========================================================

def detect_unit(metric_col):
    col = metric_col.upper()
    if "$K" in col:
        return "K"
    if "$M" in col:
        return "M"
    return None

def to_mn(val, unit):
    if pd.isna(val):
        return val
    return val / 1000 if unit == "K" else val

def fmt_mn_bn(val):
    if pd.isna(val):
        return "N/A"
    sign = "+" if val > 0 else "-" if val < 0 else ""
    val = abs(val)
    return f"{sign}{val/1000:.1f}bn" if val >= 1000 else f"{sign}{val:.0f}m"

def fmt_change_yoy(change, yoy, metric_col):
    unit = detect_unit(metric_col)
    ch = fmt_mn_bn(to_mn(change, unit))
    if pd.isna(yoy):
        return ch
    return f"{ch} ({yoy:.1f}%)"

# =========================================================
# AGGREGATION
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
# DRIVER SELECTION (SHORT)
# =========================================================

def select_key_drivers(df, metric_col, max_items=2, min_mn=5):
    if df.empty:
        return df

    unit = detect_unit(metric_col)
    df = df.copy()
    df["_abs"] = df["Change"].abs().apply(lambda x: to_mn(x, unit))
    df = df[df["_abs"] >= min_mn]

    return df.head(max_items)

# =========================================================
# WORD-STYLE LINKING
# =========================================================

def join_word_style(df, name_col, metric_col):
    items = [
        f"{r[name_col]} {fmt_change_yoy(r['Change'], r['YoY%'], metric_col)}"
        for _, r in df.iterrows()
    ]

    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    if len(items) >= 3:
        return f"{items[0]}, followed by {items[1]} and {items[2]}"
    return ""

# =========================================================
# MAIN COMMENTARY BUILDER (WITH OFFSET SUB-PARAGRAPH)
# =========================================================

def build_section_commentary(
    df_cy,
    df_py,
    metric_col,
    lvl1,
    lvl2,
    title
):
    agg = drop_noise(compute_agg(df_cy, df_py, [lvl1], metric_col))

    pos = agg[agg["Change"] > 0].sort_values("Change", ascending=False)
    neg = agg[agg["Change"] < 0].sort_values("Change")

    top = select_key_drivers(pos, metric_col, max_items=2)
    bottom = select_key_drivers(neg, metric_col, max_items=2)

    def describe(row, positive=True):
        name = row[lvl1]

        lvl2_agg = drop_noise(
            compute_agg(
                df_cy[df_cy[lvl1] == name],
                df_py[df_py[lvl1] == name],
                [lvl2],
                metric_col
            )
        )

        drivers = select_key_drivers(
            lvl2_agg[lvl2_agg["Change"] > 0]
            if positive else
            lvl2_agg[lvl2_agg["Change"] < 0],
            metric_col,
            max_items=2
        )

        base = f"{name} {fmt_change_yoy(row['Change'], row['YoY%'], metric_col)}"
        return f"{base}, {join_word_style(drivers, lvl2, metric_col)}" \
            if not drivers.empty else base

    # Main paragraph
    main_text = f"{title}: Growth was led by "
    main_text += " and ".join(describe(r, True) for _, r in top.iterrows())

    # Offsetting sub-paragraph
    offset_text = None
    if not bottom.empty:
        offset_text = "Offsetting factors: "
        offset_text += " and ".join(describe(r, False) for _, r in bottom.iterrows())

    return main_text, offset_text

# =========================================================
# BUILD ALL SECTIONS
# =========================================================

def build_commentary(df_cy, df_py, metric_col):
    sections = {}

    for title, l1, l2 in [
        ("By Segment", "Segment", "Business Line"),
        ("By Product", "Business Line", "Region2"),
        ("By Region", "Region2", "Business Line"),
    ]:
        main, offset = build_section_commentary(
            df_cy, df_py, metric_col, l1, l2, title
        )
        sections[title] = {
            "main": main,
            "offset": offset
        }

    return sections

# =========================================================
# WORD WRITER (COLORED NUMBERS)
# =========================================================

def write_word(commentary, output_file):
    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

    pattern = re.compile(r"(\+|-)\d+(\.\d+)?(m|bn)|\(\-?\+?\d+(\.\d+)?%\)")

    for section, content in commentary.items():
        doc.add_heading(section, level=2)

        for text in [content["main"], content["offset"]]:
            if not text:
                continue

            p = doc.add_paragraph()
            idx = 0

            for m in pattern.finditer(text):
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
# USAGE
# =========================================================
# commentary = build_commentary(df_cy, df_py, "Total Relationship Income ($M)")
# write_word(commentary, "TRI_Commentary_Final.docx")
