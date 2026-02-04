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

def to_mn(v, unit):
    if pd.isna(v):
        return v
    return v / 1000 if unit == "K" else v

def fmt_mn_bn(v):
    sign = "+" if v > 0 else "-" if v < 0 else ""
    v = abs(v)
    return f"{sign}{v/1000:.1f}bn" if v >= 1000 else f"{sign}{v:.0f}m"

def fmt_change_yoy(change, yoy, metric_col):
    unit = detect_unit(metric_col)
    ch = fmt_mn_bn(to_mn(change, unit))
    return f"{ch} ({yoy:.1f}%)" if not pd.isna(yoy) else ch

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
# CONTRIBUTION-BASED SELECTION
# =========================================================

def select_level1(df, metric_col, positive=True, min_mn=5):
    if df.empty:
        return df

    unit = detect_unit(metric_col)
    df = df.copy()
    df["_mn"] = df["Change"].apply(lambda x: to_mn(x, unit))
    df = df[df["_mn"].abs() >= min_mn]

    return df.sort_values("Change", ascending=not positive)

# =========================================================
# BULLET BUILDERS
# =========================================================

def build_2level_bullet(row, df_cy, df_py, lvl1, lvl2, metric_col, positive=True):
    name = row[lvl1]
    headline = f"{name}: {fmt_change_yoy(row['Change'], row['YoY%'], metric_col)}"

    lvl2_agg = drop_noise(
        compute_agg(
            df_cy[df_cy[lvl1] == name],
            df_py[df_py[lvl1] == name],
            [lvl2],
            metric_col
        )
    )

    lvl2_agg = select_level1(
        lvl2_agg[lvl2_agg["Change"] > 0] if positive else lvl2_agg[lvl2_agg["Change"] < 0],
        metric_col,
        positive
    ).head(2)

    if lvl2_agg.empty:
        return headline

    parts = [
        f"{r[lvl2]} {fmt_change_yoy(r['Change'], r['YoY%'], metric_col)}"
        for _, r in lvl2_agg.iterrows()
    ]

    return f"{headline} with " + " and ".join(parts)

# ------------------ 3-LEVEL REGION ----------------------

def build_region_bullet(region, df_cy, df_py, metric_col, positive=True):
    df_cy_r = df_cy[df_cy["Region2"] == region]
    df_py_r = df_py[df_py["Region2"] == region]

    country_agg = drop_noise(
        compute_agg(df_cy_r, df_py_r, ["Managed Country"], metric_col)
    )

    country_agg = select_level1(
        country_agg[country_agg["Change"] > 0] if positive else country_agg[country_agg["Change"] < 0],
        metric_col,
        positive
    ).head(2)

    parts = []

    for _, c in country_agg.iterrows():
        country = c["Managed Country"]
        txt = f"{country} {fmt_change_yoy(c['Change'], c['YoY%'], metric_col)}"

        biz_agg = drop_noise(
            compute_agg(
                df_cy_r[df_cy_r["Managed Country"] == country],
                df_py_r[df_py_r["Managed Country"] == country],
                ["Business Line"],
                metric_col
            )
        )

        biz_agg = select_level1(
            biz_agg[biz_agg["Change"] > 0] if positive else biz_agg[biz_agg["Change"] < 0],
            metric_col,
            positive
        ).head(2)

        if not biz_agg.empty:
            biz_parts = [
                f"{r['Business Line']} {fmt_change_yoy(r['Change'], r['YoY%'], metric_col)}"
                for _, r in biz_agg.iterrows()
            ]
            txt += " driven by " + " and ".join(biz_parts)

        parts.append(txt)

    return f"{region}: " + "; ".join(parts)

# =========================================================
# SECTION BUILDERS
# =========================================================

def build_section_2level(df_cy, df_py, metric_col, lvl1, lvl2):
    agg = drop_noise(compute_agg(df_cy, df_py, [lvl1], metric_col))

    pos = select_level1(agg[agg["Change"] > 0], metric_col, True)
    neg = select_level1(agg[agg["Change"] < 0], metric_col, False)

    bullets_pos = [
        build_2level_bullet(r, df_cy, df_py, lvl1, lvl2, metric_col, True)
        for _, r in pos.iterrows()
    ]

    bullets_neg = [
        build_2level_bullet(r, df_cy, df_py, lvl1, lvl2, metric_col, False)
        for _, r in neg.iterrows()
    ]

    return bullets_pos, bullets_neg

def build_section_region(df_cy, df_py, metric_col):
    agg = drop_noise(compute_agg(df_cy, df_py, ["Region2"], metric_col))

    pos = select_level1(agg[agg["Change"] > 0], metric_col, True)
    neg = select_level1(agg[agg["Change"] < 0], metric_col, False)

    bullets_pos = [
        build_region_bullet(r["Region2"], df_cy, df_py, metric_col, True)
        for _, r in pos.iterrows()
    ]

    bullets_neg = [
        build_region_bullet(r["Region2"], df_cy, df_py, metric_col, False)
        for _, r in neg.iterrows()
    ]

    return bullets_pos, bullets_neg

# =========================================================
# WORD WRITER (SIGN-AWARE COLOURING)
# =========================================================

def write_word(commentary, output):
    doc = Document()
    doc.add_heading("Global CIB Performance", level=1)

    pattern = re.compile(r"\-?\d+(\.\d+)?(m|bn)|\(\-?\d+(\.\d+)?%\)")

    for section, content in commentary.items():
        doc.add_heading(section, level=2)

        for block in ["positive", "negative"]:
            if block == "negative" and content["negative"]:
                doc.add_paragraph("Offsetting factors:")

            for line in content[block]:
                p = doc.add_paragraph(style="List Bullet")
                idx = 0

                for m in pattern.finditer(line):
                    start, end = m.span()
                    token = line[start:end]

                    if start > idx:
                        p.add_run(line[idx:start])

                    run = p.add_run(token)
                    run.font.color.rgb = (
                        RGBColor(192, 0, 0) if token.startswith("-") or token.startswith("(-")
                        else RGBColor(0, 176, 80)
                    )

                    idx = end

                if idx < len(line):
                    p.add_run(line[idx:])

    doc.save(output)

# =========================================================
# DRIVER
# =========================================================

def build_commentary(df_cy, df_py, metric_col):
    commentary = {}

    pos, neg = build_section_2level(df_cy, df_py, metric_col, "Segment", "Business Line")
    commentary["By Segment"] = {"positive": pos, "negative": neg}

    pos, neg = build_section_2level(df_cy, df_py, metric_col, "Business Line", "Region2")
    commentary["By Product"] = {"positive": pos, "negative": neg}

    pos, neg = build_section_region(df_cy, df_py, metric_col)
    commentary["By Region"] = {"positive": pos, "negative": neg}

    return commentary

# =======================
# USAGE
# =======================
# commentary = build_commentary(df_cy, df_py, "Total Relationship Income ($M)")
# write_word(commentary, "TRI_Commentary_Final.docx")
