```
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import RGBColor

def generate_commentary_word(df, value_col, l1_col, l2_col, output_file="Commentary.docx", top_n=3):

    df = df.copy()
    df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0)

    doc = Document()
    doc.add_heading('Performance Commentary', 0)

    # ----------------------
    # L1 Summary
    # ----------------------
    doc.add_heading('L1 Summary', 1)

    l1 = df.groupby(l1_col)[value_col].sum().reset_index()
    l1 = l1.sort_values(by=value_col, ascending=False)

    for _, row in l1.iterrows():
        p = doc.add_paragraph()
        run = p.add_run(f"{row[l1_col]}: {row[value_col]:,.2f}")
        run.bold = True

    # ----------------------
    # L2 Summary
    # ----------------------
    doc.add_heading('L2 Breakdown', 1)

    l2 = df.groupby([l1_col, l2_col])[value_col].sum().reset_index()
    l2 = l2.sort_values(by=value_col, ascending=False)

    for _, row in l2.iterrows():
        p = doc.add_paragraph(f"{row[l2_col]} under {row[l1_col]}: {row[value_col]:,.2f}")

    # ----------------------
    # L3 Explanation
    # ----------------------
    doc.add_heading('Key Drivers (L3)', 1)

    top = df.sort_values(by=value_col, ascending=False).head(top_n)
    bottom = df.sort_values(by=value_col, ascending=True).head(top_n)

    l3 = pd.concat([top, bottom]).drop_duplicates()

    for _, row in l3.iterrows():
        p = doc.add_paragraph()

        if row[value_col] > 0:
            text = f"{row[l2_col]} under {row[l1_col]} drove an increase of {row[value_col]:,.2f}"
            color = RGBColor(0, 128, 0)  # green
        elif row[value_col] < 0:
            text = f"{row[l2_col]} under {row[l1_col]} caused a decrease of {row[value_col]:,.2f}"
            color = RGBColor(255, 0, 0)  # red
        else:
            text = f"{row[l2_col]} under {row[l1_col]} had no impact"
            color = RGBColor(0, 0, 0)

        run = p.add_run(text)
        run.font.color.rgb = color

    # Save document
    doc.save(output_file)

    return output_file


   5 
