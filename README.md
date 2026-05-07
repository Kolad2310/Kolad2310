```
# =========================================================
# HEADER MAPPING
# =========================================================

header_mapping = {
    'Product Performance': ['L1', 'L3'],
    'REF Performance': ['L2', 'L4']
}

# =========================================================
# WRITE TO WORD WITH HEADERS
# =========================================================

doc = Document()

doc.add_heading(
    'Financial Commentary',
    level=1
)

for header, labels in header_mapping.items():

    # ---------------------------------------------
    # WRITE HEADER
    # ---------------------------------------------

    doc.add_heading(header, level=2)

    # ---------------------------------------------
    # WRITE COMMENTARIES UNDER HEADER
    # ---------------------------------------------

    subset = final_commentary_df[
        final_commentary_df['Label'].isin(labels)
    ]

    for _, row in subset.iterrows():

        p = doc.add_paragraph()

        write_colored_commentary(
            p,
            row['Commentary']
        )

# =========================================================
# SAVE
# =========================================================

doc.save('Financial_Commentary.docx')
