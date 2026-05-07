```
import pandas as pd
import numpy as np
import re

from docx import Document
from docx.shared import RGBColor

# =========================================================
# SAMPLE:
# df already contains below columns
# =========================================================
#
# ['Region',
#  'Country',
#  'Business Line',
#  'YTD_2025',
#  'YTD_2026',
#  'YTD_Monthly_Target_2026',
#  'FY_2025',
#  'FY_Forecast_2026',
#  'FY_Target_2026',
#  'Label']
#
# =========================================================


# =========================================================
# AGGREGATE LEVEL 1
# =========================================================

level1 = df.groupby('Label', as_index=False).agg({
    'YTD_2026': 'sum',
    'YTD_Monthly_Target_2026': 'sum',
    'FY_Forecast_2026': 'sum',
    'FY_Target_2026': 'sum'
})

commentaries = []

# =========================================================
# GENERATE COMMENTARY
# =========================================================

for _, row in level1.iterrows():

    label = row['Label']

    ytd_actual = row['YTD_2026']
    ytd_target = row['YTD_Monthly_Target_2026']

    fy_fcst = row['FY_Forecast_2026']
    fy_target = row['FY_Target_2026']

    # =====================================================
    # YTD VARIANCE
    # =====================================================

    ytd_var = ytd_actual - ytd_target

    ytd_pct = (
        (ytd_var / ytd_target) * 100
        if ytd_target != 0 else 0
    )

    if ytd_var >= 0:

        ytd_status = 'above'
        sort_ascending = False
        offset_word = 'partly offset by'

    else:

        ytd_status = 'below'
        sort_ascending = True
        offset_word = 'partly onset by'

    # =====================================================
    # COUNTRY DRIVERS
    # =====================================================

    temp = df[df['Label'] == label].copy()

    temp['YTD_VAR'] = (
        temp['YTD_2026'] -
        temp['YTD_Monthly_Target_2026']
    )

    temp = temp.sort_values(
        'YTD_VAR',
        ascending=sort_ascending
    )

    # Main driver
    top_country = temp.iloc[0]['Country']

    # Offset country
    offset_country = ''

    if len(temp) > 1:
        offset_country = temp.iloc[-1]['Country']

    # =====================================================
    # FY COMMENTARY
    # =====================================================

    fy_var = fy_fcst - fy_target

    if round(fy_var, 2) == 0:

        fy_comment = (
            "FY Forecast 2026 is on track vs target."
        )

    else:

        fy_pct = (
            (fy_var / fy_target) * 100
            if fy_target != 0 else 0
        )

        fy_status = (
            'above'
            if fy_var > 0
            else 'below'
        )

        fy_comment = (
            f"FY Forecast 2026 is "
            f"{fy_status} target by "
            f"{fy_var:,.1f} "
            f"({fy_pct:.1f}% vs target)."
        )

    # =====================================================
    # FINAL COMMENTARY
    # =====================================================

    commentary = (
        f"{label}: "
        f"YTD is {ytd_status} YTD target by "
        f"{ytd_var:,.1f} "
        f"({ytd_pct:.1f}%), "
        f"driven by {top_country}, "
        f"{offset_word} {offset_country}. "
        f"{fy_comment}"
    )

    commentaries.append(commentary)

# =========================================================
# FINAL COMMENTARY DATAFRAME
# =========================================================

final_commentary_df = pd.DataFrame({
    'Label': level1['Label'],
    'Commentary': commentaries
})

print(final_commentary_df)


# =========================================================
# FUNCTION TO WRITE COLORED TEXT TO WORD
# =========================================================

doc = Document()

def write_colored_commentary(paragraph, text):

    # Pattern captures:
    # 12.5
    # -12.5
    # 12.5%
    # -12.5%
    # (12.5%)
    # (-12.5%)
    # 12.5m
    # -12.5m

    pattern = r'''
        \(-?\$?\d+\.?\d*[%mMkKbB]?\) |
        -?\$?\d+\.?\d*[%mMkKbB]?
    '''

    parts = re.split(f'({pattern})', text, flags=re.VERBOSE)

    for part in parts:

        if not part:
            continue

        run = paragraph.add_run(part)

        nums = re.findall(r'-?\d+\.?\d*', part)

        if nums:

            value = float(nums[0])

            # =================================================
            # NEGATIVE -> RED
            # =================================================

            if '-' in part:

                run.font.color.rgb = RGBColor(255, 0, 0)

            # =================================================
            # POSITIVE -> GREEN
            # =================================================

            else:

                run.font.color.rgb = RGBColor(0, 128, 0)

# =========================================================
# WRITE TO WORD
# =========================================================

doc.add_heading('Financial Commentary', level=1)

for text in final_commentary_df['Commentary']:

    p = doc.add_paragraph()

    write_colored_commentary(
        p,
        text
    )

# =========================================================
# SAVE DOCUMENT
# =========================================================

doc.save('Financial_Commentary.docx')

print("Word document created successfully.")
