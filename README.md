```
import pandas as pd
import numpy as np
import os
import shutil

from datetime import datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# =========================================================
# TEMPLATE FILE
# =========================================================

template_file = r'Pivot_Template.xlsx'

# =========================================================
# READ REFERENCE FILE
# =========================================================

ref_df = pd.read_excel(
    r'HCIB Product_Business area split Matrix_AK.xlsx'
)

# =========================================================
# OUTPUT FOLDER
# =========================================================

timestamp = datetime.now().strftime('%d%b_%H%M')

output_folder = f'Product_slices_{timestamp}'

os.makedirs(output_folder, exist_ok=True)

# =========================================================
# SCOPE COLUMN
# =========================================================

df4['Scope'] = np.where(
    (
        df4['Description 3_mica']
        .astype(str)
        .str.contains(r'\(NIP\)', na=False)
    )
    |
    (
        df4['Description 8_mica']
        .astype(str)
        .str.contains(r'\(NIP\)', na=False)
    ),
    'Out-of-scope',
    'Inscope'
)

# =========================================================
# DESCRIPTION COLUMNS
# =========================================================

df4['Level1_mica_desc'] = (
    df4['Level 1_mica'].astype(str)
    + ' '
    + df4['Description 1_mica'].astype(str)
)

df4['Level2_mica_desc'] = (
    df4['Level 2_mica'].astype(str)
    + ' '
    + df4['Description 2_mica'].astype(str)
)

df4['Level3_mica_desc'] = (
    df4['Level 3_mica'].astype(str)
    + ' '
    + df4['Description 3_mica'].astype(str)
)

df4['Level8_mica_desc'] = (
    df4['Level 8_mica'].astype(str)
    + ' '
    + df4['Description 8_mica'].astype(str)
)

df4['Level9_mica_desc'] = (
    df4['Level 9_mica'].astype(str)
    + ' '
    + df4['Description 9_mica'].astype(str)
)

# =========================================================
# BUSINESS LOOP
# =========================================================

for business in ref_df['Business'].dropna().unique():

    print(f'Processing : {business}')

    # =====================================================
    # FILTER REFERENCE
    # =====================================================

    temp_ref = ref_df[
        ref_df['Business'] == business
    ].reset_index(drop=True)

    all_parts = []

    current_df = None

    has_cg = temp_ref['Value'].astype(str).str.startswith(
        'CG',
        na=False
    ).any()

    # =====================================================
    # CASE 1 : CG EXISTS
    # =====================================================

    if has_cg:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = row['Value']

            # ------------------------------------------------
            # START NEW DF WHEN CG COMES
            # ------------------------------------------------

            if str(filter_val).startswith('CG'):

                if current_df is not None:

                    all_parts.append(current_df)

                current_df = df4[
                    df4[filter_col].astype(str)
                    == str(filter_val)
                ].copy()

            else:

                if current_df is not None:

                    current_df = current_df[
                        current_df[filter_col].astype(str)
                        == str(filter_val)
                    ]

        if current_df is not None:

            all_parts.append(current_df)

    # =====================================================
    # CASE 2 : NO CG EXISTS
    # =====================================================

    else:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = row['Value']

            temp_df = df4[
                df4[filter_col].astype(str)
                == str(filter_val)
            ].copy()

            all_parts.append(temp_df)

    # =====================================================
    # FINAL GENERATED DF
    # =====================================================

    generated_df = pd.concat(
        all_parts,
        ignore_index=True
    ).drop_duplicates()

    # =====================================================
    # FILE NAME
    # =====================================================

    safe_business = (
        str(business)
        .replace('/', '_')
        .replace('\\', '_')
    )

    output_file = os.path.join(
        output_folder,
        f'{safe_business}_{timestamp}.xlsx'
    )

    # =====================================================
    # COPY TEMPLATE
    # =====================================================

    shutil.copy(
        template_file,
        output_file
    )

    # =====================================================
    # LOAD WORKBOOK
    # =====================================================

    wb = load_workbook(output_file)

    ws = wb['Raw_Data']

    # =====================================================
    # CLEAR OLD DATA ONLY
    # KEEP HEADER + FORMATTING + TABLE
    # =====================================================

    if ws.max_row > 1:

        ws.delete_rows(
            2,
            ws.max_row
        )

    # =====================================================
    # WRITE NEW DATA
    # =====================================================

    for row in generated_df.itertuples(
        index=False,
        name=None
    ):

        ws.append(row)

    # =====================================================
    # UPDATE TABLE RANGE
    # =====================================================

    max_row = generated_df.shape[0] + 1

    max_col = generated_df.shape[1]

    last_col_letter = get_column_letter(
        max_col
    )

    new_range = (
        f'A1:{last_col_letter}{max_row}'
    )

    ws.tables['RawTable'].ref = new_range

    # =====================================================
    # AUTO REFRESH PIVOTS ON OPEN
    # =====================================================

    for sheet in wb.worksheets:

        for pivot in sheet._pivots:

            pivot.cache.refreshOnLoad = True

    # =====================================================
    # AVOID FULL RECALCULATION
    # =====================================================

    wb.calculation.fullCalcOnLoad = False

    wb.calculation.forceFullCalc = False

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(output_file)

    print(f'Created : {output_file}')

print('All business files generated successfully.')
