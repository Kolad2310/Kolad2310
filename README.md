```
import pandas as pd
import numpy as np
import os
import shutil

from datetime import datetime
from openpyxl import load_workbook

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
    # CG LOGIC
    # =====================================================

    if has_cg:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = row['Value']

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
    # FINAL DF
    # =====================================================

    generated_df = pd.concat(
        all_parts,
        ignore_index=True
    ).drop_duplicates()

    # =====================================================
    # OUTPUT FILE
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
    # WRITE RAW DATA ONLY
    # =====================================================

    with pd.ExcelWriter(
        output_file,
        engine='openpyxl',
        mode='a',
        if_sheet_exists='replace'
    ) as writer:

        generated_df.to_excel(
            writer,
            sheet_name='Raw_Data',
            index=False
        )

    # =====================================================
    # REFRESH ON OPEN
    # =====================================================

    wb = load_workbook(output_file)

    wb.calculation.fullCalcOnLoad = True

    wb.calculation.forceFullCalc = True

    wb.save(output_file)

    print(f'Created : {output_file}')

print('All business files generated successfully.')
