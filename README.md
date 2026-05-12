```
import pandas as pd
import numpy as np
import os
from datetime import datetime

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
# LOOP BUSINESS
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
    # FINAL DF
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
    # WRITE EXCEL
    # =====================================================

    with pd.ExcelWriter(
        output_file,
        engine='xlsxwriter'
    ) as writer:

        # =================================================
        # RAW DATA
        # =================================================

        generated_df.to_excel(
            writer,
            sheet_name='Raw_Data',
            index=False
        )

        workbook = writer.book

        raw_ws = writer.sheets['Raw_Data']

        # =================================================
        # FORMATS
        # =================================================

        grey_header = workbook.add_format({
            'bold': True,
            'bg_color': '#BFBFBF',
            'border': 1
        })

        number_format = workbook.add_format({
            'num_format': '#,##0'
        })

        # =================================================
        # AUTO WIDTH
        # =================================================

        for idx, col in enumerate(generated_df.columns):

            try:

                max_len = max(
                    generated_df[col]
                    .astype(str)
                    .map(len)
                    .max(),
                    len(col)
                ) + 3

            except:

                max_len = len(col) + 3

            raw_ws.set_column(
                idx,
                idx,
                min(max_len, 60)
            )

        # =================================================
        # HEADER FORMAT
        # =================================================

        for col_num, value in enumerate(generated_df.columns):

            raw_ws.write(
                0,
                col_num,
                value,
                grey_header
            )

        # =================================================
        # NUMBER FORMAT
        # =================================================

        numeric_cols = generated_df.select_dtypes(
            include='number'
        ).columns

        for col in numeric_cols:

            col_idx = generated_df.columns.get_loc(col)

            raw_ws.set_column(
                col_idx,
                col_idx,
                None,
                number_format
            )

        # =================================================
        # CREATE EXCEL TABLE
        # =================================================

        rows, cols = generated_df.shape

        raw_ws.add_table(
            0,
            0,
            rows,
            cols - 1,
            {
                'name': 'RawTable',
                'columns': [
                    {'header': c}
                    for c in generated_df.columns
                ],
                'style': 'Table Style Medium 2'
            }
        )

        # =================================================
        # CREATE EMPTY SHEETS
        # =================================================

        workbook.add_worksheet('MICA_View_PL')
        workbook.add_worksheet('MICA_View_BS')
        workbook.add_worksheet('MICA_View_AVB')
        workbook.add_worksheet('MI_Func_RTNs')
        workbook.add_worksheet('Entity_View')

        # =================================================
        # INSTRUCTIONS SHEET
        # =================================================

        instruction_ws = workbook.add_worksheet(
            'Pivot_Instructions'
        )

        instruction_text = [

            'THIS FILE CONTAINS DYNAMIC EXCEL TABLES',
            '',
            'TO CREATE / REFRESH PIVOTS:',
            '',
            '1. Open Raw_Data sheet',
            '2. Click anywhere inside the table',
            '3. Insert -> Pivot Table',
            '4. Select Existing Worksheet',
            '',
            'P&L VIEW',
            'Rows:',
            '- Level1_mica_desc',
            '- Level3_mica_desc',
            '- Level8_mica_desc',
            '- Level9_mica_desc',
            '',
            'Columns:',
            '- Source_sys',
            '',
            'Values:',
            '- Sum of P&L',
            '',
            'Filter:',
            '- Scope',
            '',
            'BS VIEW',
            'Rows:',
            '- Level1_mica_desc',
            '- Level2_mica_desc',
            '- Level3_mica_desc',
            '',
            'Columns:',
            '- Source_sys',
            '',
            'Values:',
            '- Sum of BS',
            '',
            'Filter:',
            '- Scope',
            '',
            'AVB VIEW',
            'Rows:',
            '- Level1_mica_desc',
            '- Level2_mica_desc',
            '- Level3_mica_desc',
            '',
            'Columns:',
            '- Source_sys',
            '',
            'Values:',
            '- Sum of AVB',
            '',
            'Filter:',
            '- Scope',
            '',
            'MI FUNCTION VIEW',
            'Rows:',
            '- Consolidated Period Mi Function Code',
            '- Function Leaf Description',
            '- Function Level 3',
            '- Function Description',
            '',
            'Columns:',
            '- Source_sys',
            '',
            'Values:',
            '- Sum of AVB',
            '- Sum of BS',
            '- Sum of P&L',
            '',
            'Filter:',
            '- Scope',
            '',
            'ENTITY VIEW',
            'Rows:',
            '- Consolidated Period Entity ID',
            '',
            'Columns:',
            '- Source_sys',
            '',
            'Values:',
            '- Sum of AVB',
            '- Sum of BS',
            '- Sum of P&L',
            '',
            'Filter:',
            '- Scope'
        ]

        for row_num, line in enumerate(instruction_text):

            instruction_ws.write(
                row_num,
                0,
                line
            )

        instruction_ws.set_column(
            0,
            0,
            60
        )

    print(f'Created : {output_file}')

print('All business files generated successfully.')
