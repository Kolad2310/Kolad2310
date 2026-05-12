```
import pandas as pd
import numpy as np
import os
from datetime import datetime

import win32com.client as win32
from openpyxl import load_workbook

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
# GENERATE BUSINESS LEVEL DATAFRAMES
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
    # WRITE RAW DATA
    # =====================================================

    with pd.ExcelWriter(
        output_file,
        engine='xlsxwriter'
    ) as writer:

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
                min(max_len, 50)
            )

        # =================================================
        # HEADER FORMAT
        # =================================================

        for col_num, value in enumerate(generated_df.columns.values):

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
        # EMPTY PIVOT SHEETS
        # =================================================

        workbook.add_worksheet('MICA_View_PL')
        workbook.add_worksheet('MICA_View_BS')
        workbook.add_worksheet('MICA_View_AVB')
        workbook.add_worksheet('MI_Func_RTNs')
        workbook.add_worksheet('Entity_View')

    # =====================================================
    # CREATE REAL EXCEL PIVOTS USING WIN32
    # =====================================================

    excel = win32.gencache.EnsureDispatch('Excel.Application')

    excel.Visible = False

    wb = excel.Workbooks.Open(
        os.path.abspath(output_file)
    )

    # =====================================================
    # SOURCE RANGE
    # =====================================================

    source_sheet = wb.Sheets('Raw_Data')

    last_row = source_sheet.Cells(
        source_sheet.Rows.Count,
        1
    ).End(-4162).Row

    last_col = source_sheet.Cells(
        1,
        source_sheet.Columns.Count
    ).End(-4159).Column

    source_range = (
        f"Raw_Data!R1C1:R{last_row}C{last_col}"
    )

    pivot_cache = wb.PivotCaches().Create(
        SourceType=1,
        SourceData=source_range
    )

    # =====================================================
    # HELPER FUNCTION
    # =====================================================

    def create_pivot(
        sheet_name,
        rows,
        columns,
        values
    ):

        ws = wb.Sheets(sheet_name)

        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=f"{sheet_name}!R3C1",
            TableName=f"Pivot_{sheet_name}"
        )

        # ================================================
        # FILTERS
        # ================================================

        pivot_table.PivotFields(
            'Scope'
        ).Orientation = 3

        pivot_table.PivotFields(
            'Scope'
        ).Position = 1

        # ================================================
        # ROWS
        # ================================================

        for idx, row_field in enumerate(rows):

            pivot_table.PivotFields(
                row_field
            ).Orientation = 1

            pivot_table.PivotFields(
                row_field
            ).Position = idx + 1

        # ================================================
        # COLUMNS
        # ================================================

        for idx, col_field in enumerate(columns):

            pivot_table.PivotFields(
                col_field
            ).Orientation = 2

            pivot_table.PivotFields(
                col_field
            ).Position = idx + 1

        # ================================================
        # VALUES
        # ================================================

        for value_field in values:

            pivot_table.AddDataField(
                pivot_table.PivotFields(value_field),
                f'Sum of {value_field}',
                -4157
            )

        # ================================================
        # PIVOT STYLE
        # ================================================

        pivot_table.ShowTableStyleRowStripes = True

        pivot_table.RowAxisLayout(1)

        pivot_table.RepeatAllLabels(2)

    # =====================================================
    # P&L PIVOT
    # =====================================================

    create_pivot(
        'MICA_View_PL',
        rows=[
            'Level1_mica_desc',
            'Level3_mica_desc',
            'Level8_mica_desc',
            'Level9_mica_desc'
        ],
        columns=['Source_sys'],
        values=['P&L']
    )

    # =====================================================
    # BS PIVOT
    # =====================================================

    create_pivot(
        'MICA_View_BS',
        rows=[
            'Level1_mica_desc',
            'Level2_mica_desc',
            'Level3_mica_desc'
        ],
        columns=['Source_sys'],
        values=['BS']
    )

    # =====================================================
    # AVB PIVOT
    # =====================================================

    create_pivot(
        'MICA_View_AVB',
        rows=[
            'Level1_mica_desc',
            'Level2_mica_desc',
            'Level3_mica_desc'
        ],
        columns=['Source_sys'],
        values=['AVB']
    )

    # =====================================================
    # MI FUNCTION PIVOT
    # =====================================================

    create_pivot(
        'MI_Func_RTNs',
        rows=[
            'Consolidated Period Mi Function Code',
            'Function Leaf Description',
            'Function Level 3',
            'Function Description'
        ],
        columns=['Source_sys'],
        values=['AVB', 'BS', 'P&L']
    )

    # =====================================================
    # ENTITY PIVOT
    # =====================================================

    create_pivot(
        'Entity_View',
        rows=[
            'Consolidated Period Entity ID'
        ],
        columns=['Source_sys'],
        values=['AVB', 'BS', 'P&L']
    )

    # =====================================================
    # SAVE
    # =====================================================

    wb.Save()

    wb.Close()

    excel.Quit()

    print(f'Created : {output_file}')

print('All business files generated successfully.')
