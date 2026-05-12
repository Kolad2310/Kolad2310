```
import pandas as pd
import numpy as np
import os
from datetime import datetime

from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment
)

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
# HEADER COLORS
# =========================================================

RAW_HEADER_FILL = PatternFill(
    start_color='BFBFBF',
    end_color='BFBFBF',
    fill_type='solid'
)

PL_HEADER_FILL = PatternFill(
    start_color='F4CCCC',
    end_color='F4CCCC',
    fill_type='solid'
)

BS_HEADER_FILL = PatternFill(
    start_color='CFE2F3',
    end_color='CFE2F3',
    fill_type='solid'
)

AVB_HEADER_FILL = PatternFill(
    start_color='D9EAD3',
    end_color='D9EAD3',
    fill_type='solid'
)

SUBTOTAL_FILL = PatternFill(
    start_color='FFF2CC',
    end_color='FFF2CC',
    fill_type='solid'
)

HEADER_FONT = Font(
    bold=True
)

LEFT_ALIGN = Alignment(
    horizontal='left'
)

# =========================================================
# FUNCTIONS
# =========================================================

def auto_adjust_column_width(ws):

    for column_cells in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            try:

                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            except:
                pass

        adjusted_width = min(
            max_length + 3,
            60
        )

        ws.column_dimensions[
            column_letter
        ].width = adjusted_width


def apply_number_format(ws):

    for row in ws.iter_rows():

        for cell in row:

            if isinstance(
                cell.value,
                (
                    int,
                    float,
                    np.integer,
                    np.floating
                )
            ):

                cell.number_format = '#,##0'


def drop_all_zero_rows(df):

    numeric_cols = df.select_dtypes(
        include='number'
    ).columns

    df = df[
        ~(df[numeric_cols]
          .fillna(0)
          .eq(0)
          .all(axis=1))
    ]

    return df


# =========================================================
# ADD LEVEL 2 SUBTOTALS
# =========================================================

def add_level2_subtotals(df):

    numeric_cols = df.select_dtypes(
        include='number'
    ).columns

    result = []

    # ---------------------------------------------
    # if multi index exists
    # ---------------------------------------------

    if isinstance(df.index, pd.MultiIndex):

        level_names = df.index.names

        # subtotal at level 2
        level2_groups = df.groupby(
            level=[0, 1],
            sort=False
        )

        for keys, grp in level2_groups:

            result.append(grp)

            subtotal = grp[numeric_cols].sum()

            subtotal_index = list(keys)

            while len(subtotal_index) < len(level_names):
                subtotal_index.append('')

            subtotal_index[-1] = 'Subtotal'

            subtotal_df = pd.DataFrame(
                [subtotal],
                index=pd.MultiIndex.from_tuples(
                    [tuple(subtotal_index)],
                    names=level_names
                )
            )

            result.append(subtotal_df)

        final_df = pd.concat(result)

    else:

        final_df = df.copy()

    # ---------------------------------------------
    # grand total
    # ---------------------------------------------

    grand_total = final_df[numeric_cols].sum()

    grand_index = ['Grand Total']

    if isinstance(final_df.index, pd.MultiIndex):

        while len(grand_index) < len(final_df.index.names):
            grand_index.append('')

        grand_total_df = pd.DataFrame(
            [grand_total],
            index=pd.MultiIndex.from_tuples(
                [tuple(grand_index)],
                names=final_df.index.names
            )
        )

    else:

        grand_total_df = pd.DataFrame(
            [grand_total],
            index=['Grand Total']
        )

    final_df = pd.concat(
        [final_df, grand_total_df]
    )

    return final_df


def color_headers(ws):

    # =====================================================
    # RAW DATA
    # =====================================================

    if ws.title == 'Raw_Data':

        for cell in ws[1]:

            cell.fill = RAW_HEADER_FILL
            cell.font = HEADER_FONT

        return

    # =====================================================
    # PIVOT HEADERS
    # =====================================================

    for row in [1, 2]:

        for cell in ws[row]:

            value = str(cell.value)

            if 'P&L' in value:

                cell.fill = PL_HEADER_FILL

            elif 'BS' in value:

                cell.fill = BS_HEADER_FILL

            elif 'AVB' in value:

                cell.fill = AVB_HEADER_FILL

            else:

                cell.fill = RAW_HEADER_FILL

            cell.font = HEADER_FONT


def left_align_index_columns(ws):

    if ws.title == 'Raw_Data':
        return

    for col in range(1, 6):

        for row in range(1, ws.max_row + 1):

            ws.cell(
                row=row,
                column=col
            ).alignment = LEFT_ALIGN


def highlight_subtotals(ws):

    for row in ws.iter_rows():

        for cell in row:

            if cell.value == 'Subtotal':

                for subtotal_cell in row:

                    subtotal_cell.fill = SUBTOTAL_FILL
                    subtotal_cell.font = HEADER_FONT

            if cell.value == 'Grand Total':

                for total_cell in row:

                    total_cell.fill = SUBTOTAL_FILL
                    total_cell.font = HEADER_FONT


# =========================================================
# LOOP EACH BUSINESS
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
    # GENERATED DF
    # =====================================================

    generated_df = pd.concat(
        all_parts,
        ignore_index=True
    ).drop_duplicates()

    # =====================================================
    # DESCRIPTION COLUMNS
    # =====================================================

    generated_df['Level1_mica_desc'] = (
        generated_df['Level 1_mica'].astype(str)
        + ' '
        + generated_df['Description 1_mica'].astype(str)
    )

    generated_df['Level2_mica_desc'] = (
        generated_df['Level 2_mica'].astype(str)
        + ' '
        + generated_df['Description 2_mica'].astype(str)
    )

    generated_df['Level3_mica_desc'] = (
        generated_df['Level 3_mica'].astype(str)
        + ' '
        + generated_df['Description 3_mica'].astype(str)
    )

    generated_df['Level8_mica_desc'] = (
        generated_df['Level 8_mica'].astype(str)
        + ' '
        + generated_df['Description 8_mica'].astype(str)
    )

    generated_df['Level9_mica_desc'] = (
        generated_df['Level 9_mica'].astype(str)
        + ' '
        + generated_df['Description 9_mica'].astype(str)
    )

    # =====================================================
    # P&L VIEW
    # =====================================================

    df4_pl = generated_df[
        generated_df['MICA Leaf'].str.startswith(
            'MP',
            na=False
        )
    ]

    mica_view_pl = df4_pl.pivot_table(
        index=[
            'Level1_mica_desc',
            'Level3_mica_desc',
            'Level8_mica_desc',
            'Level9_mica_desc'
        ],
        columns='Source_sys',
        values=['P&L'],
        aggfunc='sum',
        fill_value=0
    )

    mica_view_pl[('P&L_var', 'BFA_vs_CVUK')] = (
        mica_view_pl.get(('P&L', 'BFA'), 0)
        + mica_view_pl.get(('P&L', 'CVUK'), 0)
    )

    mica_view_pl[('P&L_var', 'BFA_vs_GRC')] = (
        mica_view_pl.get(('P&L', 'BFA'), 0)
        + mica_view_pl.get(('P&L', 'GRC'), 0)
    )

    mica_view_pl[('P&L_var', 'GRC_vs_CVUK')] = (
        mica_view_pl.get(('P&L', 'GRC'), 0)
        - mica_view_pl.get(('P&L', 'CVUK'), 0)
    )

    mica_view_pl = drop_all_zero_rows(
        mica_view_pl
    )

    mica_view_pl = add_level2_subtotals(
        mica_view_pl
    )

    # =====================================================
    # BS VIEW
    # =====================================================

    df4_bs = generated_df[
        generated_df['MICA Leaf'].str.startswith(
            'MB',
            na=False
        )
    ]

    mica_view_bs = df4_bs.pivot_table(
        index=[
            'Level1_mica_desc',
            'Level2_mica_desc',
            'Level3_mica_desc'
        ],
        columns='Source_sys',
        values=['BS'],
        aggfunc='sum',
        fill_value=0
    )

    mica_view_bs[('BS_var', 'BFA_vs_CVUK')] = (
        mica_view_bs.get(('BS', 'BFA'), 0)
        + mica_view_bs.get(('BS', 'CVUK'), 0)
    )

    mica_view_bs[('BS_var', 'BFA_vs_GRC')] = (
        mica_view_bs.get(('BS', 'BFA'), 0)
        + mica_view_bs.get(('BS', 'GRC'), 0)
    )

    mica_view_bs[('BS_var', 'GRC_vs_CVUK')] = (
        mica_view_bs.get(('BS', 'GRC'), 0)
        - mica_view_bs.get(('BS', 'CVUK'), 0)
    )

    mica_view_bs = drop_all_zero_rows(
        mica_view_bs
    )

    mica_view_bs = add_level2_subtotals(
        mica_view_bs
    )

    # =====================================================
    # AVB VIEW
    # =====================================================

    df4_avb = generated_df[
        generated_df['MICA Leaf'].str.startswith(
            'AV',
            na=False
        )
    ]

    mica_view_avb = df4_avb.pivot_table(
        index=[
            'Level1_mica_desc',
            'Level2_mica_desc',
            'Level3_mica_desc'
        ],
        columns='Source_sys',
        values=['AVB'],
        aggfunc='sum',
        fill_value=0
    )

    mica_view_avb[('AVB_var', 'BFA_vs_CVUK')] = (
        mica_view_avb.get(('AVB', 'BFA'), 0)
        + mica_view_avb.get(('AVB', 'CVUK'), 0)
    )

    mica_view_avb[('AVB_var', 'BFA_vs_GRC')] = (
        mica_view_avb.get(('AVB', 'BFA'), 0)
        + mica_view_avb.get(('AVB', 'GRC'), 0)
    )

    mica_view_avb[('AVB_var', 'GRC_vs_CVUK')] = (
        mica_view_avb.get(('AVB', 'GRC'), 0)
        - mica_view_avb.get(('AVB', 'CVUK'), 0)
    )

    mica_view_avb = drop_all_zero_rows(
        mica_view_avb
    )

    mica_view_avb = add_level2_subtotals(
        mica_view_avb
    )

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
        engine='openpyxl'
    ) as writer:

        generated_df.to_excel(
            writer,
            sheet_name='Raw_Data',
            index=False
        )

        mica_view_pl.to_excel(
            writer,
            sheet_name='MICA_View_PL',
            index=True
        )

        mica_view_bs.to_excel(
            writer,
            sheet_name='MICA_View_BS',
            index=True
        )

        mica_view_avb.to_excel(
            writer,
            sheet_name='MICA_View_AVB',
            index=True
        )

        # =================================================
        # FORMAT SHEETS
        # =================================================

        for sheet in writer.book.sheetnames:

            ws = writer.book[sheet]

            auto_adjust_column_width(ws)

            apply_number_format(ws)

            color_headers(ws)

            left_align_index_columns(ws)

            highlight_subtotals(ws)

    print(f'Created : {output_file}')

print('All business files generated successfully.')
