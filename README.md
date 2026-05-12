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


def drop_all_zero_rows_pivot(df):

    numeric_df = df.select_dtypes(
        include='number'
    )

    return df[
        ~(numeric_df
          .fillna(0)
          .eq(0)
          .all(axis=1))
    ]


# =========================================================
# ADD SUBTOTALS + GRAND TOTAL
# =========================================================

def add_level_subtotals(df, subtotal_level=1):

    numeric_cols = df.select_dtypes(
        include='number'
    ).columns

    result = []

    grouped = df.groupby(
        level=list(range(subtotal_level + 1)),
        sort=False
    )

    for keys, grp in grouped:

        result.append(grp)

        subtotal_vals = grp[numeric_cols].sum()

        if not isinstance(keys, tuple):
            keys = (keys,)

        subtotal_index = list(keys)

        while len(subtotal_index) < len(df.index.names):
            subtotal_index.append('')

        subtotal_index[-1] = 'Subtotal'

        subtotal_df = pd.DataFrame(
            [subtotal_vals],
            index=pd.MultiIndex.from_tuples(
                [tuple(subtotal_index)],
                names=df.index.names
            )
        )

        result.append(subtotal_df)

    final_df = pd.concat(result)

    # =====================================================
    # GRAND TOTAL
    # =====================================================

    grand_total = final_df[numeric_cols].sum()

    grand_index = ['Grand Total']

    while len(grand_index) < len(df.index.names):
        grand_index.append('')

    grand_total_df = pd.DataFrame(
        [grand_total],
        index=pd.MultiIndex.from_tuples(
            [tuple(grand_index)],
            names=df.index.names
        )
    )

    final_df = pd.concat(
        [final_df, grand_total_df]
    )

    return final_df


# =========================================================
# HEADER COLORING
# =========================================================

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


# =========================================================
# LEFT ALIGN INDEX COLUMNS
# =========================================================

def left_align_index_columns(ws):

    if ws.title == 'Raw_Data':
        return

    for col in range(1, 6):

        for row in range(1, ws.max_row + 1):

            ws.cell(
                row=row,
                column=col
            ).alignment = LEFT_ALIGN


# =========================================================
# HIGHLIGHT SUBTOTALS
# =========================================================

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
    # GENERIC PIVOT FUNCTION
    # =====================================================

    def build_view(
        source_df,
        index_cols,
        value_cols
    ):

        pivot_df = source_df.pivot_table(
            index=index_cols,
            columns='Source_sys',
            values=value_cols,
            aggfunc='sum',
            fill_value=0
        )

        for sec in value_cols:

            pivot_df[(f'{sec}_var', 'BFA_vs_CVUK')] = (
                pivot_df.get((sec, 'BFA'), 0)
                + pivot_df.get((sec, 'CVUK'), 0)
            )

            pivot_df[(f'{sec}_var', 'BFA_vs_GRC')] = (
                pivot_df.get((sec, 'BFA'), 0)
                + pivot_df.get((sec, 'GRC'), 0)
            )

            pivot_df[(f'{sec}_var', 'GRC_vs_CVUK')] = (
                pivot_df.get((sec, 'GRC'), 0)
                - pivot_df.get((sec, 'CVUK'), 0)
            )

        pivot_df = drop_all_zero_rows_pivot(
            pivot_df
        )

        subtotal_level = min(
            1,
            len(index_cols) - 1
        )

        pivot_df = add_level_subtotals(
            pivot_df,
            subtotal_level=subtotal_level
        )

        return pivot_df

    # =====================================================
    # P&L VIEW
    # =====================================================

    df4_pl = generated_df[
        generated_df['MICA Leaf'].str.startswith(
            'MP',
            na=False
        )
    ]

    mica_view_pl = build_view(
        df4_pl,
        [
            'Level1_mica_desc',
            'Level3_mica_desc',
            'Level8_mica_desc',
            'Level9_mica_desc'
        ],
        ['P&L']
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

    mica_view_bs = build_view(
        df4_bs,
        [
            'Level1_mica_desc',
            'Level2_mica_desc',
            'Level3_mica_desc'
        ],
        ['BS']
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

    mica_view_avb = build_view(
        df4_avb,
        [
            'Level1_mica_desc',
            'Level2_mica_desc',
            'Level3_mica_desc'
        ],
        ['AVB']
    )

    # =====================================================
    # MI FUNCTION VIEW
    # =====================================================

    mifunc_view = build_view(
        generated_df,
        [
            'Consolidated Period Mi Function Code',
            'Function Leaf Description',
            'Function Level 3',
            'Function Description'
        ],
        ['AVB', 'BS', 'P&L']
    )

    # =====================================================
    # ENTITY VIEW
    # =====================================================

    entity_view = build_view(
        generated_df,
        [
            'Consolidated Period Entity ID'
        ],
        ['AVB', 'BS', 'P&L']
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

        mifunc_view.to_excel(
            writer,
            sheet_name='MI_Func_RTNs',
            index=True
        )

        entity_view.to_excel(
            writer,
            sheet_name='Entity_View',
            index=True
        )

        # =================================================
        # FORMAT ALL SHEETS
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
