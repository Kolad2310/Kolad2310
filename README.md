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
# OUTPUT FILE
# =========================================================

timestamp = datetime.now().strftime('%d%b_%H%M')

output_file = f'Full_Product_View_{timestamp}.xlsx'

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
# GENERIC PIVOT FUNCTION
# =========================================================

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

    # =====================================================
    # VARIANCES
    # =====================================================

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

    return pivot_df


# =========================================================
# P&L VIEW
# =========================================================

df4_pl = df4[
    df4['MICA Leaf'].str.startswith(
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

# =========================================================
# BS VIEW
# =========================================================

df4_bs = df4[
    df4['MICA Leaf'].str.startswith(
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

# =========================================================
# AVB VIEW
# =========================================================

df4_avb = df4[
    df4['MICA Leaf'].str.startswith(
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

# =========================================================
# MI FUNCTION VIEW
# =========================================================

mifunc_view = build_view(
    df4,
    [
        'Consolidated Period Mi Function Code',
        'Function Leaf Description',
        'Function Level 3',
        'Function Description'
    ],
    ['AVB', 'BS', 'P&L']
)

# =========================================================
# ENTITY VIEW
# =========================================================

entity_view = build_view(
    df4,
    [
        'Consolidated Period Entity ID'
    ],
    ['AVB', 'BS', 'P&L']
)

# =========================================================
# WRITE EXCEL
# =========================================================

with pd.ExcelWriter(
    output_file,
    engine='openpyxl'
) as writer:

    df4.to_excel(
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

    # =====================================================
    # FORMAT SHEETS
    # =====================================================

    for sheet in writer.book.sheetnames:

        ws = writer.book[sheet]

        auto_adjust_column_width(ws)

        apply_number_format(ws)

        color_headers(ws)

        left_align_index_columns(ws)

print(f'Created : {output_file}')
