```
import pandas as pd
import numpy as np
import os
from datetime import datetime

from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

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

HEADER_FONT = Font(
    bold=True
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

        adjusted_width = min(max_length + 3, 60)

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


def add_grand_total(df):

    numeric_cols = df.select_dtypes(
        include='number'
    ).columns

    total_row = df[numeric_cols].sum()

    total_df = pd.DataFrame(
        [total_row],
        index=['Grand Total']
    )

    df = pd.concat([df, total_df])

    return df


def color_headers(ws, fill):

    # raw data
    if ws.title == 'Raw_Data':

        for cell in ws[1]:

            cell.fill = fill
            cell.font = HEADER_FONT

    else:

        # pivots
        for row in [1, 2]:

            for cell in ws[row]:

                cell.fill = fill
                cell.font = HEADER_FONT


# =========================================================
# LOOP EACH BUSINESS
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

    # =====================================================
    # CHECK IF CG EXISTS
    # =====================================================

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

    if all_parts:

        generated_df = pd.concat(
            all_parts,
            ignore_index=True
        ).drop_duplicates()

    else:

        generated_df = pd.DataFrame()

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

    for sec in ['P&L']:

        mica_view_pl[(f'{sec}_var', 'BFA_vs_CVUK')] = (
            mica_view_pl.get((sec, 'BFA'), 0)
            + mica_view_pl.get((sec, 'CVUK'), 0)
        )

        mica_view_pl[(f'{sec}_var', 'BFA_vs_GRC')] = (
            mica_view_pl.get((sec, 'BFA'), 0)
            + mica_view_pl.get((sec, 'GRC'), 0)
        )

        mica_view_pl[(f'{sec}_var', 'GRC_vs_CVUK')] = (
            mica_view_pl.get((sec, 'GRC'), 0)
            - mica_view_pl.get((sec, 'CVUK'), 0)
        )

    mica_view_pl = add_grand_total(
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

    for sec in ['BS']:

        mica_view_bs[(f'{sec}_var', 'BFA_vs_CVUK')] = (
            mica_view_bs.get((sec, 'BFA'), 0)
            + mica_view_bs.get((sec, 'CVUK'), 0)
        )

        mica_view_bs[(f'{sec}_var', 'BFA_vs_GRC')] = (
            mica_view_bs.get((sec, 'BFA'), 0)
            + mica_view_bs.get((sec, 'GRC'), 0)
        )

        mica_view_bs[(f'{sec}_var', 'GRC_vs_CVUK')] = (
            mica_view_bs.get((sec, 'GRC'), 0)
            - mica_view_bs.get((sec, 'CVUK'), 0)
        )

    mica_view_bs = add_grand_total(
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

    for sec in ['AVB']:

        mica_view_avb[(f'{sec}_var', 'BFA_vs_CVUK')] = (
            mica_view_avb.get((sec, 'BFA'), 0)
            + mica_view_avb.get((sec, 'CVUK'), 0)
        )

        mica_view_avb[(f'{sec}_var', 'BFA_vs_GRC')] = (
            mica_view_avb.get((sec, 'BFA'), 0)
            + mica_view_avb.get((sec, 'GRC'), 0)
        )

        mica_view_avb[(f'{sec}_var', 'GRC_vs_CVUK')] = (
            mica_view_avb.get((sec, 'GRC'), 0)
            - mica_view_avb.get((sec, 'CVUK'), 0)
        )

    mica_view_avb = add_grand_total(
        mica_view_avb
    )

    # =====================================================
    # MI FUNCTION VIEW
    # =====================================================

    mifunc_view = generated_df.pivot_table(
        index=[
            'Consolidated Period Mi Function Code',
            'Function Leaf Description',
            'Function Level 3',
            'Function Description'
        ],
        columns='Source_sys',
        values=['AVB', 'BS', 'P&L'],
        aggfunc='sum',
        fill_value=0
    )

    for sec in ['AVB', 'BS', 'P&L']:

        mifunc_view[(f'{sec}_var', 'BFA_vs_CVUK')] = (
            mifunc_view.get((sec, 'BFA'), 0)
            + mifunc_view.get((sec, 'CVUK'), 0)
        )

        mifunc_view[(f'{sec}_var', 'BFA_vs_GRC')] = (
            mifunc_view.get((sec, 'BFA'), 0)
            + mifunc_view.get((sec, 'GRC'), 0)
        )

        mifunc_view[(f'{sec}_var', 'GRC_vs_CVUK')] = (
            mifunc_view.get((sec, 'GRC'), 0)
            - mifunc_view.get((sec, 'CVUK'), 0)
        )

    mifunc_view = add_grand_total(
        mifunc_view
    )

    # =====================================================
    # ENTITY VIEW
    # =====================================================

    entity_view = generated_df.pivot_table(
        index=[
            'Consolidated Period Entity ID'
        ],
        columns='Source_sys',
        values=['P&L', 'BS', 'AVB'],
        aggfunc='sum',
        fill_value=0
    )

    for sec in ['BS', 'P&L', 'AVB']:

        entity_view[(f'{sec}_var', 'BFA_vs_CVUK')] = (
            entity_view.get((sec, 'BFA'), 0)
            + entity_view.get((sec, 'CVUK'), 0)
        )

        entity_view[(f'{sec}_var', 'BFA_vs_GRC')] = (
            entity_view.get((sec, 'BFA'), 0)
            + entity_view.get((sec, 'GRC'), 0)
        )

        entity_view[(f'{sec}_var', 'GRC_vs_CVUK')] = (
            entity_view.get((sec, 'GRC'), 0)
            - entity_view.get((sec, 'CVUK'), 0)
        )

    entity_view = add_grand_total(
        entity_view
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
            sheet_name='MICA_View_PL'
        )

        mica_view_bs.to_excel(
            writer,
            sheet_name='MICA_View_BS'
        )

        mica_view_avb.to_excel(
            writer,
            sheet_name='MICA_View_AVB'
        )

        mifunc_view.to_excel(
            writer,
            sheet_name='MI_Func_RTNs'
        )

        entity_view.to_excel(
            writer,
            sheet_name='Entity_View'
        )

        # =================================================
        # FORMAT ALL SHEETS
        # =================================================

        for sheet in writer.book.sheetnames:

            ws = writer.book[sheet]

            auto_adjust_column_width(ws)

            apply_number_format(ws)

            # =============================================
            # HEADER COLORS
            # =============================================

            if sheet == 'Raw_Data':

                color_headers(
                    ws,
                    RAW_HEADER_FILL
                )

            elif 'PL' in sheet:

                color_headers(
                    ws,
                    PL_HEADER_FILL
                )

            elif 'BS' in sheet:

                color_headers(
                    ws,
                    BS_HEADER_FILL
                )

            elif 'AVB' in sheet:

                color_headers(
                    ws,
                    AVB_HEADER_FILL
                )

            else:

                color_headers(
                    ws,
                    RAW_HEADER_FILL
                )

    print(f'Created : {output_file}')

print('All business files generated successfully.')
