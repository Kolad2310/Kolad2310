```
import pandas as pd
import numpy as np
import os
from datetime import datetime

# =====================================================
# CREATE TAG COLUMNS
# =====================================================

df4['Tag'] = np.nan
df4['tag_npr'] = np.nan

# =====================================================
# OUTPUT FOLDER
# =====================================================

timestamp = datetime.now().strftime('%d%b_%H%M')

output_folder = f'Business_Output_{timestamp}'

os.makedirs(
    output_folder,
    exist_ok=True
)

# =====================================================
# STORE GENERATED DATAFRAME NAMES
# =====================================================

generated_df_names = []

# =====================================================
# LOOP BUSINESS
# =====================================================

for business in ref_df['Business'].dropna().unique():

    print(f'\nProcessing : {business}')

    # =================================================
    # FILTER REFERENCE FOR BUSINESS
    # =================================================

    temp_ref = ref_df[
        ref_df['Business'] == business
    ].reset_index(drop=True)

    # =================================================
    # CHECK IF CG EXISTS
    # =================================================

    has_cg = temp_ref['Value'].astype(str).str.startswith(
        'CG',
        na=False
    ).any()

    all_parts = []

    # =================================================
    # CASE 1 : CG EXISTS
    # =================================================

    if has_cg:

        cg_positions = temp_ref[
            temp_ref['Value']
            .astype(str)
            .str.startswith('CG', na=False)
        ].index.tolist()

        # ------------------------------------------------
        # LOOP CG BLOCKS
        # ------------------------------------------------

        for i, start_idx in enumerate(cg_positions):

            # =============================================
            # BLOCK END
            # =============================================

            if i < len(cg_positions) - 1:

                end_idx = cg_positions[i + 1]

            else:

                end_idx = len(temp_ref)

            # =============================================
            # CURRENT BLOCK
            # =============================================

            block_df = temp_ref.iloc[
                start_idx:end_idx
            ].reset_index(drop=True)

            # =============================================
            # FIRST FILTER = CG
            # =============================================

            first_row = block_df.iloc[0]

            filter_col = first_row['Filter Column']

            filter_val = str(first_row['Value'])

            cg_filtered_df = df4[
                df4[filter_col]
                .astype(str)
                == filter_val
            ].copy()

            # =============================================
            # ONLY CG EXISTS
            # =============================================

            if len(block_df) == 1:

                all_parts.append(cg_filtered_df)

            # =============================================
            # APPLY RTN / PR FILTERS
            # =============================================

            else:

                for j in range(1, len(block_df)):

                    row = block_df.iloc[j]

                    sub_filter_col = row['Filter Column']

                    sub_filter_val = str(row['Value'])

                    temp_filtered = cg_filtered_df[
                        cg_filtered_df[sub_filter_col]
                        .astype(str)
                        == sub_filter_val
                    ].copy()

                    all_parts.append(temp_filtered)

    # =================================================
    # CASE 2 : NO CG EXISTS
    # =================================================

    else:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = str(row['Value'])

            temp_filtered = df4[
                df4[filter_col]
                .astype(str)
                == filter_val
            ].copy()

            all_parts.append(temp_filtered)

    # =================================================
    # FINAL GENERATED DF
    # =================================================

    if len(all_parts) > 0:

        generated_df = pd.concat(
            all_parts,
            ignore_index=False
        ).drop_duplicates()

    else:

        generated_df = pd.DataFrame()

    # =================================================
    # NPR ADDITION
    # =================================================

    has_cg01 = temp_ref[
        'Value'
    ].astype(str).str.contains(
        'CG01',
        na=False
    ).any()

    if has_cg01:

        # -----------------------------------------------
        # GET RTNs UNDER CURRENT BUSINESS
        # -----------------------------------------------

        npr_rtns = temp_ref[
            temp_ref['Value']
            .astype(str)
            .str.startswith('RTN', na=False)
        ]['Value'].astype(str).unique().tolist()

        # -----------------------------------------------
        # FILTER NPR DATA
        # -----------------------------------------------

        npr_df = df4[
            (
                df4['MI Product Leaf Describe']
                .astype(str)
                .str.contains('NPT', na=False)
            )
            &
            (
                df4['MI GLOBALBUSINESS Level 3']
                .astype(str)
                == 'CG01'
            )
            &
            (
                df4['Level6_mica_desc']
                .astype(str)
                == 'Total opex'
            )
            &
            (
                df4['MI RTN']
                .astype(str)
                .isin(npr_rtns)
            )
        ].copy()

        # -----------------------------------------------
        # APPEND NPR DATA
        # -----------------------------------------------

        generated_df = pd.concat(
            [
                generated_df,
                npr_df
            ],
            ignore_index=False
        ).drop_duplicates()

        # -----------------------------------------------
        # UPDATE NPR TAG
        # -----------------------------------------------

        df4.loc[
            npr_df.index,
            'tag_npr'
        ] = str(business) + '_NPR'

    # =================================================
    # UPDATE TAG
    # =================================================

    if len(generated_df) > 0:

        df4.loc[
            generated_df.index,
            'Tag'
        ] = str(business)

    # =================================================
    # DATAFRAME NAME
    # =================================================

    df_name = (
        str(business)
        .replace('/', '_')
        .replace('\\', '_')
        .replace(' ', '_')
        .replace('-', '_')
    )

    # =================================================
    # CREATE VARIABLE
    # =================================================

    globals()[df_name] = generated_df

    generated_df_names.append(df_name)

    # =================================================
    # WRITE TO EXCEL
    # =================================================

    output_file = os.path.join(
        output_folder,
        f'{df_name}_{timestamp}.xlsx'
    )

    with pd.ExcelWriter(
        output_file,
        engine='xlsxwriter'
    ) as writer:

        generated_df.to_excel(
            writer,
            sheet_name='Data',
            index=False
        )

        workbook = writer.book

        worksheet = writer.sheets['Data']

        # =============================================
        # HEADER FORMATS
        # =============================================

        grey_header = workbook.add_format({
            'bold': True,
            'bg_color': '#BFBFBF',
            'border': 1
        })

        blue_header = workbook.add_format({
            'bold': True,
            'bg_color': '#D9EAF7',
            'border': 1
        })

        number_format = workbook.add_format({
            'num_format': '#,##0'
        })

        # =============================================
        # BLUE HEADERS
        # =============================================

        blue_headers = [
            'Tag',
            'tag_npr',
            'Scope',
            'P&L',
            'BS',
            'AVB'
        ]

        # =============================================
        # FORMAT HEADERS
        # =============================================

        for col_num, value in enumerate(
            generated_df.columns.values
        ):

            if value in blue_headers:

                worksheet.write(
                    0,
                    col_num,
                    value,
                    blue_header
                )

            else:

                worksheet.write(
                    0,
                    col_num,
                    value,
                    grey_header
                )

        # =============================================
        # AUTO WIDTH
        # =============================================

        for idx, col in enumerate(
            generated_df.columns
        ):

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

            worksheet.set_column(
                idx,
                idx,
                min(max_len, 60)
            )

        # =============================================
        # NUMBER FORMAT
        # =============================================

        numeric_cols = generated_df.select_dtypes(
            include='number'
        ).columns

        for col in numeric_cols:

            col_idx = generated_df.columns.get_loc(
                col
            )

            worksheet.set_column(
                col_idx,
                col_idx,
                None,
                number_format
            )

    print(f'Created : {output_file}')

# =====================================================
# PRINT GENERATED DATAFRAME NAMES
# =====================================================

print('\nGenerated DataFrames:\n')

for name in generated_df_names:

    print(name)
