```
# =====================================================
# CREATE TAG COLUMN
# =====================================================

df4['Tag'] = ''

# =====================================================
# GENERATE DATAFRAMES
# =====================================================

generated_df_names = []

for business in ref_df['Business'].dropna().unique():

    print(f'Processing : {business}')

    temp_ref = ref_df[
        ref_df['Business'] == business
    ].reset_index(drop=True)

    current_df = None

    has_cg = temp_ref['Value'].astype(str).str.startswith(
        'CG',
        na=False
    ).any()

    # =================================================
    # CASE 1 : CG EXISTS
    # =================================================

    if has_cg:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = row['Value']

            # --------------------------------------------
            # FIRST FILTER
            # --------------------------------------------

            if current_df is None:

                current_df = df4[
                    df4[filter_col].astype(str)
                    == str(filter_val)
                ].copy()

            # --------------------------------------------
            # SUBSEQUENT FILTERS
            # --------------------------------------------

            else:

                current_df = current_df[
                    current_df[filter_col].astype(str)
                    == str(filter_val)
                ]

    # =================================================
    # CASE 2 : NO CG EXISTS
    # =================================================

    else:

        temp_parts = []

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = row['Value']

            temp_df = df4[
                df4[filter_col].astype(str)
                == str(filter_val)
            ].copy()

            temp_parts.append(temp_df)

        current_df = pd.concat(
            temp_parts,
            ignore_index=False
        ).drop_duplicates()

    # =================================================
    # FINAL GENERATED DF
    # =================================================

    generated_df = current_df.copy()

    # =================================================
    # UPDATE TAG ONLY FOR FINAL ROWS
    # =================================================

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
    # CREATE DATAFRAME VARIABLE
    # =================================================

    globals()[df_name] = generated_df

    generated_df_names.append(df_name)

# =====================================================
# PRINT GENERATED DATAFRAME NAMES
# =====================================================

print('Generated DataFrames:')

for name in generated_df_names:

    print(name)
