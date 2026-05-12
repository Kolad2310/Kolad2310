```
# =====================================================
# GENERATE DIFFERENT DATAFRAMES DYNAMICALLY
# =====================================================

generated_df_names = []

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

    # =================================================
    # CASE 1 : CG EXISTS
    # =================================================

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

    # =================================================
    # CASE 2 : NO CG EXISTS
    # =================================================

    else:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = row['Value']

            temp_df = df4[
                df4[filter_col].astype(str)
                == str(filter_val)
            ].copy()

            all_parts.append(temp_df)

    # =================================================
    # FINAL GENERATED DF
    # =================================================

    generated_df = pd.concat(
        all_parts,
        ignore_index=True
    ).drop_duplicates()

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
# PRINT ALL GENERATED DATAFRAME NAMES
# =====================================================

print('Generated DataFrames:')

for name in generated_df_names:

    print(name)
