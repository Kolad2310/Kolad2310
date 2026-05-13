```
# =====================================================
# CREATE TAG COLUMN
# =====================================================

df4['Tag'] = np.nan

# =====================================================
# GENERATE DATAFRAMES
# =====================================================

generated_df_names = []

for business in ref_df['Business'].dropna().unique():

    print(f'Processing : {business}')

    temp_ref = ref_df[
        ref_df['Business'] == business
    ].reset_index(drop=True)

    all_parts = []

    current_df = None

    # =================================================
    # LOOP FILTERS IN ORDER
    # =================================================

    for _, row in temp_ref.iterrows():

        filter_col = row['Filter Column']

        filter_val = str(row['Value'])

        # =============================================
        # START NEW HIERARCHY
        # =============================================

        # CG or RTN starts new drilldown

        if (
            filter_val.startswith('CG')
            or filter_val.startswith('RTN')
        ):

            # -----------------------------------------
            # APPEND PREVIOUS HIERARCHY DF
            # -----------------------------------------

            if current_df is not None:

                all_parts.append(current_df)

            # -----------------------------------------
            # START NEW DF
            # -----------------------------------------

            current_df = df4[
                df4[filter_col]
                .astype(str)
                == filter_val
            ].copy()

        # =============================================
        # DRILLDOWN FILTERS
        # =============================================

        else:

            # -----------------------------------------
            # IF NO ROOT EXISTS YET
            # -----------------------------------------

            if current_df is None:

                current_df = df4[
                    df4[filter_col]
                    .astype(str)
                    == filter_val
                ].copy()

            # -----------------------------------------
            # APPLY DRILLDOWN
            # -----------------------------------------

            else:

                current_df = current_df[
                    current_df[filter_col]
                    .astype(str)
                    == filter_val
                ]

    # =================================================
    # APPEND FINAL DF
    # =================================================

    if current_df is not None:

        all_parts.append(current_df)

    # =================================================
    # FINAL GENERATED DF
    # =================================================

    generated_df = pd.concat(
        all_parts,
        ignore_index=False
    ).drop_duplicates()

    # =================================================
    # UPDATE TAG
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
