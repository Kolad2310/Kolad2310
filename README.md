```
# =====================================================
# CREATE TAG COLUMN
# =====================================================

df4['Tag'] = ''

# =====================================================
# GENERATE DIFFERENT DATAFRAMES
# AND UPDATE TAG COLUMN
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

            # --------------------------------------------
            # START NEW DF WHEN CG COMES
            # --------------------------------------------

            if str(filter_val).startswith('CG'):

                if current_df is not None:

                    all_parts.append(current_df)

                mask = (
                    df4[filter_col]
                    .astype(str)
                    == str(filter_val)
                )

                current_df = df4[
                    mask
                ].copy()

                # =========================================
                # UPDATE TAG
                # =========================================

                df4.loc[
                    mask,
                    'Tag'
                ] = np.where(
                    df4.loc[
                        mask,
                        'Tag'
                    ].astype(str) == '',
                    str(business),
                    df4.loc[
                        mask,
                        'Tag'
                    ].astype(str)
                    + ', '
                    + str(business)
                )

            else:

                if current_df is not None:

                    sub_mask = (
                        current_df[filter_col]
                        .astype(str)
                        == str(filter_val)
                    )

                    current_df = current_df[
                        sub_mask
                    ]

                    # =====================================
                    # UPDATE TAG IN ORIGINAL DF
                    # =====================================

                    original_mask = (
                        df4.index.isin(current_df.index)
                    )

                    df4.loc[
                        original_mask,
                        'Tag'
                    ] = np.where(
                        df4.loc[
                            original_mask,
                            'Tag'
                        ].astype(str) == '',
                        str(business),
                        df4.loc[
                            original_mask,
                            'Tag'
                        ].astype(str)
                        + ', '
                        + str(business)
                    )

        if current_df is not None:

            all_parts.append(current_df)

    # =================================================
    # CASE 2 : NO CG EXISTS
    # =================================================

    else:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']

            filter_val = row['Value']

            mask = (
                df4[filter_col]
                .astype(str)
                == str(filter_val)
            )

            temp_df = df4[
                mask
            ].copy()

            all_parts.append(temp_df)

            # =========================================
            # UPDATE TAG
            # =========================================

            df4.loc[
                mask,
                'Tag'
            ] = np.where(
                df4.loc[
                    mask,
                    'Tag'
                ].astype(str) == '',
                str(business),
                df4.loc[
                    mask,
                    'Tag'
                ].astype(str)
                + ', '
                + str(business)
            )

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
# REMOVE DUPLICATE BUSINESS TAGS
# =====================================================

df4['Tag'] = (
    df4['Tag']
    .astype(str)
    .apply(
        lambda x: ', '.join(
            dict.fromkeys(
                [
                    i.strip()
                    for i in x.split(',')
                    if i.strip()
                ]
            )
        )
    )
)

# =====================================================
# PRINT GENERATED DATAFRAME NAMES
# =====================================================

print('Generated DataFrames:')

for name in generated_df_names:

    print(name)
