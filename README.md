```
generated_dfs = {}

for business in ref_df['Business'].unique():

    business_ref = ref_df[ref_df['Business'] == business]

    result_list = []

    # Identify first filter column
    first_col = business_ref.iloc[0]['Filter Column']

    # Values of first filter
    first_values = business_ref[
        business_ref['Filter Column'] == first_col
    ]['Value'].unique()

    # If first filter values contain CG-like split logic
    # process separately and append
    if any(str(v).startswith('CG') for v in first_values):

        for val in first_values:

            temp_df = df4[
                df4[first_col] == val
            ].copy()

            # Remaining filters
            remaining = business_ref[
                ~(
                    (business_ref['Filter Column'] == first_col) &
                    (business_ref['Value'] == val)
                )
            ]

            for col, grp in remaining.groupby('Filter Column'):

                vals = grp['Value'].tolist()

                temp_df = temp_df[
                    temp_df[col].isin(vals)
                ]

            result_list.append(temp_df)

    else:
        # No CG logic
        # Apply all conditions directly on df4

        temp_df = df4.copy()

        for col, grp in business_ref.groupby('Filter Column'):

            vals = grp['Value'].tolist()

            temp_df = temp_df[
                temp_df[col].isin(vals)
            ]

        result_list.append(temp_df)

    # Final dataframe
    generated_dfs[business] = (
        pd.concat(result_list, ignore_index=True)
        .drop_duplicates()
    )


# Example
gts_df = generated_dfs['GTS']
mss_df = generated_dfs['MSS']
