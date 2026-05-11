```
import pandas as pd

# -----------------------------------
# Example:
# ref_df columns:
# Business | Filter Column | Value
#
# df4 = source dataframe
# -----------------------------------

generated_dfs = {}

# Loop each Business
for business in ref_df['Business'].dropna().unique():

    business_ref = ref_df[
        ref_df['Business'] == business
    ].copy()

    business_ref = business_ref.dropna(
        subset=['Filter Column', 'Value']
    )

    result_list = []

    # -----------------------------------
    # Detect CG rows
    # -----------------------------------
    cg_rows = business_ref[
        business_ref['Value']
        .astype(str)
        .str.startswith('CG', na=False)
    ]

    # -----------------------------------
    # CASE 1 : CG exists
    # -----------------------------------
    if not cg_rows.empty:

        for cg_value in cg_rows['Value'].unique():

            # Find corresponding filter column
            cg_col = cg_rows[
                cg_rows['Value'] == cg_value
            ]['Filter Column'].iloc[0]

            # Filter df4 for that CG first
            temp_df = df4[
                df4[cg_col].astype(str) == str(cg_value)
            ].copy()

            # Apply remaining filters
            remaining_filters = business_ref[
                ~(
                    (business_ref['Filter Column'] == cg_col) &
                    (business_ref['Value'] == cg_value)
                )
            ]

            for col, grp in remaining_filters.groupby('Filter Column'):

                vals = grp['Value'].astype(str).tolist()

                temp_df = temp_df[
                    temp_df[col].astype(str).isin(vals)
                ]

            result_list.append(temp_df)

    # -----------------------------------
    # CASE 2 : No CG exists
    # -----------------------------------
    else:

        temp_df = df4.copy()

        for col, grp in business_ref.groupby('Filter Column'):

            vals = grp['Value'].astype(str).tolist()

            temp_df = temp_df[
                temp_df[col].astype(str).isin(vals)
            ]

        result_list.append(temp_df)

    # -----------------------------------
    # Final dataframe for business
    # -----------------------------------
    if result_list:
        final_df = pd.concat(
            result_list,
            ignore_index=True
        ).drop_duplicates()

    else:
        final_df = pd.DataFrame()

    generated_dfs[business] = final_df

# -----------------------------------
# Example usage
# -----------------------------------

gts_df = generated_dfs.get('GTS')
mss_df = generated_dfs.get('MSS')

# Check outputs
print(gts_df.shape)
print(mss_df.shape)
