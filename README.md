```
import pandas as pd

# ---------------------------------------------
# ref_df columns:
# Business | Filter Column | Value
#
# df4 = source dataframe
# ---------------------------------------------

generated_dfs = {}

# Loop each business
for business in ref_df['Business'].dropna().unique():

    temp_ref = ref_df[
        ref_df['Business'] == business
    ].reset_index(drop=True)

    all_parts = []

    current_df = None

    # Check if any CG exists
    has_cg = temp_ref['Value'].astype(str).str.startswith('CG', na=False).any()

    # =====================================================
    # CASE 1 : CG EXISTS
    # =====================================================
    if has_cg:

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']
            filter_val = row['Value']

            # Start new dataframe whenever CG comes
            if str(filter_val).startswith('CG'):

                # Append previous dataframe
                if current_df is not None:
                    all_parts.append(current_df)

                # Start from df4
                current_df = df4[
                    df4[filter_col].astype(str) == str(filter_val)
                ].copy()

            else:

                # Apply filter on current CG dataframe
                if current_df is not None:

                    current_df = current_df[
                        current_df[filter_col].astype(str)
                        == str(filter_val)
                    ]

        # Append last CG dataframe
        if current_df is not None:
            all_parts.append(current_df)

    # =====================================================
    # CASE 2 : NO CG EXISTS
    # =====================================================
    else:

        # Every filter independently applied on df4
        # Then append all outputs

        for _, row in temp_ref.iterrows():

            filter_col = row['Filter Column']
            filter_val = row['Value']

            temp_df = df4[
                df4[filter_col].astype(str)
                == str(filter_val)
            ].copy()

            all_parts.append(temp_df)

    # =====================================================
    # FINAL DF
    # =====================================================
    if all_parts:

        final_df = pd.concat(
            all_parts,
            ignore_index=True
        ).drop_duplicates()

    else:
        final_df = pd.DataFrame()

    generated_dfs[business] = final_df


# ---------------------------------------------
# Example access
# ---------------------------------------------

gts_df = generated_dfs.get('GTS')
mss_df = generated_dfs.get('MSS')

print(gts_df.shape)
print(mss_df.shape)
