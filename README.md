```
import pandas as pd

# ---------------------------------------------------
# IMPORTANT FIX
# If built-in list was overwritten somewhere
# ---------------------------------------------------
try:
    del list
except:
    pass


# ---------------------------------------------------
# FUNCTION TO ADD SUBTOTALS FOR ALL INDEX LEVELS
# ---------------------------------------------------

def add_all_level_subtotals(df):

    # If not MultiIndex
    if not isinstance(df.index, pd.MultiIndex):
        return df

    nlevels = df.index.nlevels

    def recursive_subtotal(data, level=0):

        grouped = data.groupby(level=level, sort=False)

        parts = []

        for key, grp in grouped:

            # -----------------------------------
            # Process inner levels
            # -----------------------------------
            if level < nlevels - 1:

                inner_df = recursive_subtotal(
                    grp,
                    level + 1
                )

                parts.append(inner_df)

            else:
                parts.append(grp)

            # -----------------------------------
            # Create subtotal row
            # -----------------------------------
            subtotal = pd.DataFrame(grp.sum()).T

            idx = list(grp.index[0])

            # Blank lower levels
            for i in range(level + 1, nlevels):
                idx[i] = ''

            idx[level] = f'{key}_Subtotal'

            subtotal.index = pd.MultiIndex.from_tuples(
                [tuple(idx)],
                names=df.index.names
            )

            parts.append(subtotal)

        return pd.concat(parts)

    return recursive_subtotal(df)


# ---------------------------------------------------
# EXAMPLE USAGE
# ---------------------------------------------------

# Example:
# mica_view should already be pivoted dataframe
#
# mica_view = pd.pivot_table(...)

mica_view_subtotal = add_all_level_subtotals(
    mica_view
)

# Optional
mica_view_subtotal = mica_view_subtotal.fillna(0)

# Export
mica_view_subtotal.to_excel(
    'mica_view_with_subtotals.xlsx'
)

print(mica_view_subtotal)
