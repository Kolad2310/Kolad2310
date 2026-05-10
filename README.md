```
import pandas as pd

def add_all_level_subtotals(df):
    """
    Adds subtotal rows for all index levels in a pivot table.
    Works for MultiIndex rows.
    """

    # Ensure MultiIndex
    if not isinstance(df.index, pd.MultiIndex):
        return df

    levels = df.index.nlevels
    result = []

    def process_level(data, level_no):
        
        grouped = data.groupby(level=list(range(level_no + 1)), sort=False)

        temp = []

        for keys, grp in grouped:

            # Add actual rows
            temp.append(grp)

            # Add subtotal if not last level
            if level_no < levels - 1:

                subtotal = pd.DataFrame(grp.sum()).T

                if not isinstance(keys, tuple):
                    keys = (keys,)

                subtotal_index = list(keys)

                # Fill remaining levels
                while len(subtotal_index) < levels:
                    subtotal_index.append('')

                subtotal_index[level_no + 1] = 'Subtotal'

                subtotal.index = pd.MultiIndex.from_tuples(
                    [tuple(subtotal_index)],
                    names=df.index.names
                )

                temp.append(subtotal)

        return pd.concat(temp)

    final = df.copy()

    # Add subtotals progressively
    for lvl in reversed(range(levels - 1)):
        final = process_level(final, lvl)

    return final


# -------------------------------
# APPLY TO YOUR PIVOT VIEWS
# -------------------------------

# mica_view
mica_view_subtotal = add_all_level_subtotals(
    mica_view.set_index(
        ['Level 1_mica', 'MICA Leaf', 'Leaf Description_mica']
    )
)

# mifunc_view
mifunc_view_subtotal = add_all_level_subtotals(
    mifunc_view.set_index([
        'Consolidated Period Mi Function Code',
        'Leaf Description_mifunc',
        'Level 3_mifunc',
        'Description 3_mifunc',
        'Level 4_mifunc',
        'Description 4_mifunc'
    ])
)

# entity_view
entity_view_subtotal = add_all_level_subtotals(
    entity_view.set_index(['Consolidated Period Entity ID'])
)

# Optional
mica_view_subtotal = mica_view_subtotal.fillna(0)
mifunc_view_subtotal = mifunc_view_subtotal.fillna(0)
entity_view_subtotal = entity_view_subtotal.fillna(0)

# Export
with pd.ExcelWriter('pivot_with_subtotals.xlsx') as writer:
    mica_view_subtotal.to_excel(writer, sheet_name='MICA')
    mifunc_view_subtotal.to_excel(writer, sheet_name='MIFUNC')
    entity_view_subtotal.to_excel(writer, sheet_name='ENTITY')
