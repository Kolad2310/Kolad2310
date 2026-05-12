```
def build_view(
    source_df,
    index_cols,
    value_cols
):

    pivot_df = source_df.pivot_table(
        index=index_cols,
        columns='Source_sys',
        values=value_cols,
        aggfunc='sum',
        fill_value=0
    )

    # =============================================
    # VARIANCES
    # =============================================

    for sec in value_cols:

        pivot_df[(f'{sec}_var', 'BFA_vs_CVUK')] = (
            pivot_df.get((sec, 'BFA'), 0)
            + pivot_df.get((sec, 'CVUK'), 0)
        )

        pivot_df[(f'{sec}_var', 'BFA_vs_GRC')] = (
            pivot_df.get((sec, 'BFA'), 0)
            + pivot_df.get((sec, 'GRC'), 0)
        )

        pivot_df[(f'{sec}_var', 'GRC_vs_CVUK')] = (
            pivot_df.get((sec, 'GRC'), 0)
            - pivot_df.get((sec, 'CVUK'), 0)
        )

    # =============================================
    # DROP ALL ZERO ROWS
    # =============================================

    pivot_df = drop_all_zero_rows_pivot(
        pivot_df
    )

    # =============================================
    # SUBTOTAL LEVEL
    # =============================================

    if len(index_cols) > 1:
        subtotal_level = 1
    else:
        subtotal_level = 0

    # =============================================
    # ADD SUBTOTALS
    # =============================================

    pivot_df = add_level_subtotals(
        pivot_df,
        subtotal_level=subtotal_level
    )

    return pivot_df
