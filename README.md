```
def update_column_by_prefix(
    df,
    prefix,
    target_col,
    source_col_name,
    check_col='MICA Leaf'
):
    mask = df[check_col].str.startswith(prefix, na=False)

    df.loc[mask, target_col] = df.loc[mask, source_col_name]

    return df


df = update_column_by_prefix(df, 'MP', 'P&L', 'pl_col')

df = update_column_by_prefix(df, 'MB', 'BS', 'bs_col')

df = update_column_by_prefix(df, 'AV', 'AVB', 'avb_col')
