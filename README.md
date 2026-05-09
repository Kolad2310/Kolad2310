```
def update_col_from_prefix(
    df,
    source_col,
    source_val,
    check_col,
    prefix,
    target_col,
    value_col
):
    mask = (
        (df[source_col] == source_val) &
        (df[check_col].str.startswith(prefix, na=False))
    )

    df.loc[mask, target_col] = df.loc[mask, value_col]

    return df

df = update_col_from_prefix(
    df=df,
    source_col='Source',
    source_val='BFA',
    check_col='abcd',
    prefix='AV',
    target_col='AVB',
    value_col='consol avg bal'
)
