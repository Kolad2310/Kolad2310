num_cols = df.columns.difference(['Category'])

net_row = (
    df.loc[df['Category'] == 'Total', num_cols].iloc[0]
    - df.loc[df['Category'] == 'Top', num_cols].iloc[0]
)

df.loc[len(df)] = ['Net'] + net_row.tolist()|
