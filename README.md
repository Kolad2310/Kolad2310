```
```
# Update P&L
df.loc[df['MICA Leaf'].str.startswith('MP', na=False), 'P&L'] = (
    df.loc[df['MICA Leaf'].str.startswith('MP', na=False), pl_col]
    .sum(axis=1)
)

# Update BS
df.loc[df['MICA Leaf'].str.startswith('MB', na=False), 'BS'] = (
    df.loc[df['MICA Leaf'].str.startswith('MB', na=False), bs_col]
    .sum(axis=1)
)

# Update AVB
df.loc[df['MICA Leaf'].str.startswith('AV', na=False), 'AVB'] = (
    df.loc[df['MICA Leaf'].str.startswith('AV', na=False), avb_col]
    .sum(axis=1)
)
