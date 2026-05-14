```
# cols_list = list of columns to check for strings

df5['Scope'] = np.where(
    df5[cols_list].apply(
        lambda row: any(isinstance(x, str) and x.strip() != '' for x in row),
        axis=1
    ),
    'Not good',
    'Good'
)
