```
# =========================================================
# SCOPE COLUMN
# =========================================================

df4['Scope'] = np.where(
    (
        df4['Description 3_mica']
        .astype(str)
        .str.contains(r'\(NIP\)', na=False)
    )
    |
    (
        df4['Description 8_mica']
        .astype(str)
        .str.contains(r'\(NIP\)', na=False)
    ),
    'Out-of-scope',
    'Inscope'
)
