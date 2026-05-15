```
# =====================================================
# NPR PRODUCT DATAFRAME
# =====================================================

npr_prod_df = df4[
    (
        df4['MI Product Leaf Describe']
        .astype(str)
        .str.contains('NPT', na=False)
    )
    &
    (
        df4['MI GLOBALBUSINESS Level 3']
        .astype(str)
        == 'CG01'
    )
    &
    (
        df4['Level6_mica_desc']
        .astype(str)
        == 'Total opex'
    )
].copy()

# =====================================================
# UPDATE TAG
# =====================================================

df4.loc[
    npr_prod_df.index,
    'Tag'
] = 'npr_prod'

# =====================================================
# CHECK RESULT
# =====================================================

print(npr_prod_df.shape)
