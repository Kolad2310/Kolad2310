```
sections = ['AVB', 'BS', 'P&L']

for sec in sections:
    mica_view[(f'{sec}_VAR', 'BFA_vs_CVUK')] = (
        mica_view[(sec, 'BFA')] - mica_view[(sec, 'CVUK')]
    )

    mica_view[(f'{sec}_VAR', 'BFA_vs_GRC')] = (
        mica_view[(sec, 'BFA')] - mica_view[(sec, 'GRC')]
    )
