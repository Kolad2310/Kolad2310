```
import os
from datetime import datetime

# =====================================================
# OUTPUT FOLDER
# =====================================================

timestamp = datetime.now().strftime('%d%b_%H%M')

output_folder = (
    f'Business_Dataframes_{timestamp}'
)

os.makedirs(
    output_folder,
    exist_ok=True
)

# =====================================================
# WRITE EACH DATAFRAME TO SEPARATE FILE
# =====================================================

for df_name in generated_df_names:

    current_df = globals()[df_name]

    output_file = os.path.join(
        output_folder,
        f'{df_name}.xlsx'
    )

    with pd.ExcelWriter(
        output_file,
        engine='xlsxwriter'
    ) as writer:

        current_df.to_excel(
            writer,
            sheet_name='Data',
            index=False
        )

        workbook = writer.book

        worksheet = writer.sheets['Data']

        # =============================================
        # HEADER FORMAT
        # =============================================

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#BFBFBF',
            'border': 1
        })

        number_format = workbook.add_format({
            'num_format': '#,##0'
        })

        # =============================================
        # FORMAT HEADERS
        # =============================================

        for col_num, value in enumerate(
            current_df.columns.values
        ):

            worksheet.write(
                0,
                col_num,
                value,
                header_format
            )

        # =============================================
        # AUTO WIDTH
        # =============================================

        for idx, col in enumerate(
            current_df.columns
        ):

            try:

                max_len = max(
                    current_df[col]
                    .astype(str)
                    .map(len)
                    .max(),
                    len(col)
                ) + 3

            except:

                max_len = len(col) + 3

            worksheet.set_column(
                idx,
                idx,
                min(max_len, 60)
            )

        # =============================================
        # NUMBER FORMAT
        # =============================================

        numeric_cols = current_df.select_dtypes(
            include='number'
        ).columns

        for col in numeric_cols:

            col_idx = current_df.columns.get_loc(
                col
            )

            worksheet.set_column(
                col_idx,
                col_idx,
                None,
                number_format
            )

    print(f'Created : {output_file}')

print('All dataframe files generated successfully.')
