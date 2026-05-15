```
# =============================================
# HEADER FORMATS
# =============================================

grey_header = workbook.add_format({
    'bold': True,
    'bg_color': '#BFBFBF',
    'border': 1
})

blue_header = workbook.add_format({
    'bold': True,
    'bg_color': '#D9EAF7',
    'border': 1
})

number_format = workbook.add_format({
    'num_format': '#,##0'
})

# =============================================
# HEADERS TO COLOUR BLUE
# =============================================

blue_headers = [
    'Tag',
    'Scope',
    'P&L',
    'BS',
    'AVB'
]

# =============================================
# FORMAT HEADERS
# =============================================

for col_num, value in enumerate(
    generated_df.columns.values
):

    # -----------------------------------------
    # BLUE HEADERS
    # -----------------------------------------

    if value in blue_headers:

        worksheet.write(
            0,
            col_num,
            value,
            blue_header
        )

    # -----------------------------------------
    # DEFAULT GREY
    # -----------------------------------------

    else:

        worksheet.write(
            0,
            col_num,
            value,
            grey_header
        )
