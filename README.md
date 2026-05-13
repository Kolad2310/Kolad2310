```
# =========================================================
# AGGREGATE LEVEL 1
# =========================================================

level1 = df.groupby('Label', as_index=False).agg({
    'YTD_2026': 'sum',
    'YTD_Monthly_Target_2026': 'sum',
    'FY_Forecast_2026': 'sum',
    'FY_Target_2026': 'sum'
})

commentaries = []

# =========================================================
# FORMAT FUNCTION
# =========================================================

def format_value(value, label):

    # CLIENT REFERRALS -> WHOLE NUMBER
    if label == 'IWPB Client Referrals':
        return f"{value:,.0f}"

    # ALL OTHERS -> $ + m
    return f"${value:,.1f}m"


# =========================================================
# GENERATE COMMENTARY
# =========================================================

for _, row in level1.iterrows():

    label = row['Label']

    ytd_actual = row['YTD_2026']
    ytd_target = row['YTD_Monthly_Target_2026']

    fy_fcst = row['FY_Forecast_2026']
    fy_target = row['FY_Target_2026']

    # =====================================================
    # YTD VARIANCE
    # =====================================================

    ytd_var = ytd_actual - ytd_target

    ytd_pct = (
        (ytd_var / ytd_target) * 100
        if ytd_target != 0 else 0
    )

    if ytd_var >= 0:

        ytd_status = 'above'
        sort_ascending = False
        offset_word = 'partly offset by'

    else:

        ytd_status = 'below'
        sort_ascending = True
        offset_word = 'partly onset by'

    # =====================================================
    # COUNTRY DRIVERS
    # =====================================================

    temp = df[df['Label'] == label].copy()

    temp['YTD_VAR'] = (
        temp['YTD_2026']
        - temp['YTD_Monthly_Target_2026']
    )

    temp = temp.sort_values(
        'YTD_VAR',
        ascending=sort_ascending
    )

    # MAIN DRIVER
    top_country = temp.iloc[0]['Country']

    # OFFSET COUNTRY
    offset_country = ''

    if len(temp) > 1:
        offset_country = temp.iloc[-1]['Country']

    # =====================================================
    # FY COMMENTARY
    # =====================================================

    fy_var = fy_fcst - fy_target

    if round(fy_var, 2) == 0:

        fy_comment = (
            f"FYF {format_value(fy_fcst, label)} is on track."
        )

    else:

        fy_pct = (
            (fy_var / fy_target) * 100
            if fy_target != 0 else 0
        )

        fy_status = (
            'above'
            if fy_var > 0
            else 'below'
        )

        fy_comment = (
            f"FYF {format_value(fy_fcst, label)} is "
            f"{fy_status} target by "
            f"{format_value(fy_var, label)} "
            f"({fy_pct:.1f}% vs target)."
        )

    # =====================================================
    # FINAL COMMENTARY
    # =====================================================

    commentary = (
        f"{label}: "
        f"YTD target {format_value(ytd_target, label)} is "
        f"{ytd_status} target by "
        f"{format_value(ytd_var, label)} "
        f"({ytd_pct:.0f}%), "
        f"driven by {top_country}, "
        f"{offset_word} {offset_country}. "
        f"{fy_comment}"
    )

    commentaries.append(commentary)

# =========================================================
# FINAL COMMENTARY DATAFRAME
# =========================================================

final_commentary_df = pd.DataFrame({
    'Label': level1['Label'],
    'Commentary': commentaries
})

print(final_commentary_df)
