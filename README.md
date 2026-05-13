```
# =========================================================
# FORMAT FUNCTION
# =========================================================

def format_value(value, label):

    # LABELS WHICH ARE COUNTS (NO $ / NO m)
    count_labels = [
        'IWPB Client Referrals'
    ]

    # WHOLE NUMBER FORMAT
    if label in count_labels:
        return f"{value:,.0f}"

    # AMOUNT FORMAT
    sign = "-" if value < 0 else ""

    return f"{sign}${abs(value):,.1f}m"
