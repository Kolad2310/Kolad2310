```
# Customer Number exception (ONLY alphabets or blank)

if "Customer Number" in df.columns:

    cust_series = df["Customer Number"].astype(str).str.strip()

    cust_mask = (
        cust_series.eq("") |                          # blank
        cust_series.str.lower().eq("none") |          # string "None"
        cust_series.str.contains(r"[a-zA-Z]", na=False)  # contains alphabets
    )

    df.loc[cust_mask, "Exception_Reason"] += "Invalid Customer; "
