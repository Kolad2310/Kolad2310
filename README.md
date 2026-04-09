```
import os
import io
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import msoffcrypto
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ---------------- CONFIG ----------------
HEADER_ROW = 6
PRODUCT_CODE_OPTIONS = ["MD001", "MD002", "MD003", "MD004"]
VAL_TYPE_OPTIONS = ["Type1", "Type2", "Type3"]

password_cache = {}
product_code_cache = {}
type_cache = {}

# ---------------- PASSWORD ----------------
def decrypt_file(file):
    if file not in password_cache:
        password_cache[file] = simpledialog.askstring(
            "Password", f"Enter password for:\n{os.path.basename(file)}", show="*"
        )

    try:
        with open(file, "rb") as f:
            office = msoffcrypto.OfficeFile(f)
            office.load_key(password=password_cache[file])
            decrypted = io.BytesIO()
            office.decrypt(decrypted)
            return decrypted
    except:
        messagebox.showerror("Error", f"Wrong password for {file}")
        password_cache.pop(file, None)
        return None

# ---------------- READ ----------------
def get_excel_file(file):
    try:
        return pd.ExcelFile(file)
    except:
        dec = decrypt_file(file)
        if dec:
            return pd.ExcelFile(dec)
    return None

def read_excel_safe(file, sheet):
    try:
        return pd.read_excel(file, sheet_name=sheet, header=HEADER_ROW)
    except:
        dec = decrypt_file(file)
        if dec:
            return pd.read_excel(dec, sheet_name=sheet, header=HEADER_ROW)
    return None

def read_metadata(file, sheet):
    try:
        temp = pd.read_excel(file, sheet_name=sheet, header=None, nrows=6)
        return temp.iloc[4, 1], temp.iloc[4, 4]
    except:
        dec = decrypt_file(file)
        if dec:
            temp = pd.read_excel(dec, sheet_name=sheet, header=None, nrows=6)
            return temp.iloc[4, 1], temp.iloc[4, 4]
    return None, None

# ---------------- DROPDOWN ----------------
def dropdown_popup(title, options, file):
    popup = tk.Toplevel()
    popup.title(title)
    popup.geometry("350x150")
    popup.grab_set()

    tk.Label(popup, text=f"{title}\n{os.path.basename(file)}").pack(pady=10)

    var = tk.StringVar()
    combo = ttk.Combobox(popup, values=options, textvariable=var, state="readonly")
    combo.pack()
    combo.current(0)

    tk.Button(popup, text="OK", command=popup.destroy).pack(pady=10)
    popup.wait_window()

    return var.get()

# ---------------- FORMAT ----------------
def format_recon(path):
    wb = load_workbook(path)
    ws = wb["Reconciliation"]

    colors = {
        "Input Total": "ADD8E6",
        "UKMR Submission": "90EE90",
        "Exception Total": "FFD580",
        "Check": "D3D3D3"
    }

    headers = [c.value for c in ws[1]]

    for col_idx, col_name in enumerate(headers, 1):
        for key, color in colors.items():
            if key.lower() in str(col_name).lower():
                fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                for r in range(2, ws.max_row + 1):
                    ws.cell(row=r, column=col_idx).fill = fill

    wb.save(path)

# ---------------- MAIN ----------------
def process_files(folder, selection):
    usd_rate = simpledialog.askfloat("USD Rate", "Enter USD → GBP rate:")

    clean_all, exc_all, recon = [], [], []

    for file, sheets in selection.items():

        product_choice = None

        for sheet in sheets:

            b5, e5 = read_metadata(file, sheet)
            df = read_excel_safe(file, sheet)

            if df is None or df.empty:
                continue

            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(axis=1, how="all")

            # USD
            if str(e5).strip().upper() == "USD":
                df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce") / usd_rate

            # ---------------- PRODUCT CODE (FIXED) ----------------
            if "Product code" in df.columns:

                df["Product code"] = (
                    df["Product code"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                invalid = (
                    df["Product code"].eq("") |
                    df["Product code"].eq("0") |
                    ~df["Product code"].str.upper().str.startswith("MD")
                )

                if invalid.any():

                    if pd.notna(b5) and str(b5).strip():
                        replacement = str(b5).strip()
                    else:
                        if file not in product_code_cache:
                            product_code_cache[file] = dropdown_popup(
                                "Select Product Code",
                                PRODUCT_CODE_OPTIONS,
                                file
                            )
                        replacement = product_code_cache[file]

                    df.loc[invalid, "Product code"] = replacement

            # ---------------- TYPE ----------------
            if "Type" in df.columns:

                df["Type"] = df["Type"].fillna("").astype(str).str.strip()

                if df["Type"].eq("").all():

                    if file not in type_cache:
                        type_cache[file] = dropdown_popup(
                            "Select Type",
                            VAL_TYPE_OPTIONS,
                            file
                        )

                    df["Type"] = type_cache[file]

            # ---------------- EXCEPTION ----------------
            df["Exception"] = ""

            if "Customer No." in df.columns:
                cust = df["Customer No."].fillna("").astype(str).str.strip()

                invalid_cust = (
                    cust.eq("") |
                    cust.str.lower().eq("none") |
                    cust.str.match(r"^[A-Za-z]+$")
                )

                df.loc[invalid_cust, "Exception"] += "Invalid Customer; "

            if "Amount" in df.columns:
                df.loc[df["Amount"] == 0, "Exception"] += "Zero Amount; "

            clean = df[df["Exception"].str.strip() == ""].copy()
            exc = df[df["Exception"].str.strip() != ""].copy()

            clean_all.append(clean)
            exc_all.append(exc)

            # ---------------- RECON ----------------
            if "Product code" in df.columns and "Amount" in df.columns:

                input_grp = df.groupby("Product code")["Amount"].sum().reset_index()
                clean_grp = clean.groupby("Product code")["Amount"].sum().reset_index()
                exc_grp = exc.groupby("Product code")["Amount"].sum().reset_index()

                for _, row in input_grp.iterrows():

                    product_val = row["Product code"]
                    input_total = row["Amount"]

                    clean_total = clean_grp.loc[
                        clean_grp["Product code"] == product_val, "Amount"
                    ].sum() if not clean_grp.empty else 0

                    exc_total = exc_grp.loc[
                        exc_grp["Product code"] == product_val, "Amount"
                    ].sum() if not exc_grp.empty else 0

                    recon.append({
                        "File Name": os.path.basename(file),
                        "Product code": product_val,
                        "Input Total": input_total,
                        "UKMR Submission": clean_total,
                        "Exception Total": exc_total,
                        "Check (Should be 0)": input_total - (clean_total + exc_total)
                    })

    output = os.path.join(folder, "Output_for_SME.xlsx")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.concat(clean_all).to_excel(writer, "Clean_Data", index=False)
        pd.concat(exc_all).to_excel(writer, "Exceptions", index=False)
        pd.DataFrame(recon).to_excel(writer, "Reconciliation", index=False)

    format_recon(output)

    messagebox.showinfo("Done", f"Saved at {output}")

# ---------------- GUI ----------------
def browse():
    folder = filedialog.askdirectory()
    if not folder:
        return

    root.withdraw()

    files = [os.path.join(folder, f) for f in os.listdir(folder)
             if f.lower().endswith((".xlsx",".xls",".xlsb",".xlsm"))]

    win = tk.Toplevel()
    win.geometry("1200x600")

    selection = {}

    for file in files:
        xl = get_excel_file(file)
        if not xl:
            continue

        row = tk.Frame(win)
        row.pack(anchor="w")

        tk.Label(row, text=os.path.basename(file), width=30).pack(side="left")

        selection[file] = {}

        for sheet in xl.sheet_names:
            var = tk.BooleanVar(value=(sheet.lower()=="income sub."))
            tk.Checkbutton(row, text=sheet, variable=var).pack(side="left")
            selection[file][sheet] = var

    def submit():
        sel = {f:[s for s,v in sheets.items() if v.get()] for f,sheets in selection.items()}
        win.destroy()
        process_files(folder, sel)

    tk.Button(win, text="Submit", command=submit).pack()

# ---------------- RUN ----------------
root = tk.Tk()
root.geometry("400x200")

tk.Button(root, text="Browse Folder", command=browse).pack(pady=50)

root.mainloop()
