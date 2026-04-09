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
            if key in str(col_name):
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
        type_choice = None

        for sheet in sheets:

            b5, e5 = read_metadata(file, sheet)
            df = read_excel_safe(file, sheet)

            if df is None or df.empty:
                continue

            # 🔥 DROP ALL-NA COLUMNS
            df = df.dropna(axis=1, how="all")

            # USD conversion
            if str(e5).strip().upper() == "USD":
                df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce") / usd_rate

            # PRODUCT
            if "Product" in df.columns:
                df["Product"] = df["Product"].fillna("").astype(str).str.strip()

                invalid = ~df["Product"].str.upper().str.startswith("MD")

                if invalid.any():
                    if pd.notna(b5) and str(b5).strip():
                        rep = str(b5).strip()
                    else:
                        if not product_choice:
                            product_choice = dropdown_popup("Select Product Code", PRODUCT_CODE_OPTIONS, file)
                        rep = product_choice

                    df.loc[invalid, "Product"] = rep

            # TYPE
            if "Type" in df.columns:
                df["Type"] = df["Type"].fillna("").astype(str).str.strip()

                invalid = ~df["Type"].isin(VAL_TYPE_OPTIONS)

                if invalid.any():
                    if not type_choice:
                        type_choice = dropdown_popup("Select Type", VAL_TYPE_OPTIONS, file)
                    df.loc[invalid, "Type"] = type_choice

            # EXCEPTION
            df["Exception"] = ""
            df.loc[df["Amount"] == 0, "Exception"] = "Zero Amount"

            clean = df[df["Exception"] == ""]
            exc = df[df["Exception"] != ""]

            clean_all.append(clean)
            exc_all.append(exc)

            grp = df.groupby("Product")["Amount"].sum().reset_index()

            for _, r in grp.iterrows():
                recon.append({
                    "File": os.path.basename(file),
                    "Product": r["Product"],
                    "Input Total": r["Amount"],
                    "UKMR Submission": clean[clean["Product"] == r["Product"]]["Amount"].sum(),
                    "Exception Total": exc[exc["Product"] == r["Product"]]["Amount"].sum(),
                    "Check": 0
                })

    output = os.path.join(folder, "Output.xlsx")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.concat(clean_all).to_excel(writer, "Clean", index=False)
        pd.concat(exc_all).to_excel(writer, "Exceptions", index=False)
        pd.DataFrame(recon).to_excel(writer, "Reconciliation", index=False)

    format_recon(output)

    messagebox.showinfo("Done", f"Saved at {output}")

# ---------------- GUI ----------------
def browse():
    folder = filedialog.askdirectory()
    if not folder:
        return

    root.withdraw()  # 🔥 CLOSE FIRST WINDOW

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
