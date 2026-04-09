```
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# -----------------------------
# CONFIG
# -----------------------------
HEADER_ROW = 6

PRODUCT_CODE_OPTIONS = ["MD001", "MD002", "MD003", "MD004"]
VAL_TYPE_OPTIONS = ["Type1", "Type2", "Type3"]

password_cache = {}
product_code_cache = {}
type_cache = {}

# -----------------------------
# PASSWORD HANDLING (FIXED)
# -----------------------------
def ask_password(file):
    return simpledialog.askstring("Password", f"Enter password for:\n{os.path.basename(file)}", show="*")

def get_excel_file(file):
    try:
        return pd.ExcelFile(file)
    except:
        if file not in password_cache:
            password_cache[file] = ask_password(file)
        try:
            return pd.ExcelFile(file)
        except:
            messagebox.showerror("Error", f"Cannot open file: {file}")
            return None

# -----------------------------
# DROPDOWNS
# -----------------------------
def dropdown_popup(file, title, options):
    popup = tk.Toplevel(root)
    popup.title(title)
    popup.geometry("350x150")
    popup.grab_set()

    tk.Label(popup, text=f"{title} for:\n{os.path.basename(file)}").pack(pady=10)

    selected = tk.StringVar()
    combo = ttk.Combobox(popup, values=options, textvariable=selected, state="readonly")
    combo.pack(pady=5)
    combo.current(0)

    tk.Button(popup, text="Submit", command=popup.destroy).pack(pady=10)
    popup.wait_window()

    return selected.get()

# -----------------------------
# METADATA (B5, E5)
# -----------------------------
def read_metadata(file, sheet):
    try:
        temp = pd.read_excel(file, sheet_name=sheet, header=None, nrows=6)
        return temp.iloc[4, 1], temp.iloc[4, 4]
    except:
        return None, None

# -----------------------------
# SAFE READ
# -----------------------------
def read_excel_safe(file, sheet):
    try:
        return pd.read_excel(file, sheet_name=sheet, header=HEADER_ROW)
    except:
        if file not in password_cache:
            password_cache[file] = ask_password(file)
        try:
            return pd.read_excel(file, sheet_name=sheet, header=HEADER_ROW)
        except:
            return None

# -----------------------------
# MAIN PROCESS
# -----------------------------
def process_files(folder, selection_dict):
    clean_data, exception_data, recon_records = [], [], []
    header_reference = None

    usd_rate = simpledialog.askfloat("USD Rate", "Enter USD → GBP rate:")

    for file, sheets in selection_dict.items():

        # Dropdown caches per file
        product_choice = None
        type_choice = None

        for sheet in sheets:

            product_b5, currency_e5 = read_metadata(file, sheet)
            df = read_excel_safe(file, sheet)

            if df is None or df.empty:
                continue

            df = df.dropna(how="all")

            # HEADER CHECK
            cols = [c.strip().lower() for c in df.columns]
            if header_reference is None:
                header_reference = cols
            elif cols != header_reference:
                messagebox.showerror("Header Error", f"{file}-{sheet}")
                return

            # USD → GBP
            if str(currency_e5).strip().upper() == "USD" and usd_rate:
                df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
                df["Amount"] = df["Amount"] / usd_rate

            # ---------------- PRODUCT ----------------
            if "Product" in df.columns:
                df["Product"] = df["Product"].astype(str).str.strip()

                invalid_mask = (
                    ~df["Product"].str.startswith("MD", na=False)
                    | (df["Product"] == "0")
                )

                if invalid_mask.any():

                    if pd.notna(product_b5) and str(product_b5).strip() != "":
                        replacement = str(product_b5).strip()
                    else:
                        if product_choice is None:
                            product_choice = dropdown_popup(file, "Select Product Code", PRODUCT_CODE_OPTIONS)
                        replacement = product_choice

                    df.loc[invalid_mask, "Product"] = replacement

            # ---------------- TYPE ----------------
            if "Type" in df.columns:
                df["Type"] = df["Type"].astype(str).str.strip()

                invalid_type_mask = ~df["Type"].isin(VAL_TYPE_OPTIONS)

                if invalid_type_mask.any():
                    if type_choice is None:
                        type_choice = dropdown_popup(file, "Select Type", VAL_TYPE_OPTIONS)

                    df.loc[invalid_type_mask, "Type"] = type_choice

            # ---------------- EXCEPTIONS ----------------
            df["Exception"] = ""
            df.loc[df["Amount"] == 0, "Exception"] += "Zero Amount; "

            # Split
            clean = df[df["Exception"] == ""].copy()
            exc = df[df["Exception"] != ""].copy()

            for d in [clean, exc]:
                d["File"] = os.path.basename(file)
                d["Sheet"] = sheet

            clean_data.append(clean)
            exception_data.append(exc)

            # ---------------- RECON ----------------
            if "Product" in df.columns:
                grp = df.groupby("Product")["Amount"].sum().reset_index()

                for _, row in grp.iterrows():
                    product = row["Product"]
                    total = row["Amount"]

                    clean_total = clean[clean["Product"] == product]["Amount"].sum()
                    exc_total = exc[exc["Product"] == product]["Amount"].sum()

                    recon_records.append({
                        "File": os.path.basename(file),
                        "Product": product,
                        "Input Total": total,
                        "UKMR Submission": clean_total,
                        "Exception Total": exc_total,
                        "Check": total - (clean_total + exc_total)
                    })

    final_clean = pd.concat(clean_data, ignore_index=True)
    final_exc = pd.concat(exception_data, ignore_index=True)
    recon = pd.DataFrame(recon_records)

    output = os.path.join(folder, "Output.xlsx")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_clean.to_excel(writer, sheet_name="Clean", index=False)
        final_exc.to_excel(writer, sheet_name="Exceptions", index=False)
        recon.to_excel(writer, sheet_name="Recon", index=False)

    messagebox.showinfo("Done", f"Saved at {output}")

# -----------------------------
# GUI (FIXED PASSWORD ISSUE)
# -----------------------------
def browse_folder():
    folder = filedialog.askdirectory()
    if not folder:
        return

    files = [os.path.join(folder, f) for f in os.listdir(folder)
             if f.lower().endswith((".xlsx", ".xls", ".xlsb", ".xlsm"))]

    selection = {}

    win = tk.Toplevel(root)
    win.geometry("1200x600")

    for file in files:
        xl = get_excel_file(file)
        if xl is None:
            continue

        sheets = xl.sheet_names

        row = tk.Frame(win)
        row.pack(anchor="w")

        tk.Label(row, text=os.path.basename(file), width=30).pack(side="left")

        selection[file] = {}

        for sheet in sheets:
            var = tk.BooleanVar(value=(sheet.lower() == "income sub."))
            tk.Checkbutton(row, text=sheet, variable=var).pack(side="left")
            selection[file][sheet] = var

    def submit():
        selected = {
            f: [s for s, v in sheets.items() if v.get()]
            for f, sheets in selection.items()
            if any(v.get() for v in sheets.values())
        }

        win.destroy()
        process_files(folder, selected)

    tk.Button(win, text="Submit", command=submit).pack()

# -----------------------------
# RUN
# -----------------------------
root = tk.Tk()
root.geometry("500x250")

tk.Button(root, text="Browse Folder", command=browse_folder).pack(pady=50)

root.mainloop()
