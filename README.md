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

FILL_BLUE = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
FILL_GREEN = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
FILL_ORANGE = PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid")
FILL_GREY = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

password_cache = {}
product_code_cache = {}

# -----------------------------
# FILE FETCH
# -----------------------------
def get_excel_files(folder):
    files = []
    for root_dir, _, f_list in os.walk(folder):
        for f in f_list:
            if f.lower().endswith((".xls", ".xlsx", ".xlsb", ".xlsm")) and not f.startswith("~$"):
                files.append(os.path.join(root_dir, f))
    return files

# -----------------------------
# PASSWORD
# -----------------------------
def ask_password(file):
    return simpledialog.askstring("Password", f"Enter password for:\n{os.path.basename(file)}", show="*")

# -----------------------------
# DROPDOWN
# -----------------------------
def ask_product_code_dropdown(file):
    popup = tk.Toplevel(root)
    popup.title("Select Product Code")
    popup.geometry("350x150")
    popup.grab_set()

    tk.Label(popup, text=f"Select Product Code for:\n{os.path.basename(file)}").pack(pady=10)

    selected = tk.StringVar()
    combo = ttk.Combobox(popup, values=PRODUCT_CODE_OPTIONS, textvariable=selected, state="readonly")
    combo.pack(pady=5)
    combo.current(0)

    def submit():
        popup.destroy()

    tk.Button(popup, text="Submit", command=submit).pack(pady=10)
    popup.wait_window()

    return selected.get()

# -----------------------------
# READ B5 & E5
# -----------------------------
def read_metadata_cells(file, sheet):
    try:
        temp = pd.read_excel(file, sheet_name=sheet, header=None, nrows=6)
        return temp.iloc[4, 1], temp.iloc[4, 4]  # B5, E5
    except:
        return None, None

# -----------------------------
# SAFE READ
# -----------------------------
def read_excel_safe(file, sheet):
    try:
        if file.lower().endswith(".xlsb"):
            return pd.read_excel(file, sheet_name=sheet, header=HEADER_ROW, engine="pyxlsb")
        return pd.read_excel(file, sheet_name=sheet, header=HEADER_ROW)
    except:
        if file not in password_cache:
            pwd = ask_password(file)
            if not pwd:
                return None
            password_cache[file] = pwd
        try:
            return pd.read_excel(file, sheet_name=sheet, header=HEADER_ROW, engine="openpyxl")
        except:
            return None

# -----------------------------
# FORMATTING
# -----------------------------
def auto_adjust_width(ws):
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

def color_recon_sheet(ws):
    header = [c.value for c in ws[1]]
    for col_idx, col_name in enumerate(header, 1):
        fill = None
        if col_name == "Input Total":
            fill = FILL_BLUE
        elif col_name == "UKMR Submission":
            fill = FILL_GREEN
        elif col_name == "Exception Total":
            fill = FILL_ORANGE
        elif "Check" in str(col_name):
            fill = FILL_GREY

        if fill:
            for row in ws.iter_rows(min_row=1, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    cell.fill = fill

# -----------------------------
# MAIN PROCESS
# -----------------------------
def process_files(folder, selection_dict):
    try:
        clean_data, exception_data, recon_records = [], [], []
        header_reference = None

        usd_rate = simpledialog.askfloat("USD Rate", "Enter USD rate:")

        for file, sheets in selection_dict.items():
            for sheet in sheets:

                # 🔥 Read B5 & E5
                product_code_b5, currency_e5 = read_metadata_cells(file, sheet)

                df = read_excel_safe(file, sheet)
                if df is None or df.empty:
                    continue

                df = df.dropna(how="all")

                cols = [c.strip().lower() for c in df.columns]
                if header_reference is None:
                    header_reference = cols
                elif cols != header_reference:
                    messagebox.showerror("Header Error", f"{os.path.basename(file)} - {sheet}")
                    return

                # USD conversion
                if str(currency_e5).strip().upper() == "USD" and usd_rate:
                    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
                    df["Amount"] *= usd_rate

                # ---------------- PRODUCT CODE FINAL LOGIC ----------------
                if "Product code" in df.columns:
                    df["Product code"] = df["Product code"].astype(str).str.strip()

                    invalid_mask = ~df["Product code"].str.startswith("MD", na=False)

                    if invalid_mask.any():

                        # CASE 1: B5 has value → use it
                        if pd.notna(product_code_b5) and str(product_code_b5).strip() != "":
                            replacement_code = str(product_code_b5).strip()

                        # CASE 2: B5 empty → dropdown
                        else:
                            if file not in product_code_cache:
                                product_code_cache[file] = ask_product_code_dropdown(file)
                            replacement_code = product_code_cache[file]

                        df.loc[invalid_mask, "Product code"] = replacement_code

                # ---------------- EXCEPTIONS ----------------
                df["Exception_Reason"] = ""

                df.loc[df["Amount"] == 0, "Exception_Reason"] += "Zero Amount; "

                if "Customer No." in df.columns:
                    cust = df["Customer No."].astype(str).str.strip()
                    mask = cust.eq("") | cust.str.lower().eq("none") | cust.str.match(r"^[A-Za-z]+$")
                    df.loc[mask, "Exception_Reason"] += "Invalid Customer; "

                exceptions = df[df["Exception_Reason"] != ""].copy()
                clean = df[df["Exception_Reason"] == ""].copy()

                for d in [clean, exceptions]:
                    d["Source File"] = os.path.basename(file)
                    d["Source Sheet"] = sheet

                clean_data.append(clean)
                exception_data.append(exceptions)

                # ---------------- RECON ----------------
                grp = df.groupby("Product code")["Amount"].sum().reset_index()

                for _, row in grp.iterrows():
                    product = row["Product code"]
                    input_total = row["Amount"]

                    clean_total = clean[clean["Product code"] == product]["Amount"].sum()
                    exc_total = exceptions[exceptions["Product code"] == product]["Amount"].sum()

                    recon_records.append({
                        "File Name": os.path.basename(file),
                        "Product code": product,
                        "Input Total": input_total,
                        "UKMR Submission": clean_total,
                        "Exception Total": exc_total,
                        "Check (Should be 0)": input_total - (clean_total + exc_total)
                    })

        if not clean_data:
            messagebox.showwarning("No Data", "No valid data found")
            return

        final_clean = pd.concat(clean_data, ignore_index=True)
        final_exc = pd.concat(exception_data, ignore_index=True)
        recon_df = pd.DataFrame(recon_records)

        output_path = os.path.join(folder, "Output_for_SME.xlsx")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            final_clean.to_excel(writer, sheet_name="Clean_Data", index=False)
            final_exc.to_excel(writer, sheet_name="Exceptions", index=False)
            recon_df.to_excel(writer, sheet_name="Reconciliation", index=False)

        wb = load_workbook(output_path)

        for ws in wb.worksheets:
            auto_adjust_width(ws)

        color_recon_sheet(wb["Reconciliation"])

        wb.save(output_path)

        messagebox.showinfo("Success", f"Saved at:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# -----------------------------
# GUI (same as before)
# -----------------------------
def browse_folder():
    folder = filedialog.askdirectory()
    if not folder:
        return

    files = get_excel_files(folder)

    if not files:
        messagebox.showerror("Error", "No Excel files found")
        return

    sheet_window = tk.Toplevel(root)
    sheet_window.title("Select Sheets")
    sheet_window.geometry("1200x600")

    canvas = tk.Canvas(sheet_window)
    scrollbar = tk.Scrollbar(sheet_window, command=canvas.yview)
    frame = tk.Frame(canvas)

    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    selection_vars = {}

    for file in files:
        try:
            sheets = pd.ExcelFile(file).sheet_names
        except:
            sheets = []

        selection_vars[file] = {}

        row = tk.Frame(frame)
        row.pack(fill="x", pady=5)

        tk.Label(row, text=os.path.basename(file), width=30, anchor="w").pack(side="left")

        for sheet in sheets:
            var = tk.BooleanVar(value=(sheet.lower() == "income sub."))
            tk.Checkbutton(row, text=sheet, variable=var).pack(side="left")
            selection_vars[file][sheet] = var

    def submit():
        selection_dict = {
            f: [s for s, v in sheets.items() if v.get()]
            for f, sheets in selection_vars.items()
            if any(v.get() for v in sheets.values())
        }

        if not selection_dict:
            messagebox.showerror("Error", "Select at least one sheet")
            return

        sheet_window.destroy()
        process_files(folder, selection_dict)

    tk.Button(sheet_window, text="Submit", command=submit).pack(pady=10)

# -----------------------------
# RUN
# -----------------------------
root = tk.Tk()
root.title("Excel Processor")
root.geometry("500x250")

tk.Label(root, text="Select Folder").pack(pady=20)
tk.Button(root, text="Browse Folder", command=browse_folder).pack()

root.mainloop()
