```
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pyxlsb import open_workbook

# ----------------------------
# CONFIG
# ----------------------------

list_type = ["", "NA", "N/A", "None"]
list_prodcode = ["P100", "P200", "P300", "P400"]

# ----------------------------
# HELPERS
# ----------------------------

def get_excel_files(folder):
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith((".xls", ".xlsx", ".xlsb"))
    ]

def get_sheets(file):
    try:
        if file.endswith(".xlsb"):
            with open_workbook(file) as wb:
                return wb.sheets
        else:
            return pd.ExcelFile(file).sheet_names
    except:
        return []

def read_sheet(file, sheet):
    if file.endswith(".xlsb"):
        return pd.read_excel(file, sheet_name=sheet, header=6, engine="pyxlsb")
    elif file.endswith(".xls"):
        return pd.read_excel(file, sheet_name=sheet, header=6, engine="xlrd")
    else:
        return pd.read_excel(file, sheet_name=sheet, header=6, engine="openpyxl")

def read_cell(file, sheet, row, col):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None)
        return df.iloc[row, col]
    except:
        return None

# ----------------------------
# POPUPS
# ----------------------------

def ask_product_code(file, sheet):
    popup = tk.Toplevel(root)
    popup.title("Select Product Code")
    popup.grab_set()

    tk.Label(
        popup,
        text=f"Product required\n\nFile: {os.path.basename(file)}\nSheet: {sheet}"
    ).pack(padx=20, pady=10)

    selected = tk.StringVar()

    dropdown = ttk.Combobox(
        popup,
        textvariable=selected,
        values=list_prodcode,
        state="readonly"
    )
    dropdown.pack(pady=5)

    def submit():
        if not selected.get():
            messagebox.showerror("Error", "Select a product code.")
        else:
            popup.destroy()

    tk.Button(popup, text="Submit", command=submit).pack(pady=10)
    popup.wait_window()

    return selected.get()

def ask_usd_rate():
    popup = tk.Toplevel(root)
    popup.title("USD Conversion Required")
    popup.grab_set()

    tk.Label(
        popup,
        text="USD detected.\n\nEnter rate:\n1 GBP = ___ USD"
    ).pack(padx=20, pady=10)

    rate_var = tk.StringVar()
    entry = tk.Entry(popup, textvariable=rate_var)
    entry.pack(pady=5)

    def submit():
        try:
            rate = float(rate_var.get())
            if rate <= 0:
                raise ValueError
            popup.destroy()
        except:
            messagebox.showerror("Error", "Enter valid numeric rate.")

    tk.Button(popup, text="Submit", command=submit).pack(pady=10)
    popup.wait_window()

    return float(rate_var.get())

# ----------------------------
# MAIN PROCESSING
# ----------------------------

def process_files(folder, selection_dict):

    try:
        clean_data = []
        exception_data = []
        recon_records = []
        header_reference = None
        usd_rate = None

        for file, sheets in selection_dict.items():

            for sheet in sheets:

                df = read_sheet(file, sheet)
                df = df.dropna(how="all")

                # ----------------------------
                # Header validation
                # ----------------------------
                if header_reference is None:
                    header_reference = list(df.columns)
                else:
                    if list(df.columns) != header_reference:
                        messagebox.showerror(
                            "Header Error",
                            f"Header mismatch in {os.path.basename(file)} - {sheet}"
                        )
                        return

                # ----------------------------
                # USD Conversion (Correct Cell)
                # Excel row 4 col 5 -> index (3,4)
                # ----------------------------
                currency = read_cell(file, sheet, 3, 4)

                if str(currency).strip().upper() == "USD":

                    if usd_rate is None:
                        usd_rate = ask_usd_rate()

                    if "Amount" in df.columns:
                        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
                        df["Amount"] = df["Amount"] / usd_rate

                # ----------------------------
                # PRODUCT LOGIC (C4 → Dropdown)
                # Excel row 4 col 3 -> index (3,2)
                # ----------------------------
                if "Product" in df.columns:

                    mask = (
                        df["Product"].isna() |
                        df["Product"].astype(str).str.strip().isin(list_type)
                    )

                    if mask.any():

                        c4 = read_cell(file, sheet, 3, 2)

                        if pd.isna(c4) or str(c4).strip() == "":
                            c4 = ask_product_code(file, sheet)

                        df.loc[mask, "Product"] = c4

                # ----------------------------
                # EXCEPTION LOGIC
                # ----------------------------

                df["Exception_Reason"] = ""

                # Zero Amount
                if "Amount" in df.columns:
                    zero_mask = df["Amount"] == 0
                    df.loc[zero_mask, "Exception_Reason"] += "Zero Amount; "

                # Customer validation
                if "Customer Number" in df.columns:

                    cust_series = df["Customer Number"].astype(str).str.strip()

                    cust_mask = (
                        cust_series.eq("") |
                        cust_series.str.lower().eq("none") |
                        cust_series.str.contains(r"[a-zA-Z]", na=False)
                    )

                    df.loc[cust_mask, "Exception_Reason"] += "Invalid Customer; "

                exceptions = df[df["Exception_Reason"] != ""].copy()
                clean = df[df["Exception_Reason"] == ""].copy()

                for d in [exceptions, clean]:
                    d["Source File"] = os.path.basename(file)
                    d["Source Sheet"] = sheet

                clean_data.append(clean)
                exception_data.append(exceptions)

                # ----------------------------
                # RECONCILIATION (Use FINAL df)
                # ----------------------------

                if "Product" in df.columns and "Amount" in df.columns:

                    product_groups = df.groupby("Product")["Amount"].sum().reset_index()

                    for _, row in product_groups.iterrows():

                        product = row["Product"]
                        input_total = row["Amount"]

                        clean_total = clean[clean["Product"] == product]["Amount"].sum()
                        exc_total = exceptions[exceptions["Product"] == product]["Amount"].sum()

                        check_value = input_total - (clean_total + exc_total)

                        recon_records.append({
                            "File Name": os.path.basename(file),
                            "Product": product,
                            "Input Total": input_total,
                            "Exception Total": exc_total,
                            "Check (Should be 0)": check_value
                        })

        if not clean_data:
            messagebox.showwarning("No Data", "No valid data found.")
            return

        final_clean = pd.concat(clean_data, ignore_index=True)
        final_exceptions = pd.concat(exception_data, ignore_index=True)
        recon_df = pd.DataFrame(recon_records)

        output_path = os.path.join(folder, "processed_output.xlsx")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            final_clean.to_excel(writer, sheet_name="Clean_Data", index=False)
            final_exceptions.to_excel(writer, sheet_name="Exceptions", index=False)
            recon_df.to_excel(writer, sheet_name="Reconciliation", index=False)

        messagebox.showinfo("Success", f"Saved at:\n{output_path}")

        root.quit()
        root.destroy()
        sys.exit()

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ----------------------------
# GUI
# ----------------------------

def browse_folder():

    folder = filedialog.askdirectory()
    if not folder:
        return

    files = get_excel_files(folder)

    if not files:
        messagebox.showerror("Error", "No Excel files found.")
        return

    sheet_window = tk.Toplevel(root)
    sheet_window.title("Select Sheets Per File")
    sheet_window.grab_set()

    canvas = tk.Canvas(sheet_window)
    scrollbar = tk.Scrollbar(sheet_window, orient="vertical", command=canvas.yview)
    frame = tk.Frame(canvas)

    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    selection_vars = {}

    for row, file in enumerate(files):

        tk.Label(frame, text=os.path.basename(file), width=30, anchor="w").grid(row=row, column=0, sticky="w")

        sheets = get_sheets(file)
        selection_vars[file] = {}

        for col, sheet in enumerate(sheets):

            var = tk.BooleanVar()
            if sheet == "IncomeSubtype":
                var.set(True)

            chk = tk.Checkbutton(frame, text=sheet, variable=var)
            chk.grid(row=row, column=col + 1, sticky="w")

            selection_vars[file][sheet] = var

    def submit():

        selection_dict = {}

        for file, sheets in selection_vars.items():
            selected = [s for s, v in sheets.items() if v.get()]
            if selected:
                selection_dict[file] = selected

        if not selection_dict:
            messagebox.showerror("Error", "Select at least one sheet.")
            return

        sheet_window.destroy()
        process_files(folder, selection_dict)

    tk.Button(sheet_window, text="Submit", command=submit).pack(pady=10)

# ----------------------------
# RUN
# ----------------------------

root = tk.Tk()
root.title("Excel Product Cleaner")
root.geometry("500x200")

tk.Label(root, text="Select Folder Containing Excel Files").pack(pady=20)
tk.Button(root, text="Browse Folder", command=browse_folder).pack(pady=10)

root.mainloop()
