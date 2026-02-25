```
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pyxlsb import open_workbook

# ----------------------------
# CONFIG
# ----------------------------

list_type = ["", "NA", "N/A", None]
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

def read_c3(file, sheet):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None)
        return df.iloc[2, 2]
    except:
        return None

def ask_product_code(file, sheet):
    popup = tk.Toplevel()
    popup.title("Select Product Code")

    tk.Label(
        popup,
        text=f"C3 is empty in:\nFile: {os.path.basename(file)}\nSheet: {sheet}"
    ).pack(pady=10)

    selected = tk.StringVar()
    dropdown = ttk.Combobox(popup, textvariable=selected, values=list_prodcode)
    dropdown.pack(pady=5)

    def submit():
        if not selected.get():
            messagebox.showerror("Error", "Please select a product code.")
        else:
            popup.destroy()

    tk.Button(popup, text="Submit", command=submit).pack(pady=10)
    popup.grab_set()
    popup.wait_window()

    return selected.get()

# ----------------------------
# MAIN PROCESSING
# ----------------------------

def process_files(folder, selected_sheets):
    files = get_excel_files(folder)

    all_data = []
    header_reference = None

    for file in files:
        sheets = get_sheets(file)

        for sheet in selected_sheets:
            if sheet not in sheets:
                continue

            df = read_sheet(file, sheet)

            # Drop fully empty rows
            df = df.dropna(how="all")

            # Validate header
            if header_reference is None:
                header_reference = list(df.columns)
            else:
                if list(df.columns) != header_reference:
                    messagebox.showerror(
                        "Header Error",
                        f"Header mismatch in {os.path.basename(file)} - {sheet}"
                    )
                    return

            # Read C3
            c3 = read_c3(file, sheet)

            if pd.isna(c3) or c3 == "":
                c3 = ask_product_code(file, sheet)

            # Replace Product column
            if "Product" in df.columns:
                df["Product"] = df["Product"].replace(list_type, pd.NA)
                df["Product"] = df["Product"].fillna(c3)

            # Remove Amount = 0
            if "Amount" in df.columns:
                df = df[df["Amount"] != 0]

            # Remove alphanumeric Customer Number
            if "Customer Number" in df.columns:
                df = df[df["Customer Number"].astype(str).str.isnumeric()]

            df["Source File"] = os.path.basename(file)
            df["Source Sheet"] = sheet

            all_data.append(df)

    if not all_data:
        messagebox.showwarning("No Data", "No data found to process.")
        return

    final_df = pd.concat(all_data, ignore_index=True)

    output_path = os.path.join(folder, "cleaned_output.csv")
    final_df.to_csv(output_path, index=False)

    messagebox.showinfo("Success", f"File saved at:\n{output_path}")

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

    sheet_set = set()
    for file in files:
        sheet_set.update(get_sheets(file))

    sheet_list = list(sheet_set)

    sheet_window = tk.Toplevel(root)
    sheet_window.title("Select Sheets")

    tk.Label(sheet_window, text="Select Sheets to Process").pack(pady=10)

    listbox = tk.Listbox(sheet_window, selectmode=tk.MULTIPLE, width=50)
    listbox.pack(padx=20, pady=10)

    for idx, sheet in enumerate(sheet_list):
        listbox.insert(tk.END, sheet)
        if sheet == "IncomeSubtype":
            listbox.selection_set(idx)

    def submit():
        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "Please select at least one sheet.")
            return

        selected_sheets = [sheet_list[i] for i in selected_indices]
        sheet_window.destroy()
        process_files(folder, selected_sheets)

    tk.Button(sheet_window, text="Submit", command=submit).pack(pady=10)

# ----------------------------
# RUN APP
# ----------------------------

root = tk.Tk()
root.title("Excel Product Cleaner")

tk.Label(root, text="Select Folder Containing Excel Files").pack(pady=20)
tk.Button(root, text="Browse Folder", command=browse_folder).pack(pady=10)

root.mainloop()
