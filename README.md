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


def ask_product_code(file, sheet):
    popup = tk.Toplevel(root)
    popup.title("Select Product Code")
    popup.grab_set()

    tk.Label(
        popup,
        text=f"Product required\n\nFile: {os.path.basename(file)}\nSheet: {sheet}",
        justify="left"
    ).pack(padx=20, pady=10)

    selected = tk.StringVar()

    dropdown = ttk.Combobox(
        popup,
        textvariable=selected,
        values=list_prodcode,
        state="readonly",
        width=30
    )
    dropdown.pack(pady=5)

    def submit():
        if not selected.get():
            messagebox.showerror("Error", "Please select a product code.")
        else:
            popup.destroy()

    tk.Button(popup, text="Submit", command=submit).pack(pady=10)

    popup.wait_window()
    return selected.get()

# ----------------------------
# PROCESSING
# ----------------------------

def process_files(folder, selection_dict):

    try:
        all_data = []
        header_reference = None

        for file, sheets in selection_dict.items():

            for sheet in sheets:

                df = read_sheet(file, sheet)
                df = df.dropna(how="all")

                # Header validation
                if header_reference is None:
                    header_reference = list(df.columns)
                else:
                    if list(df.columns) != header_reference:
                        messagebox.showerror(
                            "Header Error",
                            f"Header mismatch in\n{os.path.basename(file)} - {sheet}"
                        )
                        return

                # ----------------------------
                # PRODUCT LOGIC (FINAL CORRECT)
                # ----------------------------
                if "Product" in df.columns:

                    mask = (
                        df["Product"].isna() |
                        df["Product"].astype(str).str.strip().isin(list_type)
                    )

                    if mask.any():

                        # Read C4 (row index 3, col index 2)
                        c4 = read_cell(file, sheet, 3, 2)

                        if pd.isna(c4) or str(c4).strip() == "":
                            c4 = ask_product_code(file, sheet)

                        df.loc[mask, "Product"] = c4

                # Remove Amount = 0
                if "Amount" in df.columns:
                    df = df[df["Amount"] != 0]

                # Remove non-numeric Customer Number
                if "Customer Number" in df.columns:
                    df = df[
                        df["Customer Number"]
                        .astype(str)
                        .str.strip()
                        .str.isnumeric()
                    ]

                df["Source File"] = os.path.basename(file)
                df["Source Sheet"] = sheet

                all_data.append(df)

        if not all_data:
            messagebox.showwarning("No Data", "No valid data found.")
            return

        final_df = pd.concat(all_data, ignore_index=True)

        output_path = os.path.join(folder, "cleaned_output.csv")
        final_df.to_csv(output_path, index=False)

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

    frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

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
            selected = [
                sheet for sheet, var in sheets.items() if var.get()
            ]
            if selected:
                selection_dict[file] = selected

        if not selection_dict:
            messagebox.showerror("Error", "Please select at least one sheet.")
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
