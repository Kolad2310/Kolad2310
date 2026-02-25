```
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pyxlsb import open_workbook

# ----------------------------
# CONFIGURATION
# ----------------------------

list_type = ["", "NA", "N/A", "None"]
list_prodcode = ["P100", "P200", "P300", "P400"]

# ----------------------------
# HELPER FUNCTIONS
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
    popup = tk.Toplevel(root)
    popup.title("Select Product Code")
    popup.grab_set()

    tk.Label(
        popup,
        text=f"C3 is empty\n\nFile: {os.path.basename(file)}\nSheet: {sheet}",
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
# MAIN PROCESSING
# ----------------------------

def process_files(folder, selected_sheets):

    try:
        files = get_excel_files(folder)
        all_data = []
        header_reference = None

        for file in files:
            sheets = get_sheets(file)

            for sheet in selected_sheets:
                if sheet not in sheets:
                    continue

                df = read_sheet(file, sheet)

                # Remove fully empty rows
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

                # Read C3
                c3 = read_c3(file, sheet)

                if pd.isna(c3) or str(c3).strip() == "":
                    c3 = ask_product_code(file, sheet)

                # ----------------------------
                # CORRECT PRODUCT COLUMN LOGIC
                # ----------------------------
                if "Product" in df.columns:
                    mask = (
                        df["Product"].isna() |
                        df["Product"].astype(str).str.strip().isin(list_type)
                    )
                    df.loc[mask, "Product"] = c3

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

        messagebox.showinfo("Success", f"File saved at:\n{output_path}")

        # Proper clean shutdown
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
        messagebox.showerror("Error", "No Excel files found in folder.")
        return

    # Collect all unique sheets across all files
    sheet_set = set()
    for file in files:
        sheet_set.update(get_sheets(file))

    sheet_list = sorted(list(sheet_set))

    # Single multi-select window
    sheet_window = tk.Toplevel(root)
    sheet_window.title("Select Sheets")
    sheet_window.grab_set()

    tk.Label(sheet_window, text="Select Sheets to Process").pack(pady=10)

    listbox = tk.Listbox(
        sheet_window,
        selectmode=tk.MULTIPLE,
        width=50,
        height=15
    )
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
# RUN APPLICATION
# ----------------------------

root = tk.Tk()
root.title("Excel Product Cleaner")
root.geometry("400x200")

root.protocol("WM_DELETE_WINDOW", root.destroy)

tk.Label(root, text="Select Folder Containing Excel Files").pack(pady=20)
tk.Button(root, text="Browse Folder", command=browse_folder).pack(pady=10)

root.mainloop()
