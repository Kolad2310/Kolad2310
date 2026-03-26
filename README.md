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
HEADER_ROW = 6  # your header row

# Colors
FILL_BLUE = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
FILL_GREEN = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
FILL_ORANGE = PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid")
FILL_GREY = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")


# -----------------------------
# HELPERS
# -----------------------------
def get_excel_files(folder):
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith((".xls", ".xlsx", ".xlsb"))
    ]


def ask_password(file):
    return simpledialog.askstring(
        "Password Required",
        f"Enter password for:\n{os.path.basename(file)}",
        show="*"
    )


def read_excel_safe(file, sheet, password_cache):
    try:
        return pd.read_excel(file, sheet_name=sheet, header=HEADER_ROW)
    except Exception:
        if file not in password_cache:
            pwd = ask_password(file)
            password_cache[file] = pwd

        try:
            return pd.read_excel(
                file,
                sheet_name=sheet,
                header=HEADER_ROW,
                engine="openpyxl",
                storage_options={"password": password_cache[file]}
            )
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open {file}\n{str(e)}")
            return None


def auto_adjust_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_len + 2


def color_recon_sheet(ws):
    header = [cell.value for cell in ws[1]]

    for col_idx, col_name in enumerate(header, start=1):
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
        clean_data = []
        exception_data = []
        recon_records = []

        password_cache = {}
        header_reference = None

        usd_rate = simpledialog.askfloat(
            "USD Rate",
            "Enter USD rate (1 GBP = ? USD):"
        )

        for file, sheets in selection_dict.items():
            for sheet in sheets:
                df = read_excel_safe(file, sheet, password_cache)
                if df is None:
                    continue

                df = df.dropna(how="all")

                # ---------------- HEADER CHECK ----------------
                cols = [c.strip().lower() for c in df.columns]

                if header_reference is None:
                    header_reference = cols
                else:
                    if cols != header_reference:
                        messagebox.showerror(
                            "Header Error",
                            f"Header mismatch:\n{os.path.basename(file)} - {sheet}"
                        )
                        return

                # ---------------- USD CONVERSION ----------------
                try:
                    currency = df.iloc[3, 4]
                except:
                    currency = None

                if str(currency).strip().upper() == "USD" and usd_rate:
                    if "Amount" in df.columns:
                        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
                        df["Amount"] *= usd_rate

                # ---------------- PRODUCT CODE ----------------
                if "Product code" in df.columns:
                    mask = df["Product code"].isna() | (df["Product code"].astype(str).str.strip() == "")

                    if mask.any():
                        try:
                            c4 = df.iloc[3, 2]
                        except:
                            c4 = None

                        if pd.isna(c4) or str(c4).strip() == "":
                            c4 = simpledialog.askstring(
                                "Product Code",
                                f"Enter product code for:\n{os.path.basename(file)} - {sheet}"
                            )

                        df.loc[mask, "Product code"] = c4

                # ---------------- EXCEPTION LOGIC ----------------
                df["Exception_Reason"] = ""

                if "Amount" in df.columns:
                    df.loc[df["Amount"] == 0, "Exception_Reason"] += "Zero Amount; "

                if "Customer No." in df.columns:
                    cust = df["Customer No."].astype(str).str.strip()

                    mask = (
                        cust.eq("") |
                        cust.str.lower().eq("none") |
                        cust.str.match(r"^[A-Za-z]+$")
                    )

                    df.loc[mask, "Exception_Reason"] += "Invalid Customer; "

                exceptions = df[df["Exception_Reason"] != ""].copy()
                clean = df[df["Exception_Reason"] == ""].copy()

                for d in [exceptions, clean]:
                    d["Source File"] = os.path.basename(file)
                    d["Source Sheet"] = sheet

                clean_data.append(clean)
                exception_data.append(exceptions)

                # ---------------- RECON ----------------
                if "Product code" in df.columns and "Amount" in df.columns:
                    grp = df.groupby("Product code")["Amount"].sum().reset_index()

                    for _, row in grp.iterrows():
                        product = row["Product code"]
                        input_total = row["Amount"]

                        clean_total = clean[clean["Product code"] == product]["Amount"].sum()
                        exc_total = exceptions[exceptions["Product code"] == product]["Amount"].sum()

                        check = input_total - (clean_total + exc_total)

                        recon_records.append({
                            "File Name": os.path.basename(file),
                            "Product code": product,
                            "Input Total": input_total,
                            "UKMR Submission": clean_total,
                            "Exception Total": exc_total,
                            "Check (Should be 0)": check
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

        # ---------------- POST FORMATTING ----------------
        wb = load_workbook(output_path)

        # Auto width
        for ws in wb.worksheets:
            auto_adjust_width(ws)

        # Color recon
        color_recon_sheet(wb["Reconciliation"])

        wb.save(output_path)

        messagebox.showinfo("Success", f"Saved at:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# -----------------------------
# GUI
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

    selection_vars = {}

    for file in files:
        sheets = pd.ExcelFile(file).sheet_names
        selection_vars[file] = {}

        tk.Label(sheet_window, text=os.path.basename(file)).pack(anchor="w")

        for sheet in sheets:
            var = tk.BooleanVar(value=True)
            chk = tk.Checkbutton(sheet_window, text=sheet, variable=var)
            chk.pack(anchor="w")
            selection_vars[file][sheet] = var

    def submit():
        selection_dict = {}

        for file, sheets in selection_vars.items():
            selected = [s for s, v in sheets.items() if v.get()]
            if selected:
                selection_dict[file] = selected

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
root.title("Excel Product Code Cleaner")
root.geometry("400x200")

tk.Label(root, text="Select Folder Containing Excel Files").pack(pady=20)
tk.Button(root, text="Browse Folder", command=browse_folder).pack()

root.mainloop()
