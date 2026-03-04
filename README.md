```
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import os
from datetime import datetime
import traceback
from openpyxl import Workbook, load_workbook

LOG_FILE = "Processing_Log.txt"
HEADER_FILE = "Header_Diagnostics.xlsx"

x_mask_value = 1000

entity_list = ["Entity1","Entity2"]
globalbusiness_list = ["Business1","Business2"]

file_store = {
    "RWA_Actuals": [],
    "RWA_Plan": [],
    "SD_Actuals": [],
    "SD_Plan": [],
    "AVBS_Actuals": [],
    "AVBS_Plan": [],
    "PBT_Actuals": [],
    "PBT_Plan": [],
    "BS_Actuals": [],
    "BS_Plan": []
}

mapper_file = None


# ------------------------------------------------
# LOGGING
# ------------------------------------------------

def log(msg):

    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"

    print(line, flush=True)

    with open(LOG_FILE,"a",encoding="utf-8") as f:
        f.write(line+"\n")

    try:
        log_box.insert(tk.END,line+"\n")
        log_box.see(tk.END)
        root.update()
    except:
        pass


# ------------------------------------------------
# HEADER DETECTION
# ------------------------------------------------

def detect_header(df):

    for i in range(min(60,len(df))):

        row = [str(v).strip().replace("\xa0","").lower() for v in df.iloc[i]]

        if (
            any("year" in r for r in row) and
            any("currency" in r for r in row) and
            any("entity" in r for r in row)
        ):
            log(f"Header detected at row {i+1}")
            return i

    log("Header not found")
    return None


# ------------------------------------------------
# NORMALIZE TEXT
# ------------------------------------------------

def normalize_series(s):

    return (
        s.astype(str)
        .str.replace("\xa0","",regex=False)
        .str.strip()
        .str.upper()
    )


# ------------------------------------------------
# CONSOLIDATE CATEGORY
# ------------------------------------------------

def consolidate_category(category):

    log(f"--- Processing {category} ---")

    collected = []
    header_info = []

    for file in file_store[category]:

        log(f"Reading file {file}")

        xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:

            log(f"Reading sheet {sheet}")

            preview = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None,
                nrows=60,
                dtype=object
            )

            header_row = detect_header(preview)

            if header_row is None:
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row,
                dtype=object
            )

            df.columns = df.columns.str.strip().str.lower()

            # ADD FILE NAME COLUMN
            df.insert(0,"source_file",os.path.basename(file))

            df = df.dropna(how="all")

            for col in df.columns:
                df[col] = pd.to_numeric(df[col],errors="ignore")

            collected.append(df)

            header_info.append({
                "Category":category,
                "File":os.path.basename(file),
                "Sheet":sheet,
                "Header_Row":header_row+1,
                "Columns":",".join(df.columns)
            })

            log(f"{len(df)} rows loaded")

    if not collected:
        return pd.DataFrame(), header_info

    df = pd.concat(collected,ignore_index=True)

    log(f"{category} rows before filter {len(df)}")

    if "currency" in df.columns:
        df["currency"] = normalize_series(df["currency"])

    if "entity" in df.columns:
        df["entity"] = normalize_series(df["entity"])

    if "global business" in df.columns:
        df["global business"] = normalize_series(df["global business"])

    entity_clean = [e.upper() for e in entity_list]
    gb_clean = [g.upper() for g in globalbusiness_list]

    if "currency" in df.columns:
        df = df[df["currency"]=="USD"]

    if "entity" in df.columns:
        df = df[df["entity"].isin(entity_clean)]

    if "global business" in df.columns:
        df = df[df["global business"].isin(gb_clean)]

    log(f"{category} rows after filter {len(df)}")

    # PBT adjustment
    if category=="PBT_Actuals" and "year" in df.columns:

        year_index = df.columns.get_loc("year")

        numeric_cols = df.columns[year_index+1:]

        for col in numeric_cols:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col]/1000

    return df, header_info


# ------------------------------------------------
# WRITE OUTPUT FILE
# ------------------------------------------------

def write_output_file(file_name, sheet_dict):

    log(f"Writing file {file_name}")

    wb = Workbook(write_only=True)

    for sheet, df in sheet_dict.items():

        ws = wb.create_sheet(sheet)

        if df.empty:
            log(f"{sheet} empty")
            continue

        ws.append(df.columns.str.title().tolist())

        for row in df.itertuples(index=False):

            clean_row = [
                None if pd.isna(v) else v
                for v in row
            ]

            ws.append(clean_row)

        log(f"{sheet} written {len(df)} rows")

    wb.save(file_name)


# ------------------------------------------------
# CREATE MAPPED FILE
# ------------------------------------------------

def create_mapped_copy(original_file):

    log("Creating mapped file")

    entity_map = pd.read_excel(mapper_file,sheet_name="Entity")
    product_map = pd.read_excel(mapper_file,sheet_name="Product")

    entity_map["Entity"] = normalize_series(entity_map["Entity"])
    product_map["Product"] = normalize_series(product_map["Product"])

    entity_dict = dict(zip(
        entity_map["Entity"],
        entity_map["Proxy Entity Code"]
    ))

    product_dict = dict(zip(
        product_map["Product"],
        product_map["Proxy Product Code"]
    ))

    wb = load_workbook(original_file)

    for sheet in wb.sheetnames:

        ws = wb[sheet]

        headers = [cell.value for cell in ws[1]]

        if "Function" in headers:

            col_index = headers.index("Function")+1
            ws.delete_cols(col_index)
            headers.pop(col_index-1)

            log(f"{sheet} function column removed")

        if "Entity" in headers:

            e_idx = headers.index("Entity")+1
            count=0

            for r in range(2,ws.max_row+1):

                val = str(ws.cell(r,e_idx).value).strip().upper()

                if val in entity_dict:

                    ws.cell(r,e_idx).value = entity_dict[val]
                    count+=1

            log(f"{sheet} entity mapped {count} rows")

        if "Product" in headers:

            p_idx = headers.index("Product")+1
            count=0

            for r in range(2,ws.max_row+1):

                val = str(ws.cell(r,p_idx).value).strip().upper()

                if val in product_dict:

                    ws.cell(r,p_idx).value = product_dict[val]
                    count+=1

            log(f"{sheet} product mapped {count} rows")

    new_file = original_file.replace(".xlsx","_Mapped.xlsx")

    wb.save(new_file)

    log(f"{new_file} created")

    return new_file


# ------------------------------------------------
# CREATE MASKED FILE
# ------------------------------------------------

def create_masked_file(mapped_file):

    log("Creating masked file")

    wb = load_workbook(mapped_file)

    for sheet in wb.sheetnames:

        ws = wb[sheet]

        headers = [c.value for c in ws[1]]

        for col in range(1,len(headers)+1):

            name = headers[col-1].lower()

            if name in ["year","entity","currency","product","source_file"]:
                continue

            for r in range(2,ws.max_row+1):

                val = ws.cell(r,col).value

                if isinstance(val,(int,float)):

                    if "pbt" in sheet.lower():
                        ws.cell(r,col).value = val/x_mask_value
                    else:
                        ws.cell(r,col).value = val/(x_mask_value/2)

    new_file = mapped_file.replace("_Mapped.xlsx","_Masked.xlsx")

    wb.save(new_file)

    log(f"{new_file} created")


# ------------------------------------------------
# MAIN PROCESS
# ------------------------------------------------

def start_processing():

    try:

        open(LOG_FILE,"w").close()

        tables={}
        headers=[]

        for key in file_store:

            df,h = consolidate_category(key)

            tables[key]=df

            headers.extend(h)

        pd.DataFrame(headers).to_excel(HEADER_FILE,index=False)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        output_file = f"Consolidated_Output_{ts}.xlsx"

        write_output_file(
            output_file,
            {
                "RWA Actuals":tables["RWA_Actuals"],
                "RWA Plan":tables["RWA_Plan"],
                "AVBS_SD Actuals":pd.concat([tables["AVBS_Actuals"],tables["SD_Actuals"]]),
                "AVBS_SD Plan":pd.concat([tables["AVBS_Plan"],tables["SD_Plan"]]),
                "PBT_BS Actuals":pd.concat([tables["PBT_Actuals"],tables["BS_Actuals"]]),
                "PBT_BS Plan":pd.concat([tables["PBT_Plan"],tables["BS_Plan"]])
            }
        )

        mapped_file = create_mapped_copy(output_file)

        create_masked_file(mapped_file)

        messagebox.showinfo("Success","Processing complete")

    except Exception:

        log(traceback.format_exc())

        messagebox.showerror("Error","Processing failed")


# ------------------------------------------------
# GUI
# ------------------------------------------------

root = tk.Tk()
root.title("Financial Consolidation Tool")
root.geometry("900x650")

frame=tk.Frame(root)
frame.pack(pady=10)

labels={}

def select_files(key):

    files=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xls")])

    if files:
        file_store[key]=list(files)
        labels[key].config(text=f"{len(files)} selected")


def select_mapper():

    global mapper_file

    mapper_file=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])

for i,key in enumerate(file_store):

    tk.Label(frame,text=key,width=20).grid(row=i,column=0)

    tk.Button(frame,text="Select Files",
              command=lambda k=key:select_files(k)).grid(row=i,column=1)

    lbl=tk.Label(frame,text="None")
    lbl.grid(row=i,column=2)

    labels[key]=lbl

tk.Button(root,text="Select Mapper File",
          command=select_mapper).pack(pady=10)

tk.Button(root,text="Run Process",
          command=start_processing,
          bg="green",fg="white").pack(pady=10)

# LOG WINDOW
log_box = ScrolledText(root,height=15)
log_box.pack(fill="both",expand=True,padx=10,pady=10)

root.mainloop()
