```
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

LOG_FILE = "Processing_Log.txt"

x_mask_value = 1000

entity_list = ["ENTITY1","ENTITY2"]
globalbusiness_list = ["BUSINESS1","BUSINESS2"]

mapper_file = None

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

# ------------------------------------------------
# LOGGING
# ------------------------------------------------

def log(msg):

    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"

    print(line)

    with open(LOG_FILE,"a",encoding="utf-8") as f:
        f.write(line+"\n")

# ------------------------------------------------
# HEADER DETECTION
# ------------------------------------------------

def detect_header(df):

    for i in range(min(60,len(df))):

        row=[str(v).strip().lower() for v in df.iloc[i]]

        if ("year" in row and "currency" in row and "entity" in row):

            log(f"Header detected at row {i+1}")
            return i

    log("Header not detected")
    return None

# ------------------------------------------------
# NORMALIZE TEXT
# ------------------------------------------------

def norm(series):

    return (
        series.astype(str)
        .str.replace("\xa0","",regex=False)
        .str.strip()
        .str.upper()
    )

# ------------------------------------------------
# CONSOLIDATE CATEGORY
# ------------------------------------------------

def consolidate_category(category):

    log(f"\n--- Consolidating {category} ---")

    dfs=[]

    for file in file_store[category]:

        log(f"Reading file: {file}")

        xls=pd.ExcelFile(file)

        for sheet in xls.sheet_names:

            log(f"Processing sheet: {sheet}")

            preview=pd.read_excel(file,sheet_name=sheet,header=None,nrows=60)

            header=detect_header(preview)

            if header is None:
                log("Skipping sheet due to missing header")
                continue

            df=pd.read_excel(file,sheet_name=sheet,header=header)

            df.columns=df.columns.str.strip().str.lower()

            df.insert(0,"source_file",os.path.basename(file))

            log(f"Rows loaded: {len(df)}")

            dfs.append(df)

    if not dfs:
        log("No data collected")
        return pd.DataFrame()

    df=pd.concat(dfs,ignore_index=True)

    log(f"Total rows before filtering: {len(df)}")

    if "currency" in df.columns:
        df["currency"]=norm(df["currency"])

    if "entity" in df.columns:
        df["entity"]=norm(df["entity"])

    if "global business" in df.columns:
        df["global business"]=norm(df["global business"])

    before=len(df)

    df=df[df["currency"]=="USD"]

    log(f"Rows after USD filter: {len(df)}")

    df=df[df["entity"].isin(entity_list)]

    log(f"Rows after entity filter: {len(df)}")

    df=df[df["global business"].isin(globalbusiness_list)]

    log(f"Rows after global business filter: {len(df)} (was {before})")

    return df

# ------------------------------------------------
# WRITE CONSOLIDATED FILE
# ------------------------------------------------

def write_excel(file_name,sheets):

    log(f"\nWriting file: {file_name}")

    wb=Workbook(write_only=True)

    for sheet,df in sheets.items():

        log(f"Writing sheet: {sheet}")

        ws=wb.create_sheet(sheet)

        if df.empty:
            log("Sheet empty")
            continue

        ws.append([c.title() for c in df.columns])

        count=0

        for r in df.itertuples(index=False):

            ws.append(list(r))
            count+=1

        log(f"{count} rows written")

    wb.save(file_name)

# ------------------------------------------------
# CREATE MAPPED FILE
# ------------------------------------------------

def create_mapped_copy(original):

    log("\nCreating mapped file")

    entity_map=pd.read_excel(mapper_file,sheet_name="Entity")
    product_map=pd.read_excel(mapper_file,sheet_name="Product")

    entity_map["Entity"]=norm(entity_map["Entity"])
    product_map["Product"]=norm(product_map["Product"])

    entity_dict=dict(zip(entity_map["Entity"],entity_map["Proxy Entity Code"]))
    product_dict=dict(zip(product_map["Product"],product_map["Proxy Product Code"]))

    wb=load_workbook(original)

    for sheet in wb.sheetnames:

        log(f"Processing mapping in sheet: {sheet}")

        ws=wb[sheet]

        headers=[c.value for c in ws[1]]

        if "Function" in headers:

            idx=headers.index("Function")+1
            ws.delete_cols(idx)

            log("Function column deleted")

            headers.pop(idx-1)

        if "Entity" in headers:

            e=headers.index("Entity")+1
            replace_count=0

            for r in range(2,ws.max_row+1):

                val=str(ws.cell(r,e).value).strip().upper()

                if val in entity_dict:
                    ws.cell(r,e).value=entity_dict[val]
                    replace_count+=1

            log(f"Entity replacements: {replace_count}")

        if "Product" in headers:

            p=headers.index("Product")+1
            replace_count=0

            for r in range(2,ws.max_row+1):

                val=str(ws.cell(r,p).value).strip().upper()

                if val in product_dict:
                    ws.cell(r,p).value=product_dict[val]
                    replace_count+=1

            log(f"Product replacements: {replace_count}")

    new=original.replace(".xlsx","_Mapped.xlsx")

    wb.save(new)

    log(f"Mapped file created: {new}")

    return new

# ------------------------------------------------
# CREATE MASKED FILE
# ------------------------------------------------

def create_masked_file(mapped):

    log("\nCreating masked file")

    wb=load_workbook(mapped)

    for sheet in wb.sheetnames:

        log(f"Masking sheet: {sheet}")

        ws=wb[sheet]

        headers=[c.value for c in ws[1]]

        for col in range(1,len(headers)+1):

            name=headers[col-1].lower()

            if name in ["year","entity","currency","product","source_file"]:
                continue

            for r in range(2,ws.max_row+1):

                val=ws.cell(r,col).value

                if isinstance(val,(int,float)):

                    if "pbt" in sheet.lower():

                        ws.cell(r,col).value=val/x_mask_value

                    else:

                        ws.cell(r,col).value=val/(x_mask_value/2)

    new=mapped.replace("_Mapped.xlsx","_Masked.xlsx")

    wb.save(new)

    log(f"Masked file created: {new}")

# ------------------------------------------------
# MAIN PROCESS
# ------------------------------------------------

def start():

    open(LOG_FILE,"w").close()

    log("PROCESS STARTED")

    tables={}

    for k in file_store:

        tables[k]=consolidate_category(k)

    output="Consolidated_Output.xlsx"

    write_excel(
        output,
        {
        "RWA Actuals":tables["RWA_Actuals"],
        "RWA Plan":tables["RWA_Plan"],
        "AVBS_SD Actuals":pd.concat([tables["AVBS_Actuals"],tables["SD_Actuals"]]),
        "AVBS_SD Plan":pd.concat([tables["AVBS_Plan"],tables["SD_Plan"]]),
        "PBT_BS Actuals":pd.concat([tables["PBT_Actuals"],tables["BS_Actuals"]]),
        "PBT_BS Plan":pd.concat([tables["PBT_Plan"],tables["BS_Plan"]])
        }
    )

    mapped=create_mapped_copy(output)

    create_masked_file(mapped)

    log("PROCESS COMPLETED")

    messagebox.showinfo("Done","All files created")

# ------------------------------------------------
# GUI
# ------------------------------------------------

root=tk.Tk()
root.title("Financial Consolidation Tool")
root.geometry("700x500")

frame=tk.Frame(root)
frame.pack(pady=20)

labels={}

def select_files(key):

    files=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xls")])

    if files:
        file_store[key]=list(files)
        labels[key].config(text=f"{len(files)} selected")

def select_mapper():

    global mapper_file

    mapper_file=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])

for i,k in enumerate(file_store):

    tk.Label(frame,text=k,width=20).grid(row=i,column=0)

    tk.Button(frame,text="Select",
              command=lambda x=k:select_files(x)).grid(row=i,column=1)

    lbl=tk.Label(frame,text="None")
    lbl.grid(row=i,column=2)

    labels[k]=lbl

tk.Button(root,text="Select Mapper",command=select_mapper).pack(pady=10)

tk.Button(root,text="Run",command=start,bg="green",fg="white").pack(pady=20)

root.mainloop()
