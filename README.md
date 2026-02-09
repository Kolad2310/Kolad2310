```
import xlwings as xw
import os

# User inputs
workbook_path = r'C:\path\to\your\workbook.xlsx'  # Adjust path
l = ['value1', 'value2', 'value3']  # Your list
sheet_to_refresh = ['Landing Page DB', 'Sheet1', 'Sheet2']  # Only these
output_dir = r'C:\path\to\output'  # Folder for versions
os.makedirs(output_dir, exist_ok=True)

# Open workbook (invisible)
app = xw.App(visible=False)
wb = app.books.open(workbook_path)

# Get Landing Page DB sheet
landing_sheet = wb.sheets['Landing Page DB']

for i, value in enumerate(l):
    # Update F1
    landing_sheet.range('F1').value = value
    
    # Recalculate only target sheets (triggers formula refresh)
    for sheet_name in sheet_to_refresh:
        sheet = wb.sheets[sheet_name]
        sheet.api.Calculate()  # Sheet-level recalc[web:13]
    
    # Create new wb with copied refreshed sheets (preserves formatting)
    new_wb = app.books.add()
    for sheet_name in sheet_to_refresh:
        wb.sheets[sheet_name].api.Copy(Before=new_wb.sheets[0].api)
    
    # Delete default sheet
    new_wb.sheets[0].delete()
    
    # Save versioned file
    versioned_name = f'version_{value}_{i+1}.xlsx'
    new_wb.save(os.path.join(output_dir, versioned_name))
    new_wb.close()

# Cleanup
wb.close()
app.quit()
