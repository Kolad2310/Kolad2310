```
import os
import win32com.client as win32

SOURCE_FOLDER = r"C:\input_excels"
TARGET_FOLDER = r"C:\value_excels"

os.makedirs(TARGET_FOLDER, exist_ok=True)

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

for file in os.listdir(SOURCE_FOLDER):
    if file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        src_path = os.path.join(SOURCE_FOLDER, file)
        tgt_path = os.path.join(TARGET_FOLDER, file)

        wb = excel.Workbooks.Open(src_path)

        for sheet in wb.Worksheets:
            used_range = sheet.UsedRange
            used_range.Value = used_range.Value  # 🔥 formula → value

        wb.SaveAs(tgt_path)
        wb.Close(False)

excel.Quit()

print("✅ All files saved as value versions successfully")
