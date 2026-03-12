```
import win32com.client as win32

def update_excel_cell(file_path, value):

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(file_path)

        sheet = wb.Worksheets("Landing Page DB")

        # Update cell F1
        sheet.Range("F1").Value = value

        # Recalculate workbook
        excel.CalculateFull()

        wb.Save()
        wb.Close()

    finally:
        excel.Quit()

    print("Cell F1 updated successfully")


file_path = r"C:\path\your_file.xlsx"

update_excel_cell(file_path, "New String Value")
