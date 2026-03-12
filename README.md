```
import win32com.client as win32
import pythoncom
import time


def refresh_selected_sheets_formula_only(file_path, value, sheet_list):

    pythoncom.CoInitialize()

    excel = win32.gencache.EnsureDispatch("Excel.Application")

    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3

        print("Opening workbook...")

        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )

        print("Workbook opened")

        # Disable automatic recalculation
        excel.Calculation = -4135

        # Update F1
        wb.Worksheets("Landing Page DB").Range("F1").Value = value
        print("Updated Landing Page DB!F1")

        print("\nStarting formula-only recalculation...\n")

        xlCellTypeFormulas = -4123

        for sheet_name in sheet_list:

            try:
                sheet = wb.Worksheets(sheet_name)

                print(f"Refreshing sheet: {sheet_name}")

                used = sheet.UsedRange

                try:
                    formula_cells = used.SpecialCells(xlCellTypeFormulas)
                    formula_cells.Calculate()
                except:
                    print("No formula cells found")

                while excel.CalculationState != 0:
                    time.sleep(0.5)

                print(f"Finished: {sheet_name}\n")

            except Exception as e:
                print(f"Error refreshing {sheet_name}: {e}")

        wb.Save()
        wb.Close(False)

        print("Workbook saved and closed")

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    print("Process completed")


file_path = r"C:\path\file.xlsx"

sheets = [
    "Landing Page DB",
    "Revenue Model",
    "Dashboard"
]

refresh_selected_sheets_formula_only(file_path, "New Value", sheets)
