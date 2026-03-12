```
import win32com.client as win32
import pythoncom
import time


def refresh_heavy_excel(file_path, value):

    pythoncom.CoInitialize()

    excel = win32.DispatchEx("Excel.Application")

    try:
        print("Starting Excel")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        # Disable macros
        excel.AutomationSecurity = 3

        # Use constants
        xlManual = -4135

        print("Opening workbook...")

        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )

        print("Workbook opened")

        # Now change calculation mode (after open)
        excel.Calculation = xlManual

        sheet = wb.Worksheets("Landing Page DB")

        print("Updating F1")

        sheet.Range("F1").Value = value

        print("Starting calculation")

        excel.Calculate()

        while excel.CalculationState != 0:
            time.sleep(1)

        print("Calculation completed")

        wb.Save()
        wb.Close(False)

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    print("Process completed")


file_path = r"C:\path\your_file.xlsx"

refresh_heavy_excel(file_path, "New String Value")
