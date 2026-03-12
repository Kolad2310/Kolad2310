```
import win32com.client as win32
import pythoncom
import time
import os


def refresh_heavy_excel(file_path, value):

    pythoncom.CoInitialize()

    excel = win32.DispatchEx("Excel.Application")

    try:
        print("Starting Excel...")

        # Run Excel silently
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        # Disable macros
        excel.AutomationSecurity = 3

        # Manual calculation while loading
        excel.Calculation = -4135  # xlCalculationManual

        # Enable multi-threaded calculation
        try:
            excel.MultiThreadedCalculation.Enabled = True
        except:
            pass

        print("Opening workbook...")

        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )

        print("Workbook opened")

        sheet = wb.Worksheets("Landing Page DB")

        print("Updating cell F1")

        sheet.Range("F1").Value = value

        print("Triggering dependency recalculation...")

        # Faster than full rebuild
        excel.Calculate()

        # Wait for calculation to complete
        while excel.CalculationState != 0:
            time.sleep(1)

        print("Calculation completed")

        wb.Save()

        print("Workbook saved")

        wb.Close(False)

    except Exception as e:
        print("Error:", e)

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    print("Process completed")


# Example usage
file_path = r"C:\path\your_file.xlsx"

refresh_heavy_excel(file_path, "New Value")
