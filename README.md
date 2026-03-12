```
import win32com.client as win32
import pythoncom
import time


def refresh_workbook_sheetwise(file_path, value):

    pythoncom.CoInitialize()

    excel = win32.gencache.EnsureDispatch("Excel.Application")

    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        # Disable macros
        excel.AutomationSecurity = 3

        print("Opening workbook...")

        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )

        print("Workbook opened")

        # Update value
        sheet = wb.Worksheets("Landing Page DB")
        sheet.Range("F1").Value = value

        print("Updated Landing Page DB!F1")

        # Set manual calculation so Excel doesn't auto calculate everything
        excel.Calculation = -4135  # xlCalculationManual

        print("Starting sheet-by-sheet recalculation...\n")

        for sheet in wb.Worksheets:

            sheet_name = sheet.Name
            print(f"Refreshing sheet: {sheet_name}")

            sheet.Calculate()

            while excel.CalculationState != 0:
                time.sleep(0.5)

            print(f"Finished: {sheet_name}\n")

        print("All sheets refreshed")

        wb.Save()
        wb.Close(False)

        print("Workbook saved and closed")

    except Exception as e:
        print("Error:", e)

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    print("Process completed successfully")


# Example usage
file_path = r"C:\path\to\your\workbook.xlsx"

refresh_workbook_sheetwise(file_path, "New String Value")
