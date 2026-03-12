```
import win32com.client as win32
import pythoncom
import time


def refresh_selected_sheets(file_path, value, l):

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

        # Prevent automatic calculation
        excel.Calculation = -4135  # xlCalculationManual

        # Update F1
        landing_sheet = wb.Worksheets("Landing Page DB")
        landing_sheet.Range("F1").Value = value

        print("Updated Landing Page DB!F1")

        print("\nStarting sheet refresh...\n")

        for sheet_name in l:

            try:
                sheet = wb.Worksheets(sheet_name)

                print(f"Refreshing sheet: {sheet_name}")

                used_range = sheet.UsedRange

                # Calculate only used cells
                used_range.Calculate()

                while excel.CalculationState != 0:
                    time.sleep(0.5)

                print(f"Finished refreshing: {sheet_name}\n")

            except Exception as e:
                print(f"Error refreshing {sheet_name}: {e}")

        print("All requested sheets refreshed")

        wb.Save()
        wb.Close(False)

        print("Workbook saved and closed")

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    print("Process completed")


# Example usage
file_path = r"C:\path\to\file.xlsx"

l = [
    "Landing Page DB",
    "Revenue Model",
    "Dashboard"
]

refresh_selected_sheets(file_path, "New Value", l)
