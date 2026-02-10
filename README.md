```
import win32com.client as win32
import os

def break_excel_links(file_path, save_as=None):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    file_path = os.path.abspath(file_path)

    wb = excel.Workbooks.Open(file_path, UpdateLinks=0)

    links = wb.LinkSources(Type=1)  # xlLinkTypeExcelLinks = 1

    if links:
        for link in links:
            wb.BreakLink(Name=link, Type=1)

    # Save logic
    if save_as:
        wb.SaveAs(os.path.abspath(save_as))
    else:
        wb.Save()

    wb.Close(False)
    excel.Quit()

    return bool(links)
