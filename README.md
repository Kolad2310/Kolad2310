```
import win32com.client as win32
import os
import sys


def break_excel_links(
    input_file,
    output_file=None,
    visible=False
):
    """
    Breaks all Excel external links in a workbook safely.

    Parameters
    ----------
    input_file : str
        Path to source Excel file
    output_file : str | None
        If None → overwrite input file
        Else → save to this path
    visible : bool
        Show Excel UI (debug)

    Returns
    -------
    dict with link details
    """

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False

    input_file = os.path.abspath(input_file)

    wb = excel.Workbooks.Open(
        input_file,
        UpdateLinks=0,   # DO NOT auto-update
        ReadOnly=False
    )

    result = {
        "file": input_file,
        "links_found": False,
        "links": []
    }

    # xlLinkTypeExcelLinks = 1
    links = wb.LinkSources(Type=1)

    if links:
        result["links_found"] = True
        result["links"] = list(links)

        for link in links:
            wb.BreakLink(Name=link, Type=1)

    # Save logic
    if output_file:
        wb.SaveAs(os.path.abspath(output_file))
    else:
        wb.Save()

    wb.Close(False)
    excel.Quit()

    return result


# ------------------ RUN DIRECTLY ------------------
if __name__ == "__main__":

    INPUT_FILE = r"C:\data\input.xlsx"
    OUTPUT_FILE = r"C:\data\input_no_links.xlsx"  # set None to overwrite

    result = break_excel_links(
        INPUT_FILE,
        OUTPUT_FILE,
        visible=False
    )

    if result["links_found"]:
        print("✅ External links were found and broken:")
        for l in result["links"]:
            print("   -", l)
    else:
        print("✅ No external links found")
