```
import subprocess
import tempfile
import os

def refresh_excel(file_path):

    ps_script = f"""
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.Calculation = -4135

    $wb = $excel.Workbooks.Open("{file_path}")

    $wb.RefreshAll()

    while ($excel.CalculationState -ne 0) {{
        Start-Sleep -Seconds 1
    }}

    $excel.CalculateFullRebuild()

    $wb.Save()
    $wb.Close($true)

    $excel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    """

    # Write PowerShell script to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".ps1") as f:
        f.write(ps_script.encode())
        ps_path = f.name

    # Run PowerShell
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path], check=True)

    os.remove(ps_path)

    print("Excel refreshed successfully")


# Usage
refresh_excel(r"C:\path\to\your\file.xlsx")
