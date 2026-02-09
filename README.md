```
import xlwings as xw
import os
import time

# Test inputs - UPDATE THESE
workbook_path = r'C:\path\to\your\workbook.xlsx'  # ← CRITICAL: Fix this path
output_dir = r'C:\path\to\output'
l = ['test1', 'test2']  # Start small
sheet_to_refresh = ['Landing Page DB']  # Your exact sheet names

print(f"Starting debug... Path exists: {os.path.exists(workbook_path)}")
print(f"Sheet names to refresh: {sheet_to_refresh}")

try:
    app = xw.App(visible=True)  # VISIBLE for debugging
    print("✓ Excel app launched")
    
    wb = app.books.open(workbook_path)
    print(f"✓ Workbook opened: {wb.name}")
    
    # List ALL sheets to verify names
    print("Available sheets:", [s.name for s in wb.sheets])
    
    landing_sheet = wb.sheets['Landing Page DB']
    print("✓ Landing sheet found")
    
    for i, value in enumerate(l):
        print(f"\n--- Processing {value} (#{i+1}) ---")
        start_time = time.time()
        
        # Update F1
        landing_sheet.range('F1').value = value
        print(f"✓ F1 updated to: {value}")
        
        # Force full recalc (safer for debug)
        wb.api.Calculate()
        print("✓ Workbook recalculated")
        
        # Simple save-as instead of complex copy
        versioned_name = f'version_{value}_{i+1}.xlsx'
        output_path = os.path.join(output_dir, versioned_name)
        wb.save(output_path)
        print(f"✓ SAVED: {versioned_name} ({time.time()-start_time:.1f}s)")
    
    wb.close()
    app.quit()
    print("\n🎉 SUCCESS! Check output folder.")
    
except Exception as e:
    print(f"❌ ERROR: {e}")
    print("Check: 1) File path, 2) Excel closed, 3) Sheet names exact")
