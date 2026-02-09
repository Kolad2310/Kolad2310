```
import xlwings as xw
import os
import time

# === UPDATE THESE PATHS AND LISTS ===
workbook_path = r'C:\path\to\your\workbook.xlsx'  # YOUR SOURCE FILE
output_dir = r'C:\path\to\output'                 # OUTPUT FOLDER
l = ['entity1', 'entity2', 'entity3']             # YOUR 31 ENTITIES HERE
sheet_to_refresh = ['Landing Page DB', 'Sheet1', 'Sheet2']  # EXACT SHEET NAMES

# Create output folder
os.makedirs(output_dir, exist_ok=True)

print("Starting Excel automation (NO VISIBLE WORKBOOK)...")
print(f"Source: {workbook_path}")
print(f"Output: {output_dir}")
print(f"Entities: {len(l)}")
print(f"Sheets: {sheet_to_refresh}")
print("-" * 50)

# Launch Excel COMPLETELY HIDDEN
app = xw.App(visible=False, screen_updating=False, display_alerts=False, add_book=False)
print("✓ Excel launched (invisible)")

try:
    # Open source workbook (stays hidden)
    wb = app.books.open(workbook_path)
    print(f"✓ Workbook opened: {wb.name} (hidden)")
    
    # Show all available sheets
    print("Available sheets:", [s.name for s in wb.sheets])
    
    # Get landing sheet
    landing_sheet = wb.sheets['Landing Page DB']
    print("✓ Landing Page DB found")
    
    # Process each entity
    for i, entity in enumerate(l):
        print(f"\n[{i+1}/{len(l)}] Processing '{entity}'...")
        start_time = time.time()
        
        # Step 1: Update F1
        landing_sheet.range('F1').value = entity
        print(f"  ✓ F1 = {entity}")
        
        # Step 2: Recalculate target sheets only
        for sheet_name in sheet_to_refresh:
            sheet = wb.sheets[sheet_name]
            sheet.api.Calculate()
        print(f"  ✓ Recalculated {len(sheet_to_refresh)} sheets")
        
        # Step 3: Create NEW workbook (also hidden)
        new_wb = app.books.add()
        
        # Copy each target sheet (preserves formatting/formulas)
        for j, sheet_name in enumerate(sheet_to_refresh):
            source_sheet = wb.sheets[sheet_name]
            source_sheet.api.Copy(Before=new_wb.sheets[0].api)
            print(f"  ✓ Copied {sheet_name}")
        
        # Delete default empty sheet
        new_wb.sheets[0].delete()
        
        # Step 4: Save versioned file
        filename = f"version_{entity}_{i+1:02d}.xlsx"
        output_path = os.path.join(output_dir, filename)
        new_wb.save(output_path)
        new_wb.close()
        
        elapsed = time.time() - start_time
        print(f"  ✓ SAVED: {filename} ({elapsed:.1f}s)")
    
    # Cleanup
    wb.close()
    app.quit()
    print("\n🎉 ALL FILES COMPLETED!")
    print(f"Check folder: {output_dir}")

except Exception as e:
    print(f"\n❌ ERROR: {e}")
    print("FIX:")
    print("1. Check file path exists")
    print("2. Close all Excel instances BEFORE running") 
    print("3. Verify exact sheet names above")
    if 'app' in locals():
        try:
            app.quit()
        except:
            pass

input("\nPress Enter to exit...")  # Keeps console open
