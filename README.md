```
import xlwings as xw
import os
import time
import shutil

# === UPDATE THESE PATHS AND LISTS ===
workbook_path = r'C:\path\to\your\workbook.xlsx'  # YOUR SOURCE FILE
output_dir = r'C:\path\to\output'                 # OUTPUT FOLDER  
l = ['entity1', 'entity2', 'entity3']             # YOUR 31 ENTITIES HERE
sheet_to_refresh = ['Landing Page DB', 'Sheet1', 'Sheet2']  # EXACT SHEET NAMES

# Create output folder
os.makedirs(output_dir, exist_ok=True)
print(f"✓ Output folder ready: {output_dir}")

print("Starting Excel automation...")
print(f"Source: {workbook_path}")
print(f"Entities: {len(l)}")
print("-" * 50)

app = xw.App(visible=False)
app.display_alerts = False

try:
    wb = app.books.open(workbook_path)
    print(f"✓ Workbook loaded")
    print("Sheets:", [s.name for s in wb.sheets])

    landing_sheet = wb.sheets['Landing Page DB']

    for i, entity in enumerate(l):
        print(f"\n[{i+1}/{len(l)}] '{entity}'...")
        start_time = time.time()
        
        # Update F1
        landing_sheet.range('F1').value = entity
        
        # Force recalc ALL workbook (safer)
        wb.api.Calculate()
        time.sleep(0.5)  # Let calcs finish
        
        # Create filename
        filename = f"version_{entity}_{i+1:02d}.xlsx"
        output_path = os.path.join(output_dir, filename)
        
        # **SIMPLE SAVEAS** - no complex copying
        wb.save(output_path)
        wb.api.Save()  # Force native save
        
        # Verify file exists and has size
        if os.path.exists(output_path):
            size = os.path.getsize(output_path) / 1024  # KB
            print(f"  ✓ SAVED: {filename} ({size:.1f} KB, {time.time()-start_time:.1f}s)")
        else:
            print(f"  ❌ File not found: {output_path}")

    print("\n🎉 ALL PROCESSED!")
    
except Exception as e:
    print(f"\n❌ ERROR: {e}")

finally:
    # **FORCE CLEANUP**
    try:
        wb.close()
        app.quit()
        print("✓ Excel closed")
    except:
        pass

print(f"\nCheck folder: {output_dir}")
print("List of files:")
for f in os.listdir(output_dir):
    size = os.path.getsize(os.path.join(output_dir, f)) / 1024
    print(f"  {f} ({size:.1f} KB)")

input("Press Enter to exit...")
