```
import xlwings as xw
import multiprocessing as mp
from functools import partial
import os
from concurrent.futures import ProcessPoolExecutor

# User inputs
workbook_path = r'C:\path\to\your\workbook.xlsx'  # Shared read-only source
l = ['entity1', 'entity2', ..., 'entity31']  # Exactly 31 entities
sheet_to_refresh = ['Landing Page DB', 'Sheet1', 'Sheet2']  # Adjust
output_dir = r'C:\path\to\output'
os.makedirs(output_dir, exist_ok=True)
num_processes = 8  # Tune: 4-12 based on CPU cores (e.g., 8 for i7)

def process_entity(entity, idx, workbook_path, sheet_to_refresh, output_dir):
    """Worker function: creates one versioned file"""
    app = xw.App(visible=False, screen_updating=False, display_alerts=False)
    try:
        wb = app.books.open(workbook_path)
        landing_sheet = wb.sheets['Landing Page DB']
        
        # Update F1
        landing_sheet.range('F1').value = entity
        
        # Recalc only targets
        for sheet_name in sheet_to_refresh:
            sheet = wb.sheets[sheet_name]
            sheet.api.Calculate()
        
        # New wb with copies
        new_wb = app.books.add()
        for sheet_name in sheet_to_refresh:
            wb.sheets[sheet_name].api.Copy(Before=new_wb.sheets[0].api)
        new_wb.sheets[0].delete()
        
        # Save
        versioned_name = f'version_{entity}_{idx+1}.xlsx'
        new_wb.save(os.path.join(output_dir, versioned_name))
        new_wb.close()
        wb.close()
        print(f"Completed {versioned_name}")  # Progress
    except Exception as e:
        print(f"Error for {entity}: {e}")
    finally:
        app.quit()

if __name__ == '__main__':
    # Prepare args
    process_func = partial(process_entity, workbook_path=workbook_path,
                           sheet_to_refresh=sheet_to_refresh, output_dir=output_dir)
    
    with ProcessPoolExecutor(max_workers=num_processes) as executor:
        futures = [executor.submit(process_func, entity, idx) 
                   for idx, entity in enumerate(l)]
        # Wait for all
        for future in futures:
            future.result()  # Raises if errors
    print("All 31 versions complete!")
