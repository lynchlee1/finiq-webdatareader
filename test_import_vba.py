import win32com.client
import os

try:
    try:
        # Try to get active Excel instance
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("Connected to active Excel instance.")
    except Exception:
        # Fallback to dispatching new instance
        excel = win32com.client.Dispatch("Excel.Application")
        print("Dispatched new Excel instance.")
        
    file_path = os.path.abspath(r"opendart-to-excel/example.xlsm")
    file_name = os.path.basename(file_path).lower()
    
    # Check if workbook is already open
    wb = None
    for open_wb in excel.Workbooks:
        if open_wb.Name.lower() == file_name:
            wb = open_wb
            print("Found already open workbook:", wb.Name)
            break
            
    if wb is None:
        print("Opening file:", file_path)
        wb = excel.Workbooks.Open(file_path)
        
    try:
        proj_name = wb.VBProject.Name
        print("Success! VBProject name:", proj_name)
        for comp in wb.VBProject.VBComponents:
            print("Component:", comp.Name, comp.Type)
    except Exception as e:
        print("Failed to access VBProject:", e)
        
    # We do NOT close the workbook or Excel if it was already open, to avoid disrupting the user.
    # But if we opened it, we can save and close.
    # For now, let's not close anything.
except Exception as e:
    print("Global Error:", e)
