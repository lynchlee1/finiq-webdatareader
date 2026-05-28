import win32com.client
import os
import subprocess

def kill_excel():
    print("Terminating Excel processes to release file locks...")
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], capture_output=True)
        print("Excel processes terminated.")
    except Exception as e:
        print("Error terminating Excel:", e)

def import_bas():
    file_path = os.path.abspath(r"opendart-to-excel/example.xlsm")
    bas_path = os.path.abspath(r"opendart-to-excel/OpenDartDownloader.bas")
    
    if not os.path.exists(file_path):
        print(f"Error: Workbook not found at {file_path}")
        return False
    if not os.path.exists(bas_path):
        print(f"Error: .bas file not found at {bas_path}")
        return False
        
    kill_excel()
    
    excel = None
    try:
        # Start a fresh instance of Excel
        print("Starting fresh Excel instance...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        print(f"Opening workbook: {file_path}")
        wb = excel.Workbooks.Open(file_path)
        print("Workbook opened. Is Read-Only:", wb.ReadOnly)
        
        if wb.ReadOnly:
            print("Warning: Workbook is still read-only! Cannot save changes.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return False
            
        # Check if project is protected
        if wb.VBProject.Protection == 1:
            print("Error: The VBA project of this workbook is protected. Cannot import module.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return False
            
        # Remove existing module if it exists
        module_name = "OpenDartDownloader"
        for comp in list(wb.VBProject.VBComponents):
            if comp.Name == module_name:
                print(f"Removing existing module: {module_name}")
                wb.VBProject.VBComponents.Remove(comp)
                break
                
        # Import .bas file
        print(f"Importing VBA module from: {bas_path}")
        wb.VBProject.VBComponents.Import(bas_path)
        
        # Save and close
        print("Saving workbook...")
        wb.Save()
        print("Workbook saved successfully with new VBA module.")
        
        wb.Close(SaveChanges=True)
        excel.Quit()
        return True
    except Exception as e:
        print("Error during import:")
        import traceback
        traceback.print_exc()
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        return False

if __name__ == "__main__":
    import_bas()
