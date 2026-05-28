import win32com.client
import os
import subprocess
import tempfile

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
                
        # Excel's VBA importer reads .bas text through the local ANSI code page.
        # Keep the repository source as UTF-8, but import a CP949 temp copy so
        # Korean VBA string literals are preserved in Excel.
        print(f"Preparing CP949 VBA import copy from UTF-8 source: {bas_path}")
        with open(bas_path, "r", encoding="utf-8-sig", newline=None) as bas_file:
            module_code = bas_file.read()

        temp_bas_path = None
        try:
            with tempfile.NamedTemporaryFile(
                "w",
                suffix=".bas",
                delete=False,
                encoding="cp949",
                newline="\r\n",
            ) as temp_bas:
                temp_bas.write(module_code)
                temp_bas_path = temp_bas.name

            print(f"Importing VBA module from CP949 temp file: {temp_bas_path}")
            wb.VBProject.VBComponents.Import(temp_bas_path)
        finally:
            if temp_bas_path and os.path.exists(temp_bas_path):
                os.remove(temp_bas_path)
        
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
