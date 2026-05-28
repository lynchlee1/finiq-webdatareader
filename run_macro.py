import win32com.client
import os

def test_macro():
    file_path = os.path.abspath(r"opendart-to-excel/example.xlsm")
    
    excel = None
    try:
        print("Starting Excel instance...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        print(f"Opening workbook: {file_path}")
        wb = excel.Workbooks.Open(file_path)
        
        print("Running macro: InitializeCorpCodes")
        excel.Run("InitializeCorpCodes")
        print("InitializeCorpCodes completed.")
        
        print("Running macro: DownloadDartData")
        # Run the macro
        excel.Run("DownloadDartData")
        print("Macro completed.")
        
        # Check output sheets
        sheet_names = [sheet.Name for sheet in wb.Worksheets]
        print("Sheet names in workbook:", sheet_names)
        
        if "OpenDART" in sheet_names:
            ws = wb.Worksheets("OpenDART")
            print("\n--- OpenDART Sheet Contents (First 30 rows, 15 columns) ---")
            for r in range(1, 31):
                row_vals = []
                for c in range(1, 16):
                    val = ws.Cells(r, c).Value
                    row_vals.append(str(val) if val is not None else "")
                # Print if row has any content
                if any(row_vals):
                    print(f"Row {r:2d}: {row_vals}")
            print("---------------------------------------------------------")
        else:
            print("Error: 'OpenDART' sheet was not found in workbook.")
            
        print("Saving workbook...")
        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        print("Finished test successfully.")
        
    except Exception as e:
        print("Error running macro:")
        import traceback
        traceback.print_exc()
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    test_macro()
