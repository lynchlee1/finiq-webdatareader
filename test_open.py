import win32com.client
import os

try:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    file_path = os.path.abspath(r"opendart-to-excel/example.xlsm")
    print("Opening file:", file_path)
    wb = excel.Workbooks.Open(file_path)
    
    print("Opened successfully.")
    print("Is Read Only:", wb.ReadOnly)
    
    wb.Close(SaveChanges=False)
    excel.Quit()
except Exception as e:
    print("Error opening file:", e)
