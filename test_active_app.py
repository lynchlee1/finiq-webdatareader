import win32com.client
import sys
import traceback

try:
    obj = win32com.client.GetActiveObject("Excel.Application")
    print("Got object from GetActiveObject:", type(obj))
    # Access .Application property
    excel = obj.Application
    print("Accessed .Application, type:", type(excel))
    print("Excel version:", excel.Version)
    print("Workbooks count:", excel.Workbooks.Count)
    for wb in excel.Workbooks:
        print("Workbook:", wb.Name)
except Exception as e:
    traceback.print_exc()
