import win32com.client
import sys
import traceback

try:
    print("Attempting DispatchEx...")
    excel = win32com.client.DispatchEx("Excel.Application")
    print("DispatchEx succeeded.")
    print("Excel version:", excel.Version)
    wbs = excel.Workbooks
    print("Workbooks count:", wbs.Count)
    excel.Quit()
except Exception as e:
    print("Error with DispatchEx:")
    traceback.print_exc()
