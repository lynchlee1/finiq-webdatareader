import win32com.client
import os

try:
    excel = win32com.client.Dispatch("Excel.Application")
    print("Excel version:", excel.Version)
    excel.Quit()
except Exception as e:
    print("Error dispatching Excel:", e)
