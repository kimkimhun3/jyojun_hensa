import sys
from win32com.client import DispatchEx
import os
import math

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python python.py <filePath>")
        sys.exit(1)

    filePath = sys.argv[1]
    savePath = filePath.replace("/", "\\")
    base, ext = os.path.splitext(savePath)

    # Initialize Excel and open the workbook
    xl = DispatchEx("Excel.Application")
    xl.Visible = False  # Run Excel in the background
    wb = xl.Workbooks.Open(savePath)

    wb.Save()

    wb.Close(SaveChanges=1)
    xl.Quit()
    print("Task Done")
    print(f"File saved: {savePath}")

