#cd C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station
#pyinstaller -F Convert_Excel.py
#輸入路徑解析xlsb成.xlsx
import pandas as pd
import datetime
import os
import fnmatch
import pyxlsb
def Convert_xlsb(File_Path):
    for file in os.listdir(File_Path):
        if fnmatch.fnmatch(file, '*.xlsb'):
            File_Path =os.path.join(File_Path, file)
            print("找到檔案：", File_Path)
    #顯示表格
    df = pd.read_excel(File_Path, sheet_name='Data1', engine='pyxlsb')
    Xlsx_File_Path = os.path.splitext(File_Path)[0] + "_converted.xlsx"
    print("輸出檔案：", Xlsx_File_Path)
    #輸出檔案
    df.to_excel(Xlsx_File_Path)
    print(File_Path)
    input("暫停一下")
    return File_Path

path = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT\HT_HT-D1-CTC-GEL-23-1171"
print("預設路徑:",path)
path = input("請輸入路徑:")
Convert_xlsb(path)