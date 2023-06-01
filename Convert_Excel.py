#cd C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station
#pyinstaller -F Convert_Excel.py
#輸入路徑解析xlsb成.xlsx
import os
import fnmatch
import pandas as pd
def Convert_xlsb(File_Path):
    for file in os.listdir(File_Path):
        if fnmatch.fnmatch(file, '*.xlsb'):
            File_Path =os.path.join(File_Path, file)
            print("找到檔案：", File_Path)
    #顯示表格
    df = pd.read_excel(File_Path, sheet_name='Data1', engine='pyxlsb')

    Xlsx_File_Path = os.path.splitext(File_Path)[0] + ".xlsx"
    print("輸出檔案：", Xlsx_File_Path)
    #輸出檔案
    df.to_excel(Xlsx_File_Path)
    return File_Path

Test_Path = input("請輸入路徑:")
Convert_xlsb(Test_Path)
input("Press enter to exit...")
