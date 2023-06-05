"轉換路徑中.xlsb檔案為.xlsx"

import os
import fnmatch
import pandas as pd
import pyxlsb
def convert_xlsb(file_path):
    "轉換路徑中.xlsb檔案為.xlsx"
    for file in os.listdir(file_path):
        if fnmatch.fnmatch(file, '*.xlsb'):
            file_path =os.path.join(file_path, file)
            print("找到檔案：", file_path)
    #顯示表格
    data_frame = pd.read_excel(file_path, sheet_name='Data1', engine='pyxlsb')
    xlsx_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
    print("輸出檔案：", xlsx_path)
    #輸出檔案
    data_frame.to_excel(xlsx_path)
    print(file_path)
    input("暫停一下")
    return file_path

PATH1 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT\HT_HT-D1-CTC-GEL-23-1171"
PATH2 = r"/workspaces/Auto-Work-Station/HT"
#print("預設路徑:",path)
#path = input("請輸入路徑:")
convert_xlsb(PATH2)
