""""
轉換路徑中.xlsb檔案為.xlsx
前置作業
pip install pyxlsb
pip install openpyxl
"""
import os
import fnmatch
import pandas as pd
import pyxlsb
import openpyxl
def convert_xlsb(file_path):
    "尋找檔案"
    for file in os.listdir(file_path):
        if fnmatch.fnmatch(file, '*.xlsb'):
            file_path =os.path.join(file_path, file)
            print("找到目標檔案：", file_path)
    #若找不到檔案，則回傳None
    if not os.path.exists(file_path):
        print("找不到檔案：", file_path)
        return None
    else:
        "篩選是否符合檔案格式條件"
        #若已經轉換過了，則回傳None
        if "~$" in file_path:
            print(file_path,"抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
            return None 
        else:
            data_frame = pd.read_excel(file_path, sheet_name='Data1', engine='pyxlsb')
            xlsx_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
            print("輸出檔案：", xlsx_path)
            if os.path.exists(xlsx_path):
                print("檔案已存在，不動作")
                return None
            else:
                data_frame.to_excel(xlsx_path)
                input("enter any keys to exit")
                return xlsx_path

def main():
    "主要執行內容"
    PATH1 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT\HT-D1-CTC-GEL-23-1169"
    PATH2 = r"/workspaces/Auto-Work-Station/HT"
    PATH_Home = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\HT"
    #print("預設路徑:",path)
    #path = input("請輸入路徑:")
    convert_xlsb(PATH1)

if __name__ == '__main__':
    main()

