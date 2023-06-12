"""
主程式
前置作業
pip install pyxlsb
pip install openpyxl
cd path
pyinstaller -F test.py
"""
import sys
import os
import doc_collection
import convert_excel
import read_ht_excel
import read_tc_excel
import text_gen
from constants import (CLOUD_HT, CLOUD_TC, COM_DESTINY_HT, COM_DESTINY_TC, DOC_NO_STRUCTURE, 
                       HOME_SOURCE_HT, HOME_DESTINY_HT, HOME_SOURCE_TC, HOME_DESTINY_TC, 
                       PLAN_LIST, COM_SOURCE_HT, COM_SOURCE_TC)

ROOT_HOME = r"C:/Users/OXO/OneDrive/01 Book/00 Test program/Auto-Work-Station"
ROOT_GIT = r"/workspaces/Auto-Work-Station"
sys.path.append(ROOT_HOME)

def main():
    "主程式"
    while True:
        letter_num, source_folder, dest_folder = doc_collection.file_path_process()
        doc_collection.move_docutment(source_folder, dest_folder)

        if "TC(C0)" in dest_folder:
            print("開始讀取「台中」檔案")
            letter_title, letter_vision, letter_num, letter_date = read_tc_excel.read_tc_excel(dest_folder)
        elif "HT" in dest_folder:
            print("開始讀取「興達」檔案")
            converted_xlsx_path = convert_excel.convert_xlsb(dest_folder)
            letter_title, letter_vision, letter_num, letter_date = read_ht_excel.read_ctc_ht_excel(dest_folder)

        else:
            print("找不到「計畫」關鍵字，無動作")
            return None
        print("dest_folder:", dest_folder)
        text_gen.text_gen(letter_title, letter_vision, letter_num, letter_date, dest_folder)
        input("完成，請按任意鍵Press enter to exit...")

def update_value(new_value1, new_value2, new_value3):
    "更新變數值"
    #global HT_path
    #HT_path = new_value
    return None

def main2():
    """[summary]"""
    return None

if __name__ == '__main__':
    main()
