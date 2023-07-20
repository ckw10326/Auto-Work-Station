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
import read_ht_excel
import read_tc_excel
import text_gen
import constants#改成seettings


ROOT_HOME = r"C:/Users/OXO/OneDrive/01 Book/00 Test program/Auto-Work-Station"
ROOT_GIT = r"/workspaces/Auto-Work-Station"
sys.path.append(ROOT_GIT)

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

'''
雲端分析功能
'''
def main2():
    path = r"/workspaces/Auto-Work-Station/00source"
    x = read_ht_excel.convert_xlsb(path)
    if x :
        letter_title, letter_vision, letter_num, letter_date = read_ht_excel.read_ctc_ht_excel(path)
        text_gen.text_gen(letter_title, letter_vision, letter_num, letter_date, path)
    else :
        print("read_ht_excel.convert_xlsb回傳None，無符合興達分析之檔案，現在開始分析「台中」")
        letter_title, letter_vision, letter_num, letter_date = read_tc_excel.read_tc_excel(path)
        text_gen.text_gen(letter_title, letter_vision, letter_num, letter_date, path)
'''
給定資料夾一次將所有excel檔案分析完成，已經分析過得不再處理
'''

def update_value(new_value1, new_value2, new_value3):
    "更新變數值"
    #global HT_path
    #HT_path = new_value
    return None


if __name__ == '__main__':
    main2()
