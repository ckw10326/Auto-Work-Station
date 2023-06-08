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

ROOT_HOME = r"C:/Users/OXO/OneDrive/01 Book/00 Test program/Auto-Work-Station"
ROOT_GIT = r"/workspaces/Auto-Work-Station"
sys.path.append(ROOT_GIT)



def main():
    "主程式"
    while True:
        global letter_num, source_folder, dest_folder
        global letter_title, letter_vision, letter_num, letter_date
        letter_num, source_folder, dest_folder = doc_collection.file_path_process()
        print("1",letter_num)
        print("2",source_folder)
        print("3",dest_folder)
        doc_collection.move_docutment(source_folder, dest_folder)
        os.startfile(dest_folder)
        if "TC(C0)" in dest_folder:
            print("TC(C0)")
            letter_title, letter_vision, letter_num, letter_date = read_tc_excel.read_tc_excel(dest_folder)
        elif "HT" in dest_folder:
            converted_xlsx_path = convert_excel.convert_xlsb(dest_folder)
            letter_title, letter_vision, letter_num, letter_date = read_ht_excel.read_ctc_ht_excel(dest_folder)
        else:
            return None

        text_gen(letter_title, letter_vision, letter_num, letter_date)
        input("Press enter to exit...")

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
