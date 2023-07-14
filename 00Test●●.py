'''
測試子程式
測試整合流程
'''
import sys
from file_list import files_list1
from read_ht_excel_cloud0712 import convert_xlsb
from read_ht_excel_cloud0712 import read_ctc_ht_excel
from read_tc_excel_cloud0712 import read_tc_excel
from doc_collection import file_path_process
from doc_collection import move_docutment
from text_gen import text_gen

#測試file_path_process、move_docutment
#輸入基本訊息，並顯示各項目路徑
def test_folder_path():
    file_doc_num = ""
    file_source_folder = "/workspaces/Auto-Work-Station/00source"
    file_dest_folder = "/workspaces/Auto-Work-Station/00source"
    x = [file_doc_num, file_source_folder, file_dest_folder] = file_path_process()
    for asss in x:
        print(asss)
    #move_docutment(file_source_folder, file_dest_folder)

#測試雲端解析興達文件
def test_cloud_ht():

    path = r"/workspaces/Auto-Work-Station/00source"
    #列表，檔案清單
    the_xlsb_file_list = files_list1(path, ".xlsb")

    #列表，有轉換檔案後的清單
    for filepath in the_xlsb_file_list:
        convert_xlsb(filepath)
    print("convered_xlsb Done\n")

    #列表，輸出成converted.xlsx清單
    the_xlsx_file_list = files_list1(path, "converted.xlsx")
    for the_file in the_xlsx_file_list:
        read_ctc_ht_excel(the_file)

#測試雲端解析台中文件
def test_cloud_tc():
    path = r"/workspaces/Auto-Work-Station/00source"
    #列表，檔案清單
    the_xlsb_file_list = files_list1(path, ".xlsm")

    #讀取列表中的清單
    for the_file in the_xlsb_file_list:
        read_tc_excel(the_file)

#測試路徑是否存在
def test_path():
    file_source_folder = "/workspaces/Auto-Work-Station/00source"
    file_dest_folder = "/workspaces/Auto-Work-Station/00source"
    print(file_source_folder)
    print(file_dest_folder)
    if os.path.exists(file_source_folder):
        print("路徑file_source_folder確實存在")
    else:
        print("不存在")

#測試生成套印文字
def test_text_gen():
    dest_folder = r"/workspaces/Auto-Work-Station/00source"
    text_gen('HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab',
             'A', 'TPC-TC(C0)-CD-23-0002', '2023/02/02', dest_folder)
    input("Press enter to exit...")
    return None

#總流程測試
def cloud_total_run():
    ROOT_GIT = r"/workspaces/Auto-Work-Station"
    sys.path.append(ROOT_GIT)

    path = r"/workspaces/Auto-Work-Station/00source"
    #解析.xlsb(興達資料)，用來判斷是否有興達.xlsb檔案
    judge_ht = 0
    the_xlsb_file_list = files_list1(path, ".xlsb")
    for filepath in the_xlsb_file_list:
        #x為判斷是否有轉檔
        x = convert_xlsb(filepath)
    print("convered_xlsb Done\n")

    if x :
        letter_title, letter_vision, letter_num, letter_date = read_ctc_ht_excel(path)
        text_gen(letter_title, letter_vision, letter_num, letter_date, path)
    else :
        print("convert_xlsb回傳None，無符合興達分析之檔案，現在開始分析「台中」")
        letter_title, letter_vision, letter_num, letter_date = read_tc_excel(path)
        text_gen(letter_title, letter_vision, letter_num, letter_date, path)

if __name__ == '__main__':
    #test_cloud_ht()
    #test_cloud_tc()
    #test_path()
    #test_folder_path()
    #test_text_gen()
    cloud_total_run()
