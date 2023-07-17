'''
測試子程式
測試整合流程
'''
import sys
import os
import shutil
from file_list import files_list1
from read_ht_excel_cloud0712 import convert_xlsb
from read_ht_excel_cloud0712 import read_ctc_ht_excel
from read_tc_excel_cloud0712 import read_tc_excel
from doc_collection import file_path_process
from doc_collection import move_docutment
from text_gen import text_gen

file_doc_num = ""
file_source_folder = "/workspaces/Auto-Work-Station/00source"
file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
test_folder = "/workspaces/Auto-Work-Station/08Reference_Files"

#測試file_path_process、move_docutment
#輸入基本訊息，並顯示各項目路徑
#Done
def test_folder_path():
    x = [file_doc_num, file_source_folder, file_dest_folder] = file_path_process()
    for asss in x:
        print(asss)
    #move_docutment(file_source_folder, file_dest_folder)

#測試雲端解析興達文件
#Done
def test_cloud_ht():
    sample_ht_folder = "/workspaces/Auto-Work-Station/08Reference_Files/HT"
    destination_dir = "/workspaces/Auto-Work-Station/00dest/HT_Sample"
    #複製Sample資料夾結構
    shutil.copytree(sample_ht_folder, destination_dir)
    #列表，檔案清單
    the_xlsb_file_list = files_list1(destination_dir, ".xlsb")

    #列表，有轉換檔案後的清單
    print("開始表列符合.xlsb清單")
    for filepath in the_xlsb_file_list:
        print("符合.xlsb清單", filepath)
        convert_xlsb(filepath)
    print("convered_xlsb Done\n")

    #列表，輸出符合條件"converted.xlsx"清單
    the_xlsx_file_list = files_list1(destination_dir, "converted.xlsx")
    print("開始表列符合converted.xlsb清單")
    for the_file in the_xlsx_file_list:
        print("符合converted.xlsb清單:",the_file)
        input("enter any keys to exit")
        read_ctc_ht_excel(the_file)
        input("enter any keys to exit")

#測試雲端解析台中文件
#Done
def test_cloud_tc():
    sample_tc_folder = "/workspaces/Auto-Work-Station/08Reference_Files/TC"
    destination_dir = "/workspaces/Auto-Work-Station/00dest/TC_Sample"
    #複製Sample資料夾結構
    shutil.copytree(sample_tc_folder, destination_dir)
    #列表，檔案清單
    the_xlsm_file_list = files_list1(destination_dir, ".xlsm")

    #列表，有符合條件清單
    print("開始表列符合.xlsm清單")
    for filepath in the_xlsm_file_list:
        print("符合.xlsm清單", filepath)
    input("enter any keys to exit")

    #讀取列表中的清單
    for the_file in the_xlsm_file_list:
        read_tc_excel(the_file)
        input("enter any keys to exit")

#測試路徑是否存在
def test_path():
    file_source_folder = "/workspaces/Auto-Work-Station/00source"
    file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
    print(file_source_folder)
    print(file_dest_folder)
    if os.path.exists(file_source_folder):
        print("路徑file_source_folder確實存在")
    else:
        print("不存在")
    if os.path.exists(file_dest_folder):
        print("路徑file_dest_folder確實存在")
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

    sample_folder = "/workspaces/Auto-Work-Station/00source"
    destination_dir = "/workspaces/Auto-Work-Station/00dest"
    #複製Sample資料夾結構
    shutil.copytree(sample_folder, destination_dir)
    print("複製完成")
    
    #解析.xlsb(興達資料)，用來判斷是否有興達.xlsb檔案
    the_xlsb_file_list = files_list1(destination_dir, ".xlsb")
    for filepath in the_xlsb_file_list:
        print("符合.xlsb清單", filepath)
        convert_xlsb(filepath)
    the_xlsx_file_list = files_list1(destination_dir, "converted.xlsx")
    for the_file in the_xlsx_file_list:
        print("符合converted.xlsb清單:",the_file)
        input("enter any keys to exit")
        read_ctc_ht_excel(the_file)
        input("enter any keys to exit")

    #測試雲端解析台中文件
    the_xlsm_file_list = files_list1(destination_dir, ".xlsm")
    for filepath in the_xlsm_file_list:
        print("符合.xlsm清單", filepath)
        input("enter any keys to exit")
    for the_file in the_xlsm_file_list:
        read_tc_excel(the_file)
        input("enter any keys to exit")

    print("解析完成!")
    print("開始產生文字")
    input("enter any keys to exit")

if __name__ == '__main__':
    file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
    #test_folder_path()

    #shutil.rmtree(file_dest_folder)
    #test_cloud_ht()

    #shutil.rmtree(file_dest_folder)
    #test_cloud_tc()

    #test_path()
    
    #test_text_gen()

    if os.path.exists(file_dest_folder):
        shutil.rmtree(file_dest_folder)
    cloud_total_run()