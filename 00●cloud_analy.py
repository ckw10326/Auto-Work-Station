'''
測試子程式
測試整合流程
'''
import sys
import os
import shutil
from file_process import files_list, move_document
from read_ht_excel_cloud0712 import convert_xlsb
from read_ht_excel_cloud0712 import read_ctc_ht_excel
from read_tc_excel_cloud0712 import read_tc_excel
from text_gen import copy_plan_file
from text_gen import text_gen

file_doc_num = ""
file_source_folder = "/workspaces/Auto-Work-Station/00source"
file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
test_folder = "/workspaces/Auto-Work-Station/08Reference_Files"

def test_cloud_ht():
    '''
    測試雲端解析興達文件 Done 2023/07/20
    1.複製資料夾  2.分析Excel檔案  3.生成套印文字
    '''
    sample_ht_folder = "/workspaces/Auto-Work-Station/08Reference_Files/HT"
    destination_dir = "/workspaces/Auto-Work-Station/00dest"
    # 複製Sample資料夾結構
    shutil.copytree(sample_ht_folder, destination_dir)
    # 列表，檔案清單
    the_xlsb_file_list = files_list(destination_dir, ".xlsb")

    # 列表，有轉換檔案後的清單
    print("開始表列符合.xlsb清單")
    for filepath in the_xlsb_file_list:
        print("符合.xlsb清單", filepath)
        convert_xlsb(filepath)
    print("convered_xlsb Done\n")

    # 列表，輸出符合條件"converted.xlsx"清單
    the_xlsx_file_list = files_list(destination_dir, "converted.xlsx")
    print("開始表列符合converted.xlsb清單")
    for the_file in the_xlsx_file_list:
        print("符合converted.xlsb清單:", the_file)
        input("enter any keys to exit")
        letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value = read_ctc_ht_excel(
            the_file)
        text_gen(letter_titl_value, drawing_vision_value,
                 letter_num_value, letter_date_value, the_file)
        input("enter any keys to exit")

# 測試雲端解析台中文件 Done 2023/07/20
# 1.複製資料夾  2.分析Excel檔案  3.生成套印文字

def test_cloud_tc():
    '''測試雲端台中檔案'''
    sample_tc_folder = "/workspaces/Auto-Work-Station/08Reference_Files/TC"
    destination_dir = "/workspaces/Auto-Work-Station/00dest"
    # 複製Sample資料夾結構
    shutil.copytree(sample_tc_folder, destination_dir)
    # 列表，檔案清單
    the_xlsm_file_list = files_list(destination_dir, ".xlsm")

    # 列表，有符合條件清單
    print("開始表列符合.xlsm清單")
    for filepath in the_xlsm_file_list:
        print("符合.xlsm清單", filepath)
    input("enter any keys to exit")

    # 讀取列表中的清單
    for the_file in the_xlsm_file_list:
        letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value = read_tc_excel(
            the_file)
        text_gen(letter_titl_value, drawing_vision_value,
                 letter_num_value, letter_date_value)
        input("enter any keys to exit")

def cloud_total_run():
    """總流程測試"""
    root_git = r"/workspaces/Auto-Work-Station"
    sys.path.append(root_git)

    sample_folder = "/workspaces/Auto-Work-Station/00source"
    destination_dir = "/workspaces/Auto-Work-Station/00dest"
    # 複製Sample資料夾結構
    move_document(sample_folder, destination_dir)
    print("複製destination_dir到00dest資料夾完成", destination_dir, "\n")

    listnum = 0
    # 解析.xlsb(興達資料)，判斷是否有興達.xlsb檔案
    if 1:
        the_xlsb_file_list = files_list(destination_dir, ".xlsb")
        listnum = len(the_xlsb_file_list)
        # 若是清單有內容才會帶入解析
        if listnum:
            for filepath in the_xlsb_file_list:
                print("符合興達計畫.xlsb清單", filepath)
                # 若符合.xlsb，則轉檔
                convert_xlsb(filepath)
        else:
            pass
        # 將.xlsb轉換成.xlsx後，開始分析內容
        the_xlsx_file_list = files_list(destination_dir, "converted.xlsx")
        listnum = len(the_xlsx_file_list)
        # 若是清單有內容才會帶入解析
        if listnum:
            for the_file in the_xlsx_file_list:
                print("符合興達計畫converted.xlsb清單:", the_file)
                input("enter any keys to exit")
                letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value = read_ctc_ht_excel(
                    the_file)
                print(letter_titl_value, drawing_vision_value,
                      letter_num_value, letter_date_value)
                print("-----------將要執行text_gen--------------")
                text_gen(letter_titl_value, drawing_vision_value,
                         letter_num_value, letter_date_value)
                # 複製套印、傳真檔案
                copy_plan_file(the_file)
                input("enter any keys to exit")
        else:
            pass

    listnum = 0
    # 解析.xlsx(台中資料)，判斷是否有台中.xlsb檔案
    if 1:
        the_xlsm_file_list = files_list(destination_dir, ".xlsm")
        listnum = len(the_xlsm_file_list)
        if listnum:
            for the_file in the_xlsm_file_list:
                print("符合台中計畫.xlsm清單", the_file)
                input("enter any keys to exit")
                letter_title, letter_vision, letter_num, letter_date_value = read_tc_excel(
                    the_file)
                print(letter_title, letter_vision,
                      letter_num, letter_date_value)
                text_gen(letter_title, letter_vision, letter_num,
                         letter_date_value)
                # 複製套印、傳真檔案
                copy_plan_file(the_file)
                input("enter any keys to exit")
        input("----END----")

if __name__ == '__main__':
    file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
    # test_folder_path()
    if os.path.exists(file_dest_folder):
        shutil.rmtree(file_dest_folder)
    # test_cloud_ht()

    # shutil.rmtree(file_dest_folder)
    # test_cloud_tc()

    # test_path()

    # test_text_gen()

    # shutil.rmtree(file_dest_folder)

    cloud_total_run()

    # 手動需要資料
