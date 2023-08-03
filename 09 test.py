'''
測試子程式
測試整合流程
'''
import sys
import os
import shutil
import class_format
from file_process import files_list, move_document
from read_ht_class import read_ctc_ht_excel
from read_tc_excel_cloud0712 import read_tc_excel
from text_gen import file_path_process
from text_gen import text_gen
from text_gen import copy_stand_file

file_doc_num = ""
file_source_folder = "/workspaces/Auto-Work-Station/00source"
file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
test_folder = "/workspaces/Auto-Work-Station/08Reference_Files"


# 測試雲端解析興達文件 Done 2023/07/20
# 1.複製資料夾  2.分析Excel檔案  3.生成套印文字
def test_cloud_ht():
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
                 letter_num_value, letter_date_value, the_file)
        input("enter any keys to exit")

# 測試路徑是否存在 Done 2023/07/20


def test_path():
    file_source_folder = "/workspaces/Auto-Work-Station/00source"
    file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
    paths = [file_source_folder, file_dest_folder]
    for path in paths:
        if os.path.exists(path):
            print(f"路徑 {path} 確實存在")
        else:
            print(f"路徑 {path} 不存在")

# 測試生成套印文字


def test_text_gen():
    dest_folder = r"/workspaces/Auto-Work-Station/00source"
    text_gen('HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab',
             'A', 'TPC-TC(C0)-CD-23-0002', '2023/02/02', dest_folder)
    input("Press enter to exit...")
    return None

def cloud_total_run_class():
    """# 總流程測試"""
    ROOT_GIT = r"/workspaces/Auto-Work-Station"
    sys.path.append(ROOT_GIT)

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
                         letter_num_value, letter_date_value, destination_dir)
                # 複製套印、傳真檔案
                copy_stand_file(the_file)
                input("enter any keys to exit")
        else:
            pass


if __name__ == '__main__':
    file_dest_folder = "/workspaces/Auto-Work-Station/00dest"
    # test_folder_path()

    shutil.rmtree(file_dest_folder)
    cloud_total_run()
