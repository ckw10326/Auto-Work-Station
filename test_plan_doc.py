'''
測試「讀取表格指定位置」功能
2023/10/3 
read_ht
    test_convert_xlsb
    test_rh_r_ctc_xlsb_sheet done
read_tc
    test_r_tc_xlsm_sheet

    
cloud_total_run，雲端分析
'''
import sys, os, shutil
import pandas
from function_file_process import files_list, move_document, make_folder, del_folder
import read_ht as rh
import read_tc as rt

def test_convert_xlsb():
    '''
    測試功能 確認
    1.路徑
    2.轉換xlsb成xlsx
    3.復原
    '''

    # 設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2867.xlsb")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00temp")
    dest_path = os.path.join(dest_folder, file_name)

    # 1.路徑測試
    print("1.測試路徑:\n"
          "source_path", source_path, "\n"
          "file_name=", file_name, "\n"
          "dest_folder=", dest_folder, '\
            n'
          "dest_path=" , dest_path, "\n")
    # 刪除並創建新資料夾
    del_folder("00temp")
    make_folder("00temp")
    # 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 2. 開始轉換
    print("2. 開始轉換:")
    rh.convert_xlsb(dest_path)
    input("\n\nplz enter any key")

    # 3.復原
    print("3.復原\n")
    del_folder("00temp")

def test_rh_r_ctc_xlsb_sheet():
    '''
    測試功能 確認
    1.路徑
    2.轉換xlsb成xlsx
    3.讀取產出letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value
    4.復原
    '''
    # 設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2867.xlsb")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00temp")
    dest_path = os.path.join(dest_folder, file_name)

    # 1.路徑測試
    print("1.測試路徑:\n"
          "source_path", source_path, "\n"
          "file_name=", file_name, "\n"
          "dest_folder=", dest_folder, '\n'
          "dest_path=" , dest_path, "\n")
    # 刪除並創建新資料夾
    del_folder("00temp")
    make_folder("00temp")
    # 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 2.轉換xlsb成xlsx
    print("2.轉換xlsb成xlsx")
    xlsx_path = rh.convert_xlsb(dest_path)
    if dest_path:
        print("xlsx_path=", xlsx_path, "\n")

    # 3.讀取產出letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value
    print("3.讀取產出letter_titl_value, drawing_vision_value,\n letter_num_value, letter_date_value\n")
    text_list =rh.r_ctc_xlsx_sheet(xlsx_path)
    for elements in text_list:
        print(elements)
    input("plz enter any key")

    # 4.復原資料夾
    print("4.復原資料夾")
    del_folder("00temp")

def test_rt_r_tc_xlsm_sheet():
    '''
    測試功能 確認
    1.路徑
    2.讀取xlsm產出letter_titl_value, drawing_vision_value, 
    letter_num_value, letter_date_value
    3.復原資料夾
    '''
    # 設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/TC/TPC-TC(C0)-CD-23-2078/TPC-TC(C0)-CD-23-2078.xlsm")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00temp")
    dest_path = os.path.join(dest_folder, file_name)

    # 1.路徑測試
    print("1.測試路徑:\n"
          "source_path", source_path, "\n"
          "file_name=", file_name, "\n"
          "dest_folder=", dest_folder, '\n'
          "dest_path=" , dest_path, "\n")
    # 刪除並創建新資料夾
    del_folder("00temp")
    make_folder("00temp")
    # 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 2.讀取xlsm產出letter_titl_value, drawing_vision_value, \
    #   letter_num_value, letter_date_value
    print("2.讀取xlsm產出letter_titl_value, drawing_vision_value")
    if dest_path:
        text_list =rt.r_tc_xlsm_sheet(dest_path)
        for elements in text_list:
            print(elements)
        input("plz enter any key")

    # 4.復原資料夾
    print("3.復原資料夾")
    del_folder("00temp")


if __name__ == '__main__':
    # 已完成測試
    #test_convert_xlsb()
    #test_rh_r_ctc_xlsb_sheet()
    #test_rt_r_tc_xlsm_sheet()
