'''
執行pandas的excel
'''
import os
import shutil
import sys
from file_process import move_document, files_list, check_plan, del_dest
from read_ht import convert_xlsb
from read_ht_pandas import read_pandas, read_pandas_sec, read_csv, combine_csv
from read_tc import read_tc_excel

'''
分析EXCEL檔案，【批次】處理
'''
def test_map(excel_path):
    """測試映射值"""
    field_mapping = {
        'No': ['批次序號'],
        'drawing_no_value': ['圖號:', 'CLIENTDOCNO', 'DOCVERSIONDESC'],
        'drawing_title_value': ['圖名:', 'DESCRIPTION', 'DOCCLASS'],
        'drawing_vision_value': ['版次:', 'DOCVERSIONDESC'],
        'letter_num_value': ['來文號碼:', 'TRANSMITTALNO'],
        'letter_date_value': ['來文日期:', 'REVDATE', 'PLANNEDCLIENTRETURNDATE'],
        'letter_titl_value': ['來文名稱:', 'DESCRIPTION'],
        'file_path': ['路徑']
    }
    # 輸入value，輸出key

def work_flow_pandas():
    '''
    讀取【一個】檔案
    1.設定目錄路徑  2.複製檔案  3.解析檔案(HT TC皆完成)
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    source_folder = os.path.join(root_path, "00source")
    dest_folder = os.path.join(root_path, "00dest")

    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.分析plan檔案，興達返回1 台中返回2
    # 【目前只有HT檔案測試，尚無TC測試】
    if str(check_plan(dest_folder)) == "1":
        # 讀取興達
        # 3.1 產生xlsb列表，並轉換成xlsx
        xlsb_file_list = files_list(dest_folder, ".xlsb")
        if xlsb_file_list:
            for xlsb_path in xlsb_file_list:
                convert_xlsb(xlsb_path)

        # 3.2 產生converted.xlsx列表，開始分析內容
        xlsx_file_list = files_list(dest_folder, "converted.xlsx")
        if xlsx_file_list:
            for converted_path in xlsx_file_list:
                read_pandas_sec(converted_path)

    elif str(check_plan(dest_folder)) == "2":
        pass
        # 讀取台中  
    else:
        print("沒有符合檔案")
        sys.exit()

def test_combine_csv():
    """
    測試整合csv功能函數 2023/8/28完成
    目前combine_csv(file1,rawdata)功能為【一對一】
    """
    # 設定目錄
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 複製標準檔案
    file1 = os.path.join(root_path, "01Class/HT-D1-CTC-GEL-23-2710_csv.csv")
    file2 = os.path.join(root_path, "01Class/HT-D1-CTC-GEL-23-2867_csv.csv")
    file_list = [file1, file2]
    rawdata = os.path.join(root_path, "01Class/rawdata.csv")

    # 刪除rawdata
    if os.path.exists(rawdata):
        os.remove(rawdata)
    # 重新複製rawdata
    shutil.copy(file2, rawdata)
    print("已經刪除初始檔案rawdata，並複製新的rawdata")
    # 顯示合併前資料
    read_csv(rawdata)
    # 合併檔案
    combine_csv(file1,rawdata)
    # 顯示合併後資料
    read_csv(rawdata)

if __name__ == '__main__':
    '''測試輸出dataframe'''
    #刪除00dest
    del_dest()
    work_flow_pandas()
    
    '''測試整合csv檔案'''
    #test_combine_csv()
