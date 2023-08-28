'''
執行pandas的excel
'''
import os
import shutil
import sys
from file_process import move_document, files_list, check_plan, del_dest
from read_ht import convert_xlsb
from read_ht_pandas import read_pandas, read_csv, combine_csv
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

def main():
    '''
    1.設定目錄路徑  2.複製檔案  3.解析檔案pandas
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    source_folder = os.path.join(root_path, "00source")
    destination_dir = os.path.join(root_path, "00dest")
    # 檢查source_folder是否存在
    excel_path = "/workspaces/Auto-Work-Station/00dest/00source/HT-D1-CTC-GEL-23-2710_converted.xlsx"
    old_csv_path = "/workspaces/Auto-Work-Station/01Class/data.csv"
    read_ctc_ht_excel(excel_path, old_csv_path)

    def test_movedoc():
        '''複製完整結構'''
        path1 = "/workspaces/Auto-Work-Station/09Past"
        path2 = "/workspaces/Auto-Work-Station/00dest"
        move_document(path1, path2)

def work_flow_pandas():
    '''
    讀取【一個】檔案
    1.設定目錄路徑  2.複製檔案  3.解析檔案(HT TC皆完成)
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    dest_folder = os.path.join(root_path, "00dest")
    if not os.path.exists(dest_folder):
        os.mkdir(dest_folder)
    source_test_ht_file0 = os.path.join(root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2710.xlsb")
    source_test_ht_file = os.path.join(root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2867.xlsb")
    # 測試檔案，並測試是否存在
    if not os.path.exists(source_test_ht_file):  # 注意這裡的小寫 "n" in "not"
        print("source_test_ht_file不存在，故結束")
        sys.exit()

    # 2.複製檔案
    dest_ht_file = os.path.join(dest_folder, os.path.split(source_test_ht_file)[1])
    print(source_test_ht_file)
    print(dest_ht_file)
    shutil.copy(source_test_ht_file, dest_ht_file)

    # 3.分析plan檔案，興達返回1 台中返回2
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
                read_pandas(converted_path)

    elif str(check_plan(dest_folder)) == "2":
        pass
        # 讀取台中  
    else:
        print("沒有符合檔案")
        sys.exit()

def test_combine_csv():
    file1 = "/workspaces/Auto-Work-Station/01Class/HT-D1-CTC-GEL-23-2710.csv"
    file2 = "/workspaces/Auto-Work-Station/01Class/HT-D1-CTC-GEL-23-2867.csv"
    file_list = [file1, file2]
    rawdata = "/workspaces/Auto-Work-Station/01Class/rawdata.csv"
    os.remove(rawdata)
    # 1.刪除rawdata
    input("1刪除")
    # 2.重新複製rawdata
    shutil.copy(file2, rawdata)

    read_csv(rawdata)
    combine_csv(file1,rawdata)
    read_csv(rawdata)

if __name__ == '__main__':
    # 刪除00dest
    del_dest()
    if 0:
        # 刪除data.csv
        file_path = "/workspaces/Auto-Work-Station/01Class/data.csv"
        if os.path.exists(file_path):
            os.remove(file_path)
    work_flow_pandas()
    #test_combine_csv()

