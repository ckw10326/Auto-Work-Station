'''
已完成的測試
    test_csv_to_df

測試
read_ht_pandas
read_tc_pandas
執行pandas的excel
1. work_flow_pandas，分析00source內的檔案，並解析出csv檔案
2. test_combine_csv，整合單一csv檔案
3. test_combine_csv_list，整合列表內合併到dst
4. test_read_csv，顯示csv檔案
'''
# pylint: disable=W0613
import os, sys, shutil, pandas
from IPython.display import display
import pandas as pd
import read_ht_pandas as RHP
import read_tc_pandas as RTP
from function_file_process import move_document, files_list, check_plan, del_folder, make_folder
from read_ht import convert_xlsb
import function_table_process
from read_ht_pandas import r_ht_ctc_xlsx_sheet_tocsv, xlsx_to_csv, xlsb_to_csv,  clean_csv_letter_cover, read_all_xlsb_df
from read_tc_pandas import xlsm_to_csv, all_xlsm_to_csv




def test_RHP_r_ctc_xlsb_sheet_tocsv():
    '''
    測試功能 確認
    1.路徑
    2.讀取檔案產生表格
    3.csv輸出
    4.csv讀取
    5.復原
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
    if os.path.exists(dest_path):
        print("dest_folder檔案存在\n")

    # 2.讀取檔案產生表格
    RH.convert_xlsb(dest_folder)

    # # 4. 開始轉換
    # # 讀取Data1分頁
    # judgement = 1
    # if judgement:
    #     xlsb_to_csv(dest_path)
    # else:
    #     xlsb_to_csv(dest_path, "Data1")
    # return None

    # 5.復原資料夾
    #del_folder("00temp")

def test_combine_csv():
    """
    測試整合csv功能函數 2023/8/28完成
    目前combine_csv(file1,rawdata)功能為【一對一】
    """
    # 設定目錄
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 複製標準檔案
    file1 = os.path.join(
        root_path, r"08Reference_Files/CSV/HT-D1-CTC-GEL-23-2710_csv.csv")
    file2 = os.path.join(
        root_path, r"08Reference_Files/CSV/HT-D1-CTC-GEL-23-2979_csv.csv")
    rawdata = os.path.join(root_path, "00dest/rawdata.csv")

    # 刪除rawdata
    if os.path.exists(rawdata):
        os.remove(rawdata)
    # 重新複製rawdata
    shutil.copy(file2, rawdata)
    print("已經刪除初始檔案rawdata，並複製新的rawdata")
    # 顯示合併前資料
    csv_to_df(rawdata)
    # 合併檔案
    combine_csv(file1, rawdata)
    # 顯示合併後資料
    csv_to_df(rawdata)

def test_combine_csv_list():
    """
    2023/8/29完成
    combine_csv_list(scr, dst)
    整合檔案列表所有csv檔案
    scr需要為完整路徑的列表
    dst需要為完整csv檔案
    結果資料夾請看00dest
    """
    # 設定目錄
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 複製標準檔案
    file2 = os.path.join(
        root_path, r"08Reference_Files/CSV/HT-D1-CTC-GEL-23-2979_csv.csv")
    dst_data = os.path.join(root_path, "00dest/dst_data.csv")
    # 設定CSV檔案路徑
    csv_path = os.path.join(root_path, "08Reference_Files/CSV")
    if os.path.exists(csv_path):
        # 刪除00dset資料夾
        del_folder("00dest")
        # 創造00dest資料夾
        make_folder("00dest")
        # 刪除dst_data
        if os.path.exists(dst_data):
            os.remove(dst_data)
        # 重新複製dst_data
        shutil.copy(file2, dst_data)
        print("已經刪除初始檔案dst_data，並複製新的dst_data")
        # 顯示合併前資料
        csv_to_df(dst_data)
        print("----------------------")
        # 合併檔案
        csv_lists = files_list(csv_path, ".csv")
        combine_csv_list(csv_lists, dst_data)
        # 顯示合併後資料
        csv_to_df(dst_data)
        print("----------------------")

    else:
        return False

  

def test_txt_to_df():
    """
    2023/08/30 完成測試
    輸入txt檔案路徑
    輸出dataframe
    """
    txt_path = "08Reference_Files/data.txt"
    data_frame, ss = txt_to_df(txt_path)
    print(data_frame)

def test_xlsx_to_csv():
    """
    測試
    1.做作測試檔案
    2.抓取xlsb路徑
    """
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    dest_folder = os.path.join(root_path, "00dest")
    xlsb_path = os.path.join(
        root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2867.xlsb")
    dest_path = os.path.join(
        root_path, ("00dest/" + os.path.basename(xlsb_path)))
    # 複製檔案
    del_folder("00dest")
    make_folder("00dest")
    shutil.copy(xlsb_path, dest_path)

    # 2.抓取xlsb路徑
    xlsb_list = files_list(dest_folder, ".xlsb")
    for xlsbs in xlsb_list:
        if "-CTC-" in xlsbs:
            xlsx_path = convert_xlsb(xlsbs)
            csv_path = xlsx_to_csv(xlsx_path)
            # 讀取後，xlsx便沒有使用，故刪除
            os.remove(xlsx_path)
            print(csv_path)

def test_xlsb_to_csv():
    '''
    轉換xlsb to csv
    自動複製樣本到00dest
    轉換檔案，並存在轉換檔案的位置所在
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2867.xlsb")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00dest")
    dest_path = os.path.join(dest_folder, file_name)

    # 2. 刪除並創建新資料夾
    del_folder("00dest")
    make_folder("00dest")

    # 3. 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 4. 開始轉換
    # 讀取Data1分頁
    judgement = 1
    if judgement:
        xlsb_to_csv(dest_path)
    else:
        xlsb_to_csv(dest_path, "Data1")
    return None

def test_read_all_xlsb_df():
    """測試test_read_all_xlsb_df"""
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定範本檔案路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2867.xlsb")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00dest")
    dest_path = os.path.join(dest_folder, file_name)
    csv_result_path = os.path.join(root_path, "00dest/result.csv")

    # 2. 刪除並創建新資料夾
    del_folder("00dest")
    make_folder("00dest")

    # 3. 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 4.讀取檔案分頁
    df1 = read_all_xlsb_df(dest_path)
    print("讀取合併後xlsb資料為:\n", df1)
    df1.to_csv(csv_result_path)
    return df1

def test_xlsm_to_csv():
    "輸入xlsm路徑，輸出csv檔案"
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/TC/TPC-TC(C0)-CD-23-2078/TPC-TC(C0)-CD-23-2078.xlsm")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00dest")
    dest_path = os.path.join(dest_folder, file_name)

    # 2. 刪除並創建新資料夾
    del_folder("00dest")
    make_folder("00dest")

    # 3. 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 4. 開始轉換
    xlsm_to_csv(dest_path)

def test_clean_df():
    "輸入xlsm路徑，輸出csv檔案"
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/TC/TPC-TC(C0)-CD-23-2078/TPC-TC(C0)-CD-23-2078.xlsm")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00dest")
    dest_path = os.path.join(dest_folder, file_name)

    # 2. 刪除並創建新資料夾
    del_folder("00dest")
    make_folder("00dest")

    # 3. 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 4. 開始轉換
    raw_df = all_xlsm_to_csv(dest_path)
    new_df = clean_df(raw_df, threshold=5)
    print(new_df)

def test_all_xlsm_to_csv():
    "輸入xlsm路徑，輸出csv檔案"
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/TC/TPC-TC(C0)-CD-23-2078/TPC-TC(C0)-CD-23-2078.xlsm")
    file_name = os.path.basename(source_path)
    dest_folder = os.path.join(root_path, "00dest")
    dest_path = os.path.join(dest_folder, file_name)

    # 2. 刪除並創建新資料夾
    del_folder("00dest")
    make_folder("00dest")

    # 3. 複製指定檔案
    shutil.copy2(source_path, dest_path)

    # 4. 開始轉換
    all_xlsm_to_csv(dest_path)

def test_clean_csv_letter_cover():
    """測試test_clean_csv_letter_cover"""
    del_folder("00dest")
    make_folder("00dest")
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_path = os.path.join(
        root_path, "08Reference_Files/HTindex.csv")
    basename = os.path.basename(source_path)  # 結果應為HTindex.csv
    dest_folder = os.path.join(root_path, "00dest")
    dest_path = os.path.join(dest_folder, basename)
    shutil.copyfile(source_path, dest_path)
    # 讀取csv資料
    csv_data_frame = pd.read_csv(dest_path)
    # 篩選csv資料
    csv_data_frame = clean_csv_letter_cover(csv_data_frame, null_num=2)

    '''
    大規模 TC檔案分析
    使用處理方法 xlsm > csv
    1.設定目錄路徑 
    2.複製檔案
    3.解析路徑
    4.【將分頁產生csv檔案】
    '''
    def step_first():
        """複製檔案"""
        source_folder_str="08Reference_Files/collectTC20230728"
        dest_folder_str="00dest"
        # 清空資料夾
        del_folder("00dest")
        make_folder("00dest")

        # 1.設定目錄路徑
        root_path = os.path.dirname(os.path.abspath(__file__))
        # 將根目錄路徑添加到 sys.path
        sys.path.append(root_path)
        # 指定資料夾路徑
        source_folder = os.path.join(root_path, source_folder_str)
        dest_folder = os.path.join(root_path, dest_folder_str)
        # 2.複製檔案
        move_document(source_folder, dest_folder)

    def step_sec():
        """每個資料夾產生csv檔案"""
        # 3.1 解析路徑
        # 3.2 將xlsb檔案轉換為xlsx
        # 3.3 將xlsx轉換換為csv
        xlsm_list = files_list(dest_folder_str, ".xlsm")
        for xlsm in xlsm_list:
            if "TPC-TC(C0)" in xlsm:
                csv_path = all_xlsm_to_csv(xlsm)
                if csv_path:
                    print(csv_path)

    def step_third():
        """設定csv_list，一件自動合併"""
        # 設定csv_list，一件自動合併
        # 1. 設定跟目錄路徑
        root = os.path.dirname(os.path.abspath(__file__))
        # 抓取清單
        m_csv_path = os.path.join(root, "00dest/collectTC20230728")
        csv_list = files_list(m_csv_path, ".csv")
        # 複製標準樣本
        sample_csv_path = os.path.join(
            root, "08Reference_Files/TC_standard_csv.csv")
        test_csv_path = os.path.join(root, "00dest/stander_csv.csv")
        if not os.path.exists(test_csv_path):
            shutil.copy2(sample_csv_path, test_csv_path)
        # 合併清單中檔案
        combine_csv_list(csv_list, test_csv_path)
        # 查看清單檔案
        csv_to_df(test_csv_path)

    step_first()
    step_sec()
    step_third()

if __name__ == '__main__':
    #test_csv_to_df()
    test_RHP_r_ctc_xlsb_sheet_tocsv()