# pylint: disable=W0613
'''
執行pandas的excel
1. work_flow_pandas，分析00source內的檔案，並解析出csv檔案
2. test_combine_csv，整合單一csv檔案
3. test_combine_csv_list，整合列表內合併到dst
4. test_read_csv，顯示csv檔案
'''
import os
import shutil
import sys
import pandas as pd
from file_process import move_document, files_list, check_plan, del_folder, make_folder
from read_ht import convert_xlsb
from table_process import combine_csv, combine_csv_list, read_csv, txt_to_df, add_column_dataframe
from read_ht_pandas import read_pandas_sec, xlsx_to_csv, xlsb_to_csv
from read_tc_pandas import xlsm_to_csv

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


def generateCsvFilesBatch(source_folder_str="08Reference_Files/collectHT20230728",
                          dest_folder_str="00dest"):
    '''
    使用read_pandas_sec，該功能為暴力分析Excel(抓取指定位置)
    缺點是只是用新檔案，舊檔案會有index格不相容的問題
    分析【指定】資料夾下所有Excel檔案
    1.設定目錄路徑 
    2.複製檔案
    3.解析路徑
    4.產生csv檔案
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_folder = os.path.join(root_path, source_folder_str)
    dest_folder = os.path.join(root_path, dest_folder_str)

    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.1 解析路徑，產出xlsb清單
    # 3.2 將xlsb檔案轉換為xlsx
    # 3.3 將xlsx轉換
    xlsb_list = files_list(dest_folder, ".xlsb")
    for xlsbs in xlsb_list:
        if "-CTC-" in xlsbs:
            xlsx_path = convert_xlsb(xlsbs)
            csv_path = read_pandas_sec(xlsx_path)
            # 讀取後，xlsx便沒有使用，故刪除
            os.remove(xlsx_path)
            print(csv_path)


def generateCsvFilesBatch2(source_folder_str="08Reference_Files/collectHT20230728",
                           dest_folder_str="00dest"):
    '''
    改善新舊檔案不相容，直接將檔案分頁 xlsb > xlsx > csv
    分析【指定】資料夾下所有Excel檔案
    1.設定目錄路徑 
    2.複製檔案
    3.解析路徑
    4.【將分頁產生csv檔案
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_folder = os.path.join(root_path, source_folder_str)
    dest_folder = os.path.join(root_path, dest_folder_str)

    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.1 解析路徑
    # 3.2 將xlsb檔案轉換為xlsx
    # 3.3 將xlsx轉換換為csv
    xlsb_list = files_list(dest_folder, ".xlsb")
    for xlsbs in xlsb_list:
        if "-CTC-" in xlsbs:
            xlsx_path = convert_xlsb(xlsbs)
            csv_path = xlsx_to_csv(xlsx_path)
            # 讀取後，xlsx便沒有使用，故刪除
            os.remove(xlsx_path)
            print(csv_path)


def generateCsvFilesBatch3(source_folder_str="08Reference_Files/collectHT20230728",
                           dest_folder_str="00dest"):
    '''
    優化分析路徑，
    舊處理方法 xlsb > xlsx > csv
    新處理方法 xlsb > csv
    測試使用xlsb_to_csv
    1.設定目錄路徑 
    2.複製檔案
    3.解析路徑
    4.【將分頁產生csv檔案】
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_folder = os.path.join(root_path, source_folder_str)
    dest_folder = os.path.join(root_path, dest_folder_str)

    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.1 解析路徑
    # 3.2 將xlsb檔案轉換為xlsx
    # 3.3 將xlsx轉換換為csv
    xlsb_list = files_list(dest_folder, ".xlsb")
    for xlsbs in xlsb_list:
        if "-CTC-" in xlsbs:
            csv_path = xlsb_to_csv(xlsbs)
            print(csv_path)


def generateCsvFilesBatch4(source_folder_str="08Reference_Files/collectHT20230728",
                           dest_folder_str="00dest"):
    '''
    大規模 檔案分析開
    使用處理方法 xlsb > csv
    1.設定目錄路徑 
    2.複製檔案
    3.解析路徑
    4.【將分頁產生csv檔案】
    '''
    # 0.清除不需要資料夾
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
    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.1 解析路徑
    # 3.2 將xlsb檔案轉換為xlsx
    # 3.3 將xlsx轉換換為csv
    xlsb_list = files_list(dest_folder, ".xlsb")
    for xlsbs in xlsb_list:
        if "-CTC-" in xlsbs:
            csv_path = xlsb_to_csv(xlsbs)
            print(csv_path)

def batch_gen_tc_csv(source_folder_str="08Reference_Files/collectTC20230728",
                           dest_folder_str="00dest"):
    '''
    大規模 TC檔案分析
    使用處理方法 xlsm > csv
    1.設定目錄路徑 
    2.複製檔案
    3.解析路徑
    4.【將分頁產生csv檔案】
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    source_folder = os.path.join(root_path, source_folder_str)
    dest_folder = os.path.join(root_path, dest_folder_str)
    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.1 解析路徑
    # 3.2 將xlsb檔案轉換為xlsx
    # 3.3 將xlsx轉換換為csv
    xlsm_list = files_list(dest_folder_str, ".xlsm")
    for xlsm in xlsm_list:
        if "TPC-TC(C0)" in xlsm:
            csv_path = xlsm_to_csv(xlsm)
            if csv_path:
                print(csv_path)

def show_tc_csv():
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    csv_folder = os.path.join(root_path, "00dest")
    print(csv_folder)
    tc_csv_list = files_list(csv_folder, ".csv")
    print(tc_csv_list)
    for csv in tc_csv_list:
        print(csv)

def make_stander_csv():
    """表列表頭清單"""
    lista = ["TRANSMITTALNO", "CTCIDOCNO", "VENDORDOCNO", "CLIENTDOCNO", "DOCVERSIONDESC", "DOCREV", "REVDATE",
             "ISSUEPURPOSE", "IFDPLAN", "ISSUEDATE", "RETUREDATE", "TAGNO", "SUBMITENGINEER", "DESCRIPTION"]
    listb = ["TRANSMITTALNO", "ISSUEPURPOSE", "PLANNEDCLIENTRETURNDATE", "ISSUEDATE", "DOCVERSIONNO",
             "DOCVERSIONDESC", "DOCREV", "DOCNATURE", "DOCCLASS", "TAGNO", "CLIENTDOCNO", "DESCRIPTION"]
    union_list = list(set(lista).union(set(listb)))
    print(union_list)

    # 建立一個空的 DataFrame
    data_frame = pd.DataFrame(columns=union_list)
    # 指定輸出的 CSV 檔案名稱
    root_path = os.path.dirname(os.path.abspath(__file__))
    stander_csv_path = os.path.join(root_path, r"00dest/stander_csv.csv")
    if os.path.exists(stander_csv_path):
        os.remove(stander_csv_path)
    print(stander_csv_path)
    data_frame.to_csv(stander_csv_path, index=False)  # 將 DataFrame 寫入 CSV 檔案
    print("CSV 檔案已成功輸出。")

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
    read_csv(rawdata)
    # 合併檔案
    combine_csv(file1, rawdata)
    # 顯示合併後資料
    read_csv(rawdata)


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
        read_csv(dst_data)
        print("----------------------")
        # 合併檔案
        csv_lists = files_list(csv_path, ".csv")
        combine_csv_list(csv_lists, dst_data)
        # 顯示合併後資料
        read_csv(dst_data)
        print("----------------------")

    else:
        return False


def test_read_csv():
    """讀取csv檔案"""
    # 設定目錄
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 複製標準檔案
    filepath = os.path.join(
        root_path, r"08Reference_Files/CSV/HT-D1-CTC-GEL-23-3046_csv.csv")
    read_csv(filepath)


def test_txt_to_df():
    """
    2023/08/30 完成測試
    輸入txt檔案路徑
    輸出dataframe
    """
    data_frame, ss = txt_to_df()
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
    xlsb_to_csv(dest_path)

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


if __name__ == '__main__':
    # 代號1，辦文專用，使用work_flow_pandas()抓取指定位置，暴力法辦文
    #       分析檔案後，產生1表格，放置到00source資料夾，並自動清空
    # 代號21，批次處理，廠商xlsb > xlsx > 抓取表格指定位置(read_pandas_sec) > csv，簡稱暴力抓取，批次產生csv檔案，index較好看
    # 代號22，批次處理，廠商xlsb > xlsx > 將分頁轉乘csv(xlsx_to_csv)，批次產生csv檔案，index為原生
    # 代號23，主要使用，批次處理，廠商xlsb > csv，批次產生csv檔案，index為原生
    # 代號24，家中大規模測試，批次處理，廠商xlsb > csv，批次產生csv檔案，index為原生
    # 代號3，務必先執行後generateCsvFilesBatch()後，才有可合併csv的檔案
    # 代號4，【新功能】原生excel分頁直接轉換成csv，開發中
    ht_judge = 1054
    if ht_judge == 1:
        del_folder("00dest")
        work_flow_pandas()
        del_folder("00source")
        make_folder("00source")

    if ht_judge == 21:
        del_folder("00dest")
        generateCsvFilesBatch()

    if ht_judge == 22:
        del_folder("00dest")
        generateCsvFilesBatch2()

    if ht_judge == 23:
        del_folder("00dest")
        generateCsvFilesBatch3()

    if ht_judge == 24:
        source = r"C:\Users\S\Downloads\collectTC20230728"
        dest = r"C:\Users\S\Downloads\00dest"
        generateCsvFilesBatch4(source, dest)
    
    if ht_judge == 3:
        # stander_csv為廠商原生index
        # 設定csv_list，一件自動合併
        # 1. 設定跟目錄路徑
        root = os.path.dirname(os.path.abspath(__file__))
        # 抓取清單
        m_csv_path = os.path.join(root, "00dest/collectHT20230728")
        csv_list = files_list(m_csv_path, ".csv")
        # 複製標準樣本
        sample_csv_path = os.path.join(
            root, "08Reference_Files/stander_csv.csv")
        test_csv_path = os.path.join(root, "00dest/stander_csv.csv")
        if not os.path.exists(test_csv_path):
            shutil.copy2(sample_csv_path, test_csv_path)
        # 合併清單中檔案
        combine_csv_list(csv_list, test_csv_path)
        # 查看清單檔案
        read_csv(test_csv_path)

    # make_stander_csv()
    # test_xlsx_to_csv()
    # test_txt_to_df()
    # test_read_csv()
    # test_combine_csv()
    # test_combine_csv_list()
    # test_xlsb_to_csv()

    
    # show_tc_csv()

    # 代號50，批次產生TC的csv檔案
    # 代號51，批次合併標準檔案
    tc_judge = 52
    if tc_judge == 50:
        # 測試轉換csv檔案，須加上path路徑
        test_xlsm_to_csv()

    if tc_judge == 51:
        batch_gen_tc_csv()

    if tc_judge == 52:
        # stander_csv為廠商原生index
        # 設定csv_list，一件自動合併
        # 1. 設定跟目錄路徑
        root = os.path.dirname(os.path.abspath(__file__))
        # 抓取清單
        m_csv_path = os.path.join(root, "00dest/collectTC20230728")
        csv_list = files_list(m_csv_path, ".csv")
        # 複製標準樣本
        sample_csv_path = os.path.join(
            root, "08Reference_Files/stander_csv.csv")
        test_csv_path = os.path.join(root, "00dest/stander_csv.csv")
        if not os.path.exists(test_csv_path):
            shutil.copy2(sample_csv_path, test_csv_path)
        # 合併清單中檔案
        combine_csv_list(csv_list, test_csv_path)
        # 查看清單檔案
        read_csv(test_csv_path)