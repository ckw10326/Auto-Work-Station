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
from function_file_process import move_document, files_list, check_plan, del_folder, make_folder
from read_ht import convert_xlsb
from function_table_process import combine_csv, combine_csv_list, csv_to_df, txt_to_df, clean_df
from read_ht_pandas import r_ht_ctc_xlsx_sheet_tocsv, xlsx_to_csv, xlsb_to_csv,  clean_csv_letter_cover, read_all_xlsb_df
from read_tc_pandas import xlsm_to_csv, all_xlsm_to_csv

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
                r_ht_ctc_xlsx_sheet_tocsv(converted_path)

    elif str(check_plan(dest_folder)) == "2":
        pass
        # 讀取台中
    else:
        print("沒有符合檔案")
        sys.exit()

def show_tc_csv():
    "show_tc_csv"
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

def tc_total_table():
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
    # 代號1，辦文專用，使用work_flow_pandas()抓取指定位置，暴力法辦文
    #       分析檔案後，產生1表格，放置到00source資料夾，並自動清空
    # 代號21，批次處理，廠商xlsb > xlsx > 抓取表格指定位置(read_pandas_sec) > csv，簡稱暴力抓取，批次產生csv檔案，index較好看
    # 代號22，批次處理，廠商xlsb > xlsx > 將分頁轉乘csv(xlsx_to_csv)，批次產生csv檔案，index為原生
    # 代號23，主要使用，批次處理，廠商xlsb > csv，批次產生csv檔案，index為原生
    # 代號24，家中大規模測試，批次處理，廠商xlsb > csv，批次產生csv檔案，index為原生
    # 代號3，務必先執行後generateCsvFilesBatch()後，才有可合併csv的檔案
    # 代號4，【新功能】原生excel分頁直接轉換成csv，開發中

    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    # 指定資料夾路徑
    TC_STANDARD_FILE = os.path.join(
        root_path, "08Reference_Files/TC/TPC-TC(C0)-CD-23-2078/TPC-TC(C0)-CD-23-2078.xlsm")
    dest_folder = os.path.join(root_path, "00dest")


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
        csv_to_df(test_csv_path)

    # make_stander_csv()
    # test_xlsx_to_csv()
    # test_txt_to_df()
    # test_read_csv()
    # test_combine_csv()
    # test_combine_csv_list()
    # test_xlsb_to_csv()
    # test_read_all_xlsb_df()
    # test_clean_csv_letter_cover()
    # test_clean_df()
    # show_tc_csv()
    # test_all_xlsm_to_csv()
    tc_total_table()
    
    # 代號50，批次產生TC的csv檔案
    # 代號51，批次合併標準檔案
    # 代號52，合併csV檔案內容
    tc_judge = 500
    if tc_judge == 50:
        # 測試轉換csv檔案，須加上path路徑
        # test_xlsm_to_csv()
        test_all_xlsm_to_csv()

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
        csv_to_df(test_csv_path)
