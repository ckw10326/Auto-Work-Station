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
from table_process import combine_csv, combine_csv_list, read_csv, txt_to_df
from read_ht_pandas import read_pandas_sec, ht_to_csv

def test_map(excel_path):
    """測試映射值，【未完成】"""
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

def generateCsvFilesBatch(source_folder_str = "08Reference_Files/collectHT20230728", 
                        dest_folder_str = "00dest"):
    '''
    功能已經非常完整，可惜廠商文件，新舊檔案不相容
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
    source_folder = os.path.join(
        root_path, source_folder_str)
    dest_folder = os.path.join(root_path, dest_folder_str)

    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.1 解析路徑
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

def generateCsvFilesBatch2(source_folder_str = "08Reference_Files/collectHT20230728", 
                        dest_folder_str = "00dest"):
    '''
    改善新舊檔案不相容
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
    source_folder = os.path.join(
        root_path, source_folder_str)
    dest_folder = os.path.join(root_path, dest_folder_str)

    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.1 解析路徑
    # 3.2 將xlsb檔案轉換為xlsx
    # 3.3 將xlsx轉換
    xlsb_list = files_list(dest_folder, ".xlsb")
    for xlsbs in xlsb_list:
        if "-CTC-" in xlsbs:
            xlsx_path = convert_xlsb(xlsbs)
            csv_path = ht_to_csv(xlsx_path)
            # 讀取後，xlsx便沒有使用，故刪除
            os.remove(xlsx_path)
            print(csv_path)

def make_stander_csv():
    """表列表頭清單"""
    lista = ["TRANSMITTALNO","CTCIDOCNO","VENDORDOCNO","CLIENTDOCNO","DOCVERSIONDESC","DOCREV","REVDATE","ISSUEPURPOSE","IFDPLAN","ISSUEDATE","RETUREDATE","TAGNO","SUBMITENGINEER","DESCRIPTION"]
    listb = ["TRANSMITTALNO","ISSUEPURPOSE","PLANNEDCLIENTRETURNDATE","ISSUEDATE","DOCVERSIONNO","DOCVERSIONDESC","DOCREV","DOCNATURE","DOCCLASS","TAGNO","CLIENTDOCNO","DESCRIPTION"]
    union_list = list(set(lista).union(set(listb)))
    print(union_list)

    # 建立一個空的 DataFrame
    df = pd.DataFrame(columns=union_list)
    # 指定輸出的 CSV 檔案名稱
    root_path = os.path.dirname(os.path.abspath(__file__))
    stander_csv_path = os.path.join(root_path, r"00dest/stander_csv.csv")
    if os.path.exists(stander_csv_path):
        os.remove(stander_csv_path)
    print(stander_csv_path)
    df.to_csv(stander_csv_path, index=False)  # 將 DataFrame 寫入 CSV 檔案
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
    df, data = txt_to_df()
    print(df)

def test_ht_to_csv():
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
    xlsb_path = os.path.join(root_path, "08Reference_Files/HT/HT-D1-CTC-GEL-23-2867.xlsb")
    dest_path = os.path.join(root_path, ("00dest/" + os.path.basename(xlsb_path)))
    # 複製檔案
    del_folder("00dest")
    make_folder("00dest")
    shutil.copy(xlsb_path, dest_path)
    
    # 2.抓取xlsb路徑
    xlsb_list = files_list(dest_folder, ".xlsb")
    for xlsbs in xlsb_list:
        if "-CTC-" in xlsbs:
            xlsx_path = convert_xlsb(xlsbs)
            csv_path = ht_to_csv(xlsx_path)
            # 讀取後，xlsx便沒有使用，故刪除
            os.remove(xlsx_path)
            print(csv_path)

if __name__ == '__main__':
    # 代號1，分析單Excel檔案，放置到00source資料夾，並自動清空
    # 代號2，批次分析資料夾Excel檔案，批次產生csv檔案
    # 代號3，合併csv檔案
    # 代號4，【新功能】將excel分頁轉換成csv
    xx_judge = 1
    if xx_judge == 1:
        del_folder("00dest")
        work_flow_pandas()
        del_folder("00source")
        make_folder("00source")

    if xx_judge == 2:
        # 批次轉換excel檔案成為csv檔
        del_folder("00dest")
        # generateCsvFilesBatch()，功能在美化廠商檔案轉化為CSV，表頭【標題】等
        #generateCsvFilesBatch()
        # generateCsvFilesBatch2()，功能在於直接將廠商檔案轉化為CSV，不做修飾
        generateCsvFilesBatch2()

    if xx_judge == 31:
        # 抓取清單
        csv_list = files_list("/workspaces/Auto-Work-Station/00dest/collectHT20230728", ".csv")
        # 複製標準樣本
        sample_csv_path = r"00dest/collectHT20230728/HT-D1-CTC-GEL-22-0024/HT-D1-CTC-GEL-22-0024_csv.csv"
        test_csv_path = r"/workspaces/Auto-Work-Station/00dest/raw.csv"
        shutil.copy2(sample_csv_path, test_csv_path)
        # 合併清單中檔案
        combine_csv_list(csv_list, test_csv_path)
        # 查看清單檔案
        read_csv(test_csv_path)

    if xx_judge == 32:
        '''嘗試改良31的問題，把欄位增加'''
        root_path = os.path.dirname(os.path.abspath(__file__))
        # 抓取清單
        csv_path = os.path.join(root_path, "00dest/collectHT20230728")
        csv_list = files_list(csv_path, ".csv")
        # 複製標準樣本
        sample_csv_path = os.path.join(root_path, "08Reference_Files/stander_csv.csv")
        test_csv_path = os.path.join(root_path, "00dest/stander_csv.csv")
        shutil.copy2(sample_csv_path, test_csv_path)
        # 合併清單中檔案
        combine_csv_list(csv_list, test_csv_path)
        # 查看清單檔案
        read_csv(test_csv_path)

    if xx_judge == 4:
        test_ht_to_csv()
    
    # make_stander_csv()
    # test_ht_to_csv()
    # test_txt_to_df()
    # test_read_csv()
    # test_combine_csv()
    # test_combine_csv_list()
