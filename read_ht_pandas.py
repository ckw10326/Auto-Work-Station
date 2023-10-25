'''
轉換興達計畫excel檔案，大量使用pandas
1.read_pandas_sec
  輸入excel路徑，輸出csv檔案到路徑底下
2.convert_xlsb      轉換符合條件的檔案，並輸出歷列表
3.read_ctc_ht_excel 讀取轉換後檔案，並逐項輸出
'''
import os
import openpyxl
import pandas as pd
from function_file_process import move_document, files_list, check_plan, del_folder, make_folder

def r_ht_ctc_xlsx_sheet_tocsv(converted_xlsx_path):
    """
    前置動作    convert_xlsb(XLSB>XLSX)
    讀取計畫    HT
    讀取關鍵字  CTC-GEL
    讀取類型    XLSX
    方法        SHEET代表讀取表格指定位置
    結果        在原始路徑輸出CSV檔案
    輸出        csv檔案路徑
    """
    # 建立空的 DataFrame
    data_frame = pd.DataFrame()
    file_path = []
    drawing_no = []
    draw_title = []
    draw_vision = []
    l_num = []
    l_date = []
    l_title = []
    # 1.讀取 Excel 檔案
    workbook = openpyxl.load_workbook(converted_xlsx_path)
    worksheet = workbook.active
    # 讀取表格文件數量
    drawings_nums = 0
    for i in range(0, 19):
        if worksheet["B"+str(2+i)].value:
            drawings_nums = 1 + drawings_nums
        else:
            break
    print("文件數量", drawings_nums)
    # 讀取資料
    for i in range(0, drawings_nums):
        drawing_no.append(worksheet["E"+str(2+i)].value)
        draw_title.append(worksheet["O"+str(2+i)].value)
        draw_vision.append(worksheet["G"+str(2+i)].value)
        l_num.append(worksheet["B2"].value)
        l_date.append(worksheet["H2"].value)
        l_title.append(worksheet["O2"].value)
        file_path.append(converted_xlsx_path)

    # 將資料添加到DataFrame中
    row_data = {
        '批次序號': i,
        '圖號:': drawing_no,
        '圖名:': draw_title,
        '版次:': draw_vision,
        '來文號碼:': l_num,
        '來文日期:': l_date,
        '來文名稱:': l_title,
        '路徑': file_path}

    data_frame = pd.DataFrame(row_data)
    print(data_frame)

    # 分割檔案名稱、副檔名 > 儲存CSV檔案
    filename, _ = os.path.splitext(converted_xlsx_path)
    # 將字串中的"_converted"剝離
    csv_path = filename.replace("_converted", "") + "_csv.csv"
    data_frame.to_csv(csv_path, index=True)

    workbook.close()
    return csv_path

def gei_xlsx(excel_path):
    """
    測試HT-D1-GEI-GEL-23-0290
    輸入路徑，輸出df表格，表格應包含關鍵字
    """
    # 讀取xlsx檔案
    
    # 刪除無用

def test_gei_xlsx():
    excel_path = "/workspaces/Auto-Work-Station/08Reference_Files/HT/HT-D1-GEI-GEL-23-0290.xlsx"
    gei_xlsx(excel_path)

def xlsx_to_csv(excel_path):
    """
    # 輸入，讀取excel檔案路徑
    # 1.直接將表格轉換成csv檔案，【不做指定位置輸出】
    """
    # 建立空的 DataFrame
    workbook = openpyxl.load_workbook(excel_path)
    worksheet = workbook.active

    # 將工作表轉換成 DataFrame
    data = worksheet.values
    columns = next(data)  # 使用第一列作為欄位名稱
    data_frame = pd.DataFrame(data, columns=columns)

    # 儲存 DataFrame 為 CSV 檔案
    csv_output_path = excel_path.replace("_converted.xlsx", ".csv")
    data_frame.to_csv(csv_output_path, index=True)
    workbook.close()
    return csv_output_path


def xlsb_to_csv(file_path, workbook=None):
    """
    功能：將 xlsb 文件轉換為 csv
    注意：具備分頁選擇功能
    Parameters:
        file_path (str): xlsb 文件的路徑
        workbook (str or None): 指定要讀取的工作簿名稱
    Returns:
        str or None: 轉換後的 csv 文件路徑，如果不是 xlsb 文件則返回 None
    """

    # 檢查輸入參數合法性

    assert os.path.isfile(file_path), "Invalid file path"
    assert file_path.endswith(".xlsb"), "File format is not xlsb"

    # 處理隱藏檔案

    if "~$" in file_path:
        print(file_path, "2.convert_xlsb，抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
        return None

    # 生成輸出檔案名稱

    converter_csv = os.path.splitext(file_path)[0] + ".csv"

    # 讀取 xlsb 文件

    try:
        if workbook:
            data_frame = pd.read_excel(
                file_path, sheet_name=workbook, engine='pyxlsb')
        else:
            data_frame = pd.read_excel(file_path, engine='pyxlsb')
    except Exception as e:
        print(e)
        return None

    # 寫入 csv 文件

    with open(converter_csv, 'w', newline='') as f:
        data_frame.to_csv(f)

    return converter_csv
def test_xlsb_to_csv():
    file_path = "/workspaces/Auto-Work-Station/00source/HT-D1-CTC-GEL-23-0005.xlsb"
    workbook = "Data1"
    converter_csv = xlsb_to_csv(file_path, workbook)
    print(f"4.轉換完成，輸出檔案：{converter_csv}")


def read_all_xlsb_df(file_path):
    '''
    1.讀取XLSB全分業
    2.合併不同表格，
    3.轉換xlsb to csv 文在檔案的位置所在
    '''
    if not file_path.endswith('.xlsb'):
        # 若檔案不存在則不動作
        print("非xlsb檔案，請在確認")
        return None

    # 4.讀取檔案分頁
    df1 = read_xlsb_df(file_path, "Data1")
    # 新增path欄位
    df1['path'] = "dest_path"

    df2 = read_xlsb_df(file_path)
    df2 = clean_csv_letter_cover(df2, null_num=2)
    df2 = df2.reset_index(drop=True)

    # 進行合併操作
    merged_df = pd.concat([df1, df2], axis=1)
    return merged_df


def clean_csv_letter_cover(data_frame,
                           threshold=None,
                           null_num=1):
    '''
    似乎沒有使用必要，重要資訊Data1分頁都擁有
    目的：將複雜的cover頁轉化為DataFrame當輸入，輸出純淨的表格
    輸入：data_frame
    輸出：dataf_frame
    參數
    threshold：用於指定保留至少多少個非空值的行或列
    nullnum：用於保留那些缺少值數量少於2的行
    重新設定表頭
    '''
    # 讀取df
    temp_data_frame = data_frame

    if threshold is not None:
        # 如果 `threshold` 不是 `None`，即有指定閾值
        clean_df = temp_data_frame.dropna(thresh=threshold)

    elif null_num is not None:
        # 只保留缺少值數量少於 null_num 的行
        clean_df = temp_data_frame[temp_data_frame.isnull().sum(
            axis=1) < null_num]

    # 指定第一列為表頭
    # 設定 DataFrame 中df.iloc[0] 用於選取 DataFrame 的第一列數據
    clean_df.columns = clean_df.iloc[0]
    # 刪除第一列
    clean_df = clean_df[1:]

    return clean_df

    # 將只包含完整資料的資料行寫入新的 CSV 檔案
    # clean_df.to_csv('cleaned_data.csv', index=False)
    # print("已整理並輸出整潔的資料到 cleaned_data.csv 檔案中。")


if __name__ == '__main__':
    if 0:
        print("請執行main_pandas.py")
    else:
        test_xlsb_to_csv()
