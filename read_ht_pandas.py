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


def read_pandas_sec(excel_path):
    """
    # 輸入，讀取excel檔案路徑，輸出csv檔案路徑
    # 1.讀取指定路徑檔案，並讀取表格上指定位置
    # 2.自訂義相關文字
    # 3.新增到csv檔案中
    """
    # 建立空的 DataFrame
    df = pd.DataFrame()
    file_path = []
    drawing_no = []
    draw_title = []
    draw_vision = []
    l_num = []
    l_date = []
    l_title = []
    # 1.讀取 Excel 檔案
    workbook = openpyxl.load_workbook(excel_path)
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
        file_path.append(excel_path)

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

    df = pd.DataFrame(row_data)
    print(df)

    # 分割檔案名稱、副檔名 > 儲存CSV檔案
    filename, _ = os.path.splitext(excel_path)
    # 將字串中的"_converted"剝離
    csv_path = filename.replace("_converted", "") + "_csv.csv"
    df.to_csv(csv_path, index=True)

    workbook.close()
    return csv_path

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

def xlsb_to_csv(file_path, workbook = None):
    """轉換xlsb to csv"""
    if '.xlsb' in file_path:
        print("1.convert_xlsb，檔案路徑：", file_path, "，符合「.xlsb」的檔案")
        converter_csv = os.path.splitext(file_path)[0] + ".csv"
        # 抓到隱藏檔案(檔名有關鍵字：~$H，不動作
        if "~$" in file_path:
            print(file_path, "2.convert_xlsb，抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
        else:
            if os.path.exists(converter_csv):
                # print("3.convert_xlsb，_converted.xlsx" + "檔案已存在，不動作")
                pass
            else:
                # os.path.splitext [0] = 檔案名稱，讀取分頁名稱Data1
                # 若有指定分頁名稱，則使用sheetname功能
                if workbook:
                    data_frame = pd.read_excel(file_path, sheet_name = workbook, engine='pyxlsb')
                else:
                    data_frame = pd.read_excel(file_path, engine='pyxlsb')
                data_frame.to_csv(converter_csv)
                print(data_frame)
                print("2.convert_csv，輸出檔案：", converter_csv)
                print(
                    r"3.----------------------Convert_xlsb Done!------------------------\n")
                return converter_csv
    else:
        return None

def clean_csv(file_path, 
              threshold = None, 
              null_num = 1):
    '''
    threshold，用於指定保留至少多少個非空值的行或列
    nullnum，用於保留那些缺少值數量少於2的行
    '''
    # 讀取 CSV 檔案
    df = pd.read_csv(file_path)

    if threshold is not None:
        # 如果 `threshold` 不是 `None`，即有指定閾值
        clean_df = df.dropna(thresh=threshold)
    
    elif null_num is not None:
        # 只保留缺少值數量少於 null_num 的行
        clean_df = df[df.isnull().sum(axis=1) < null_num]
    
    print(clean_df.to_string)
    return clean_df

    # 將只包含完整資料的資料行寫入新的 CSV 檔案
    #clean_df.to_csv('cleaned_data.csv', index=False)
    #print("已整理並輸出整潔的資料到 cleaned_data.csv 檔案中。")


def to_df_analyze(excel_path, combined_csv_path="/workspaces/Auto-Work-Station/01Class/data.csv"):
    """ 測試df = pd.read_excel格式是否會跑掉"""
    if "Done" in excel_path:
        print("包含Done檔案，已處理過不再處理")
        return None

    # 建立首列DataFrame
    df = pd.DataFrame(columns=['批次序號', '圖號:', '圖名:', '版次:', '來文號碼:',
                               '來文日期:', '來文名稱:', '路徑'])
    # 表格映射值
    field_mapping = {
        'No': ['批次序號'],
        'drawing_no_value': ['圖號', 'CLIENTDOCNO', 'DOCVERSIONDESC'],
        'drawing_title_value': ['圖名', 'DESCRIPTION', 'DOCCLASS'],
        'drawing_vision_value': ['版次', 'DOCVERSIONDESC'],
        'letter_num_value': ['來文號碼', 'TRANSMITTALNO'],
        'letter_date_value': ['來文日期', 'REVDATE', 'PLANNEDCLIENTRETURNDATE'],
        'letter_titl_value': ['來文名稱', 'DESCRIPTION'],
        'file_path': ['路徑']
    }

    # 1.讀取 Excel 檔案，設置為df
    df = pd.read_excel(excel_path)
    print(df, '\n------------------------------------')

    field_mapping = {
        'No': ['批次序號'],
        'drawing_no_value': ['圖號', 'CLIENTDOCNO', 'DOCVERSIONDESC'],
        'drawing_title_value': ['圖名', 'DESCRIPTION', 'DOCCLASS'],
        'drawing_vision_value': ['版次', 'DOCVERSIONDESC'],
        'letter_num_value': ['來文號碼', 'TRANSMITTALNO'],
        'letter_date_value': ['來文日期', 'REVDATE', 'PLANNEDCLIENTRETURNDATE'],
        'letter_titl_value': ['來文名稱', 'DESCRIPTION'],
        'file_path': ['路徑']
    }

    # 指定要移動到第一列和第二列的表頭
    code_headers = ['No', 'drawing_no_value', 'drawing_title_value', 'drawing_vision_value',
                    'letter_num_value', 'letter_date_value', 'letter_titl_value', 'file_path']
    zh_headers = ['批次序號', '圖號', '圖名', '版次', '來文號碼', '來文日期', '來文名稱', '路徑']
    vender1_zh_headers = ['', 'CLIENTDOCNO', 'DESCRIPTION',
                          'DOCVERSIONDESC', 'TRANSMITTALNO', 'REVDATE', 'DESCRIPTION', '']
    vender2_zh_headers = ['', 'DOCVERSIONDESC', 'DOCCLASS',
                          '', '', 'PLANNEDCLIENTRETURNDATE', '', '']

    # 使用 reindex() 函數重新排序列
    df1 = df.reindex(columns=vender1_zh_headers)
    df2 = df.reindex(columns=vender2_zh_headers)
    print("start")
    print(df1, '\n------------------------------------')
    print(df2, '\n------------------------------------')

    # worksheet = workbook.active
    # # 讀取表格文件數量
    # drawings_nums = 0
    # for i in range(0, 19):
    #     if worksheet["B"+str(2+i)].value:
    #         drawings_nums = 1 + drawings_nums
    #     else:
    #         break
    # print("文件數量", drawings_nums)
    # # 讀取資料
    # for i in range(0, drawings_nums):
    #     # print("讀取路徑", file_path, "第",i+1,"個檔案", )
    #     drawing_no_value = worksheet["E"+str(2+i)].value
    #     drawing_title_value = worksheet["O"+str(2+i)].value
    #     drawing_vision_value = worksheet["G"+str(2+i)].value
    #     letter_num_value = worksheet["B2"].value
    #     letter_date_value = worksheet["H2"].value
    #     letter_titl_value = worksheet["O2"].value
    #     # 將資料添加到DataFrame中
    #     row_data = {
    #         '批次序號': i+1,
    #         '圖號:': drawing_no_value,
    #         '圖名:': drawing_title_value,
    #         '版次:': drawing_vision_value,
    #         '來文號碼:': letter_num_value,
    #         '來文日期:': letter_date_value,
    #         '來文名稱:': letter_titl_value,
    #         '路徑': excel_path}
    #     # 合併儲存格
    #     df = pd.concat([df, pd.DataFrame(row_data, index=[0])], ignore_index=True)

    # # 分割檔案名稱、副檔名 > 儲存CSV檔案
    # filename, extension = os.path.splitext(excel_path)
    # csv_path = filename.replace("_converted", ".csv")
    # df.to_csv(csv_path, mode='a', header=False, index=False)
    # print(df, "\n----------------以上為讀取表格-----------------")
    # # 整合舊資料與新資料
    # if os.path.exists(combined_csv_path):
    #     df.to_csv(combined_csv_path, mode='a', header=False, index=False)
    #     combined_df = pd.read_csv(combined_csv_path)
    #     print(combined_df, "\n----------------整合後表格-----------------")
    # else:
    #     df.to_csv(combined_csv_path, index=False)
    # workbook.close()
    # return None


if __name__ == '__main__':
    print("請執行main_pandas.py")
