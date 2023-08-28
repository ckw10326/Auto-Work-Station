'''
轉換興達計畫excel檔案，大量使用pandas
1.read_ht_pandas
2.convert_xlsb      轉換符合條件的檔案，並輸出歷列表
3.read_ctc_ht_excel 讀取轉換後檔案，並逐項輸出
'''
import os
import openpyxl
import pandas as pd
from read_ht import convert_xlsb

def read_pandas(excel_path, combined_csv_path="/workspaces/Auto-Work-Station/01Class/data.csv"):
    """
    # 輸入：讀取excel檔案路徑，輸出csv檔案路徑
    # 1.讀取指定路徑檔案
    # 2.新增到csv檔案中
    """

    # 建立首列DataFrame
    # df = pd.DataFrame(columns=['批次序號', '圖號:', '圖名:', '版次:', '來文號碼:',
    #                            '來文日期:', '來文名稱:', '路徑'])
    # 遍歷所有資料夾、檔案
    file_type = ".xlsx"
    file_path = excel_path
    drawing_no_value = ""
    drawing_title_value = ""
    drawing_vision_value = ""
    letter_num_value = ""
    letter_date_value = ""
    letter_titl_value = ""

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
        # print("讀取路徑", file_path, "第",i+1,"個檔案", )
        drawing_no_value = worksheet["E"+str(2+i)].value
        drawing_title_value = worksheet["O"+str(2+i)].value
        drawing_vision_value = worksheet["G"+str(2+i)].value
        letter_num_value = worksheet["B2"].value
        letter_date_value = worksheet["H2"].value
        letter_titl_value = worksheet["O2"].value
        # 將資料添加到DataFrame中
        row_data = {
            '批次序號': i,
            '圖號:': drawing_no_value,
            '圖名:': drawing_title_value,
            '版次:': drawing_vision_value,
            '來文號碼:': letter_num_value,
            '來文日期:': letter_date_value,
            '來文名稱:': letter_titl_value,
            '路徑': excel_path}
        # 合併儲存格
        df = pd.DataFrame(row_data, index = [])

    # 分割檔案名稱、副檔名 > 儲存CSV檔案
    filename, _ = os.path.splitext(excel_path)
    csv_path = filename.replace("_converted", ".csv")
    df.to_csv(csv_path, mode='a', header=False, index=False)
    print(df, "\n----------------以上為讀取表格-----------------")

    workbook.close()
    return None

def read_pandas_sec(excel_path, combined_csv_path="/workspaces/Auto-Work-Station/01Class/data.csv"):
    """
    # 輸入：讀取excel檔案路徑，輸出csv檔案路徑
    # 1.讀取指定路徑檔案
    # 2.新增到csv檔案中
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
    return None

def combine_csv(src, dst):
    try:
        # 讀取源檔案
        df_src = pd.read_csv(src)
        # 讀取目標檔案
        df_dst = pd.read_csv(dst)
        # 合併檔案
        combined = pd.concat([df_src, df_dst], ignore_index=True)
        # 寫入合併後的檔案
        combined.to_csv(dst, index=False)
        print("檔案合併完成")
        
    except FileNotFoundError:
        print("找不到檔案")
        
    except Exception as e:
        print("發生錯誤:", str(e))

def combine_csv_list(scr, dst):
    # 整合舊資料與新資料
    if os.path.exists(combined_csv_path):
        df.to_csv(combined_csv_path, mode='a', header=False, index=False)
        combined_df = pd.read_csv(combined_csv_path)
        print(combined_df, "\n----------------整合後表格-----------------")
    else:
        print("合併檔案不存在")

def read_csv(file_path):
    try:
        # 讀取 CSV 文件
        df = pd.read_csv(file_path)
        print(df.iloc[:5, :5])
    except FileNotFoundError:
        print("找不到文件")
    except Exception as e:
        print("發生錯誤:", str(e))

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
