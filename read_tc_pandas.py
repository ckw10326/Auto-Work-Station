'''
轉換台中計畫excel檔案
2.convert_xlsb      轉換符合條件的檔案，並輸出歷列表
3.read_ctc_ht_excel 讀取轉換後檔案，並逐項輸出
'''
import os
import openpyxl
import pandas as pd

def read_tc_excel(def_dest_path):
    """讀取台中文件"""
    if "Done" in def_dest_path:
        print("包含Done檔案，已處理過不再處理")
        return None
    # 遍歷所有資料夾、檔案
    keyword = {"CLIENTDOCNO":"圖號", "DOCVERSIONDESC":"圖名",
               "DOCREV":"版次", "TRANSMITTALNO":"來文號碼",
               "RETUREDATE":"來文日期", "DESCRIPTION":"來文名稱"}
    #遍歷所有資料夾、檔案
    excute_num_block = []
    drawing_no_block = []
    drawing_title_block = []
    drawing_vision_block = []
    letter_num_block = []
    letter_date_block = []
    letter_title_block = []
    file_path_block = []
    file_type = ".xlsm"
    file_path = ""
    drawing_no_value = ""
    drawing_title_value = ""
    drawing_vision_value = ""
    letter_num_value = ""
    letter_date_value = ""
    letter_titl_value = ""

    # 開啟 Excel 檔案
    workbook = openpyxl.load_workbook(def_dest_path)
    # 選取Data1工作表
    worksheet = workbook['Data1']
    # 讀取表格文件數量
    drawings_nums = 0
    for i in range(0, 19):
        if worksheet["A"+str(2+i)].value:
            drawings_nums = 1 + drawings_nums
        else:
            break
    print("文件數量", drawings_nums)
    # 讀取資料
    for i in range(0, drawings_nums):
        # print("讀取路徑", file_path, "第",i+1,"個檔案", )
        drawing_no_value = worksheet["J"+str(2+i)].value
        drawing_title_value = worksheet["F"+str(2+i)].value
        drawing_vision_value = worksheet["G"+str(2+i)].value
        letter_num_value = worksheet["A2"].value
        letter_date_value = worksheet["D2"].value
        letter_titl_value = worksheet["I2"].value

        excute_num_block.append(i)
        drawing_no_block.append(drawing_no_value)
        drawing_title_block.append(drawing_title_value)
        drawing_vision_block.append(drawing_vision_value)
        letter_num_block.append(letter_num_value)
        letter_date_block.append(letter_date_value)
        letter_title_block.append(letter_titl_value)
        file_path_block.append(file_path)
        print("批次序號", i)
        print("圖號:", drawing_no_value)
        print("圖名:", drawing_title_value)
        print("版次:", drawing_vision_value)
        print("來文號碼:", letter_num_value)
        print("來文日期:", letter_date_value)
        print("來文名稱:", letter_titl_value)
        print("路徑", file_path)
        print("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<")
    workbook.close()

    # 寫入檔案
    data = {'批次序號': excute_num_block,
            '圖號': drawing_no_block,
            '圖名': drawing_title_block,
            '版次': drawing_vision_block,
            '來文號碼': letter_num_block,
            '來文日期': letter_date_block,
            '來文名稱': letter_title_block,
            '路徑': file_path_block
            }
    df_data = pd.DataFrame(data)

    #輸出Excel檔案
    filename_without_extension = os.path.splitext(def_dest_path)[0]
    output_path  = filename_without_extension + "_Done.xlsx"
    print("輸出路徑:", output_path)
    #輸出成路徑 + "HT-D1-CTC-GEL-23-1171_Done.xlsx"
    df_data.to_excel(output_path, index=False)
    print("--------------.xlsx分析完成-----------------")
    return letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value

def xlsm_to_csv(file_path):
    """
    輸入xlsm路徑，產生csv檔案，並輸出該路徑
    台中計畫轉換xlsm to csv
    """
    if '.xlsm' in file_path:
        print(file_path, "，符合「.xlsm」的檔案")
        # 產生新檔名
        csv_path = os.path.splitext(file_path)[0] + ".csv"
        # 抓到隱藏檔案(檔名有關鍵字：~$H，不動作
        if "~$" in file_path:
            print(file_path, "抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
        else:
            if os.path.exists(csv_path):
                print(csv_path, "檔案已存在，不動作")
                return False
            else:
                sheet_name = 'Data1' 
                data_frame = pd.read_excel(file_path, sheet_name=sheet_name)
                data_frame.to_csv(csv_path, encoding='utf-8', index=True)
                print("輸出檔案：", csv_path)
                print(
                    r"3.----------------------Convert_xlsm to csv Done!------------------------\n")
                return csv_path
    else:
        return None

def test2() -> None:
    '''測試二'''
    path = r"/workspaces/Auto-Work-Station/00source"
    #列表，檔案清單
    the_xlsb_file_list = file_process.files_list(path, ".xlsm")

    #讀取列表中的清單
    for the_file in the_xlsb_file_list:
        read_tc_excel(the_file)

def main():
    "main"
    return None

def test_all():
    """測試to_df_analyze"""
    convert_xls = "/workspaces/Auto-Work-Station/00dest/00source/HT-D1-CTC-GEL-23-2146/HT-D1-CTC-GEL-23-2146_converted.xlsx"
    to_df_analyze(convert_xls)
    return None


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
