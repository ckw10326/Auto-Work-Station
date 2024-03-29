'''
轉換興達計畫excel檔案 2023/08/28
1.r_ctc_xlsb_sheet
    輸入xlsx檔案，依照位置分析頁首
2.convert_xlsb      轉換符合條件的檔案，並輸出轉換後路徑
'''
import os
import openpyxl
import pandas as pd

def r_ctc_xlsx_sheet(def_dest_path):
    """讀取HT文件"""
    if "Done" in def_dest_path:
        print("包含Done檔案，已處理過不再處理")
        return None
    # 遍歷所有資料夾、檔案
    excute_num_block = []
    drawing_no_block = []
    drawing_title_block = []
    drawing_vision_block = []
    letter_num_block = []
    letter_date_block = []
    letter_title_block = []
    file_path_block = []
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
    # ●●注意因為.xlsb轉檔為.xlsx，所以已不須讀取分頁●●
    # worksheet = workbook['Data1']
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
    data_frame = pd.DataFrame(data)

    # 輸出Excel檔案
    filename_without_extension = os.path.splitext(def_dest_path)[0]
    # 檢查是否有內容，有則輸出_Done，無則_Not
    if letter_num_block and letter_title_block and drawing_vision_block:
        output_path = filename_without_extension + "_Done.xlsx"
    else:
        output_path = filename_without_extension + "_Not.xlsx"
    print("輸出路徑:", output_path)
    # 輸出成路徑 + "HT-D1-CTC-GEL-23-1171_Done.xlsx"
    data_frame.to_excel(output_path, index=False)
    print("--------------.xlsx分析完成-----------------")
    return letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value

def convert_xlsb(file_path):
    """
    轉換xlsb to xlsx
    讀取計畫    HT
    讀取關鍵字  CTC-GEL
    讀取類型    XLSB
    方法        engine = pyxlsb,  分頁DATA1
    結果        在原始路徑輸出CSV檔案
    輸出        csv檔案路徑
    """
    if ".xlsb" in file_path:
        print(" (1).convert_xlsb，檔案路徑：", file_path, "，符合「.xlsb」的檔案")
        converter_xlsx = os.path.splitext(file_path)[0] + "_converted.xlsx"
        print("converter_xlsx=" ,converter_xlsx)
        # 抓到隱藏檔案(檔名有關鍵字：~$H，不動作
        if "~$" in file_path:
            print(file_path, " (2).convert_xlsb，抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
        else:
            if os.path.exists(converter_xlsx):
                # print("3.convert_xlsb，_converted.xlsx" + "檔案已存在，不動作")
                pass
            else:
                # os.path.splitext [0] = 檔案名稱，讀取分頁名稱Data1
                data_frame = pd.read_excel(
                    file_path, sheet_name='Data1', engine='pyxlsb')
                data_frame.to_excel(converter_xlsx)
                print(" (2).convert_xlsb，輸出檔案：", converter_xlsx)
                print(
                    r" (3).------Convert_xlsb Done!--------\n")
                return converter_xlsx
    else:
        return None

if __name__ == '__main__':
    print("請運行main.py")
