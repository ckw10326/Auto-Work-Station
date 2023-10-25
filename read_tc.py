'''
轉換台中計畫excel檔案
read_tc_excel     讀取轉換後檔案，並逐項輸出
'''
import os
import openpyxl
import pandas as pd

def r_tc_xlsm_sheet(def_dest_path):
    """
    讀取台中文件
    輸出
        列表各字串列表
        產生excel檔案
    """
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

    # 輸出Excel檔案
    filename_without_extension = os.path.splitext(def_dest_path)[0]
    output_path = filename_without_extension + "_Done.xlsx"
    print("輸出路徑:", output_path)
    # 輸出成路徑 + "HT-D1-CTC-GEL-23-1171_Done.xlsx"
    df_data.to_excel(output_path, index=False)
    print("--------------.xlsx分析完成-----------------")
    return letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value


if __name__ == '__main__':
    print("請運行main.py")
