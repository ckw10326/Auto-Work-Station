'''
轉換興達計畫excel檔案
1.files_list        表列所有檔案，並輸出列表
2.convert_xlsb      轉換符合條件的檔案，並輸出歷列表
3.read_ctc_ht_excel 讀取轉換後檔案，並逐項輸出
'''
import os
import fnmatch
import openpyxl
import pandas as pd
import pyxlsb
import shutil
import file_process

def read_ctc_ht_excel(def_dest_path):
    if "Done" in def_dest_path:
        print("包含Done檔案，已處理過不再處理")
        return None
    # 遍歷所有資料夾、檔案
    keyword = {"CLIENTDOCNO":"圖號", "DOCVERSIONDESC":"圖名",
               "DOCREV":"版次", "TRANSMITTALNO":"來文號碼",
               "RETUREDATE":"來文日期", "DESCRIPTION":"來文名稱"
                }
    #遍歷所有資料夾、檔案
    excute_num_block = []
    drawing_no_block = []
    drawing_title_block = []
    drawing_vision_block = []
    letter_num_block = []
    letter_date_block = []
    letter_title_block = []
    file_path_block = []
    file_type = ".xlsx"
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
    df = pd.DataFrame(data)

    # 輸出Excel檔案
    filename_without_extension = os.path.splitext(def_dest_path)[0]
    # 檢查是否有內容，有則輸出_Done，無則_Not
    print("len(letter_num_block)", len(letter_num_block))
    print("len(letter_title_block)", len(letter_title_block))
    print("len(drawing_vision_block)", len(drawing_vision_block))
    if len(letter_num_block) and len(letter_title_block) and len(drawing_vision_block):
        Output_path  = filename_without_extension + "_Done.xlsx"
    else:
        Output_path  = filename_without_extension + "_Not.xlsx"
    print("輸出路徑:", Output_path)
    #輸出成路徑 + "HT-D1-CTC-GEL-23-1171_Done.xlsx"
    df.to_excel(Output_path, index=False)
    print("--------------.xlsx分析完成-----------------")
    return letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value


def read_ctc_ht_excel_list(def_dest_path):
    if "Done" in def_dest_path:
        print("包含Done檔案，已處理過不再處理")
        return None
    # 遍歷所有資料夾、檔案
    keyword = {"CLIENTDOCNO":"圖號", "DOCVERSIONDESC":"圖名",
               "DOCREV":"版次", "TRANSMITTALNO":"來文號碼",
               "RETUREDATE":"來文日期", "DESCRIPTION":"來文名稱"
                }
    #遍歷所有資料夾、檔案
    excute_num_block = []
    drawing_no_block = []
    drawing_title_block = []
    drawing_vision_block = []
    letter_num_block = []
    letter_date_block = []
    letter_title_block = []
    file_path_block = []
    file_type = ".xlsx"
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
        #用來出示範的
        list22 = [drawing_no_value, drawing_title_value, drawing_vision_value, 
                  letter_num_value, letter_date_value, letter_titl_value]
        return list22
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
    df = pd.DataFrame(data)

    # 輸出Excel檔案
    filename_without_extension = os.path.splitext(def_dest_path)[0]
    # 檢查是否有內容，有則輸出_Done，無則_Not
    print("len(letter_num_block)", len(letter_num_block))
    print("len(letter_title_block)", len(letter_title_block))
    print("len(drawing_vision_block)", len(drawing_vision_block))
    if len(letter_num_block) and len(letter_title_block) and len(drawing_vision_block):
        Output_path  = filename_without_extension + "_Done.xlsx"
    else:
        Output_path  = filename_without_extension + "_Not.xlsx"
    print("輸出路徑:", Output_path)
    #輸出成路徑 + "HT-D1-CTC-GEL-23-1171_Done.xlsx"
    df.to_excel(Output_path, index=False)
    print("--------------.xlsx分析完成-----------------")
    return letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value

'''
輸入檔案名稱，函數內比對資格，轉換後輸出新路徑
這邊比較麻煩的是，當執行return converter_xlsx，分一個檔案後就會退出
'''
def convert_xlsb(file_path):
    if '.xlsb' in file_path:
        print("有抓取到.xlsb檔案")
        print("1.convert_xlsb，檔案路徑：", file_path, "，符合「.xlsb」的檔案")
        converter_xlsx = os.path.splitext(file_path)[0] + "_converted.xlsx"
        #抓到隱藏檔案(檔名有關鍵字：~$H，不動作
        if "~$" in file_path:
            print(file_path,"2.convert_xlsb，抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
        else: 
            if os.path.exists(converter_xlsx):
                #print("3.convert_xlsb，_converted.xlsx" + "檔案已存在，不動作")
                pass
            else:
                #os.path.splitext [0] = 檔案名稱
                print("2.convert_xlsb，將要輸出檔案：", converter_xlsx)
                data_frame = pd.read_excel(file_path, sheet_name='Data1', engine='pyxlsb')
                data_frame.to_excel(converter_xlsx)
                print(r"3.----------------------Convert_xlsb Done!------------------------\n")
                input("enter any keys to exit")
                return converter_xlsx
    else:
        return None

"測試中程式"
def test1():
    path = r"/workspaces/Auto-Work-Station/00source"
    #列表，檔案清單
    the_xlsb_file_list = file_process.files_list1(path, ".xlsb")

    #列表，有轉換檔案後的清單
    for filepath in the_xlsb_file_list:
        convert_xlsb(filepath)
    print("convered_xlsb Done\n")

    #列表，輸出成converted.xlsx清單
    the_xlsx_file_list = file_process.files_list1(path, "converted.xlsx")
    for the_file in the_xlsx_file_list:
        read_ctc_ht_excel(the_file)

"測試讀取錯誤時，輸出NOt"
def test2():
    path = r"/workspaces/Auto-Work-Station/00source"
    dest = r"/workspaces/Auto-Work-Station/00dest"
    if 0:
        #刪除dest檔案
        if os.path.exists(dest):
            shutil.rmtree(dest)
        #列表，檔案清單
        the_xlsb_file_list = file_process.files_list1(path, ".xlsb")

        #列表，有轉換檔案後的清單
        if len(the_xlsb_file_list):
            for filepath in the_xlsb_file_list:
                convert_xlsb(filepath)
            print("convered_xlsb Done\n")

    #列表，輸出成converted.xlsx清單
    the_xlsx_file_list = file_process.files_list1(path, "converted.xlsx")
    for the_file in the_xlsx_file_list:
        read_ctc_ht_excel(the_file)

def main():
    #測試輸入無內容Excel檔案是否會輸出not資料
    path = "/workspaces/Auto-Work-Station/00source/HT-D1-CTC-GEL-23-2867_converted.xlsx"
    print(read_ctc_ht_excel_list(path))
    return None


if __name__ == '__main__':
    main()