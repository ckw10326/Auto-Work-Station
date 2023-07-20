'''
轉換興達計畫excel檔案
'''
import os
import fnmatch
import openpyxl
import pandas as pd
import pyxlsb

'''
預設只能讀取一個檔案
2023/6/26 補充多個檔案
'''
def read_ctc_ht_excel(def_dest_folder):
    # 遍歷所有資料夾、檔案
    keyword = {"CLIENTDOCNO":"圖號", "DOCVERSIONDESC":"圖名",
               "DOCREV":"版次", "TRANSMITTALNO":"來文號碼",
               "RETUREDATE":"來文日期", "DESCRIPTION":"來文名稱"
                }
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

    for root, dirs, files in os.walk(def_dest_folder):
        for file in files:
            #比對副檔名條件，跳過含有Done的檔案
            if file.endswith(file_type) * file.count("-") * ("Done" not in file)> 3:
                file_path = os.path.join(root, file)
                #確認 1.無檔案 2.非有效Excel檔
                try:
                    # 開啟 Excel 檔案
                    workbook = openpyxl.load_workbook(file_path)
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
                except FileNotFoundError:
                    # 處理找不到檔案的錯誤
                    print("找不到檔案:", file_path)
                except openpyxl.utils.exceptions.InvalidFileException:
                    # 處理無效的 Excel 檔案的錯誤
                    print("無效的 Excel 檔案:", file_path)
            else:
                print("無興達Excel關鍵字「.xlsx」")

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
    basename = os.path.basename(file_path)
    filename_without_extension = os.path.splitext(basename)[0]
    NewExcelfile = "Done_" + filename_without_extension + ".xlsx"
    Output_path = os.path.join(def_dest_folder, NewExcelfile)
    #輸出成"Done_HT-D1-CTC-GEL-23-1171.xlsx"(範例)
    df.to_excel(Output_path, index=False)
    return letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value

'''
轉換資料夾內檔案
'''
def convert_xlsb(folder_path):
    xlsx_path = ""
    "尋找檔案是否存在"
    if os.path.exists(folder_path):
        print("convert_xlsb，找到目標資料夾：", folder_path)
        for file in os.listdir(folder_path):
            #確認xlsb檔案是否存在
            print("convert_xlsb，", file)
            if fnmatch.fnmatch(file, '*.xlsb'):
                file_path =os.path.join(folder_path, file)
                print("convert_xlsb，找到目標檔案路徑：", file_path, "有符合「.xlsb」的檔案")
                #抓到隱藏檔案(檔名有關鍵字：~$H，不動作
                if "~$" in file_path:
                    print(file_path,"convert_xlsb，抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
                else:
                    xlsx_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
                    print("convert_xlsb，輸出檔案：", xlsx_path)
                    if os.path.exists(xlsx_path):
                        print("convert_xlsb，_converted.xlsx" + "檔案已存在，不動作")
                    else:
                        data_frame = pd.read_excel(file_path, sheet_name='Data1', engine='pyxlsb')
                        data_frame.to_excel(xlsx_path)
                        print(r"----------------------Convert_xlsb Done!------------------------")
                        input("enter any keys to exit")
            else:
                print("convert_xlsb，找不到目標檔案：", file)
        return xlsx_path
    else:
        "若找不到資料夾，則回傳None"
        print("convert_xlsb，找不到目標資料夾：", folder_path)
        return None

def test1():
    "主程式"
    path = r"/workspaces/Auto-Work-Station/00source"
    convert_xlsb(path)
    read_ctc_ht_excel(path)
    input("測試功能完成...")

    return None


def main():
    "main"
    return None


if __name__ == '__main__':
    test1()
