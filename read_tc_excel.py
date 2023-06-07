#cd C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station
#pyinstaller -F Read_TC_Excel.py

import os
import openpyxl
import pandas as pd

def read_tc_excel(def_dest_folder):
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
    for root, dirs, files in os.walk(def_dest_folder):
        for file in files:
            #比對副檔名條件
            "測試0"
            if file.endswith(file_type) * file.count("-") > 3:
                "測試1"
                print("檔案名稱",file)
                file_path = os.path.join(root, file)
                #確認 1.無檔案 2.非有效Excel檔
                try:
                    # 開啟 Excel 檔案
                    workbook = openpyxl.load_workbook(file_path)
                    # 選取Data1工作表
                    worksheet = workbook['Data1']
                    #讀取表格文件數量
                    drawings_nums = 0
                    for i in range(0,19):
                      if worksheet["A"+str(2+i)].value:
                        drawings_nums = 1 + drawings_nums
                      else:
                          break
                    print("文件數量",drawings_nums)
                    #讀取資料
                    for i in range(0,drawings_nums):
                        #print("讀取路徑", file_path, "第",i+1,"個檔案", )
                        drawing_no_value = worksheet["J"+str(2+i)].value
                        drawing_title_value = worksheet["F"+str(2+i)].value
                        drawing_vision_value = worksheet["G"+str(2+i)].value
                        letter_num_value = worksheet["A2"].value
                        letter_date_value = worksheet["C2"].value
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
                except FileNotFoundError:
                    # 處理找不到檔案的錯誤
                    print("找不到檔案:", file_path)
                except openpyxl.utils.exceptions.InvalidFileException:
                    # 處理無效的 Excel 檔案的錯誤
                    print("無效的 Excel 檔案:", file_path)

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

    #輸出Excel檔案
    basename = os.path.basename(file_path)
    filename_without_extension = os.path.splitext(basename)[0]
    NewExcelfile = "Done_" + filename_without_extension + ".xlsx"
    Output_path = os.path.join(def_dest_folder, NewExcelfile)
    #輸出成"Done_HT-D1-CTC-GEL-23-1171.xlsx"(範例)
    df.to_excel(Output_path, index=False)
    return letter_titl_value, drawing_vision_value, letter_num_value, letter_date_value

def main():
    "主程式"
    path = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC\TPC-TC(C0)-CD-23-0002"
    #convert_excel.convert_xlsb(path)
    print(read_tc_excel(path))
    input("Press enter to exit...")
    return None


def test():
    "測試程式"
    return None


if __name__ == '__main__':
    main()
