# -*- coding: utf-8 -*-

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

# 輸入：讀取excel檔案路徑，輸出csv檔案路徑
# 1.讀取指定路徑檔案
# 2.新增到csv檔案中
def read_ctc_ht_excel0(excel_path, csv_path = "/workspaces/Auto-Work-Station/01Class/data.csv"):
    if "Done" in excel_path:
        print("包含Done檔案，已處理過不再處理")
        return None
    # 建立首列DataFrame
    df = pd.DataFrame(columns=['批次序號', '圖號:', '圖名:', '版次:', '來文號碼:', 
                               '來文日期:', '來文名稱:', '路徑'])

    # XX已失效，但先保留遍歷所有資料夾、檔案
    keyword = {"CLIENTDOCNO":"圖號", "DOCVERSIONDESC":"圖名",
               "DOCREV":"版次", "TRANSMITTALNO":"來文號碼",
               "RETUREDATE":"來文日期", "DESCRIPTION":"來文名稱"
                }
    # 遍歷所有資料夾、檔案
    file_type = ".xlsx"

    # 1.讀取 Excel 檔案
    workbook = openpyxl.load_workbook(excel_path)

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
            '批次序號': i+1,
            '圖號:': drawing_no_value,
            '圖名:': drawing_title_value,
            '版次:': drawing_vision_value,
            '來文號碼:': letter_num_value,
            '來文日期:': letter_date_value,
            '來文名稱:': letter_titl_value,
            '路徑': excel_path
        }
        df = pd.concat([df, pd.DataFrame(row_data, index=[0])], ignore_index=True)
    #將每一列讀取完成後，看結果
    csv_path = "/workspaces/Auto-Work-Station/01Class/data.csv"
    csv_file_path2 = "/workspaces/Auto-Work-Station/00source/data.csv"
    if os.path.exists(csv_path):
        df.to_csv(csv_path, mode='a', header=False, index=False)
    else:
        df.to_csv(csv_path, index=False)
    workbook.close()

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

#測試pandas
def main():
    #測試輸入無內容Excel檔案是否會輸出not資料
    sample_folder = "/workspaces/Auto-Work-Station/00source"
    destination_dir = "/workspaces/Auto-Work-Station/00dest"

    if os.path.exists(destination_dir):
        shutil.rmtree(destination_dir)

    #複製Sample資料夾結構
    file_process.move_document(sample_folder, destination_dir)
    print("複製destination_dir到00dest資料夾完成", destination_dir, "\n")
    
    the_xlsb_file_list = file_process.files_list(destination_dir, ".xlsb")
    #若是清單有內容才會帶入解析
    for filepath in the_xlsb_file_list:
        print("符合興達計畫.xlsb清單", filepath)
        #若符合.xlsb，則轉檔
        convert_xlsb(filepath)

    #將.xlsb轉換成.xlsx後，開始分析內容
    the_xlsx_file_list = file_process.files_list(destination_dir, "converted.xlsx")
    #若是清單有內容才會帶入解析
    for the_file in the_xlsx_file_list:
        print("符合興達計畫converted.xlsb清單:",the_file)
        input("enter any keys to exit")
        read_ctc_ht_excel0(the_file)
    
    csv_file_path = "/workspaces/Auto-Work-Station/00source/data.csv"
    data_frame = pd.read_csv(csv_file_path)
    print(data_frame)

if __name__ == '__main__':
    main()