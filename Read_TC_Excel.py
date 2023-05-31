
import os
import openpyxl
import pandas as pd

#輸入 目錄路徑
#輸出 1.螢幕相關資料   2.整個資料夾的Excel檔案彙整

def Read_TC_Excel(folder_path):
    #遍歷所有資料夾、檔案
    Excute_Num_block = []
    DrawingNo_block = []#J2
    Drawing_Title_block = []#F2
    DrawinVision_block = []#G2
    Letter_Num_block = []#A2
    Letter_Date_block = []#C2
    Letter_title_block =[]#I2
    file_path_block = []
    filetype = ".xlsm"
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            #比對副檔名條件
            if file.endswith(filetype) * file.count("-") > 3:
                file_path = os.path.join(root, file)
                #確認 1.無檔案 2.非有效Excel檔
                try:
                    # 開啟 Excel 檔案
                    workbook = openpyxl.load_workbook(file_path)
                    # 選取Data1工作表
                    worksheet = workbook['Data1']
                    #讀取表格文件數量
                    DrawingsNums = 0
                    for i in range(0,19):
                      if worksheet["A"+str(2+i)].value:
                        DrawingsNums = 1 + DrawingsNums
                      else:
                          break
                    #讀取資料
                    for i in range(0,DrawingsNums):
                        #print("讀取路徑", file_path, "第",i+1,"個檔案", )
                        DrawingNo_value = worksheet["J"+str(2+i)].value
                        Drawing_Title_value = worksheet["F"+str(2+i)].value
                        DrawinVision_value = worksheet["G"+str(2+i)].value
                        Letter_Num_value = worksheet["A2"].value
                        Letter_Date_value = worksheet["C2"].value
                        Letter_titl_value = worksheet["I2"].value
                        
                        Excute_Num_block.append(i)
                        DrawingNo_block.append(DrawingNo_value)
                        Drawing_Title_block.append(Drawing_Title_value)
                        DrawinVision_block.append(DrawinVision_value)
                        Letter_Num_block.append(Letter_Num_value)#來文號碼B2
                        Letter_Date_block.append(Letter_Date_value)#來文日期H2
                        Letter_title_block.append(Letter_titl_value)#來文標題O2
                        file_path_block.append(file_path)

                        print("------------------------")
                        print("批次序號",i)
                        print("圖號:",DrawingNo_value)
                        print("圖名:",Drawing_Title_value)
                        print("版次:",DrawinVision_value)
                        print("來文號碼:",worksheet["A2"].value)
                        print("來文日期:",worksheet["C2"].value)
                        print("來文名稱:",worksheet["I2"].value)
                        print("路徑",file_path)
                        print("------------------------")
                        print(Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value)
                        return Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value
                    workbook.close()
                except FileNotFoundError:
                    # 處理找不到檔案的錯誤
                    print("找不到檔案:", file_path)
                except openpyxl.utils.exceptions.InvalidFileException:
                    # 處理無效的 Excel 檔案的錯誤
                    print("無效的 Excel 檔案:", file_path)

    data = {'批次序號': Excute_Num_block,
            '圖號': DrawingNo_block,
            '圖名': Drawing_Title_block,
            '版次': DrawinVision_block,
            '來文號碼': Letter_Num_block,
            '來文日期': Letter_Date_block,
            '來文名稱': Letter_title_block,
            '路徑': file_path_block
            }
    df = pd.DataFrame(data)

    #輸出Excel檔案
    OutputName = 'TC_output.xlsx'
    Output_path = os.path.join(folder_path, OutputName)
    df.to_excel(Output_path, index=False)
    print("完成輸出檔案\n路徑:",Output_path,"\n檔名:",OutputName)

#全路徑參考目錄
Test_Path = r"/content/sample_data/"
Cloud_path = r"/content/sample_data/TC"
Home_path = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC"
Company_path = r"D:\\00 興達計劃\\05 EPC提供資料\\TC/0TPC-TC(C0)-CD-23-0004"
path_map = {"0":Test_Path, "1" : Cloud_path, "2" : Home_path, "3" : Company_path}
#調試路徑

Test_Path = input("請輸入路徑:")
Read_TC_Excel(Test_Path)
