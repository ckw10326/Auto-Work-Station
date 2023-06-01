#cd C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station
#pyinstaller -F Read_HT_Excel.py
import os
import openpyxl
import pandas as pd
def Read_HT_Excel(folder_path):
    #遍歷所有資料夾、檔案
    Excute_Num_block = []
    DrawingNo_block = []#J2
    Drawing_Title_block = []#F2
    DrawinVision_block = []#G2
    Letter_Num_block = []#A2
    Letter_Date_block = []#C2
    Letter_title_block =[]#I2
    file_path_block = []
    filetype = ".xlsx"
    file_path = ""
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
                    #●●注意因為.xlsb轉檔為.xlsx，所以已不須讀取分頁●●
                    #worksheet = workbook['Data1']
                    worksheet = workbook.active
                    #讀取表格文件數量
                    DrawingsNums = 0
                    for i in range(0,19):
                        if worksheet["B"+str(2+i)].value:
                            DrawingsNums = 1 + DrawingsNums
                        else:
                            break
                    print("文件數量",DrawingsNums)

                    #讀取資料
                    for i in range(0,DrawingsNums):
                        #print("讀取路徑", file_path, "第",i+1,"個檔案", )
                        DrawingNo_value = worksheet["F"+str(2+i)].value
                        Drawing_Title_value = worksheet["G"+str(2+i)].value
                        DrawinVision_value = worksheet["H"+str(2+i)].value
                        Letter_Num_value = worksheet["B2"].value
                        Letter_Date_value = worksheet["D2"].value
                        Letter_titl_value = worksheet["M2"].value
                        
                        Excute_Num_block.append(i)
                        DrawingNo_block.append(DrawingNo_value)
                        Drawing_Title_block.append(Drawing_Title_value)
                        DrawinVision_block.append(DrawinVision_value)
                        Letter_Num_block.append(Letter_Num_value)
                        Letter_Date_block.append(Letter_Date_value)
                        Letter_title_block.append(Letter_titl_value)
                        file_path_block.append(file_path)
                        print("批次序號",i)
                        print("圖號:",DrawingNo_value)
                        print("圖名:",Drawing_Title_value)
                        print("版次:",DrawinVision_value)
                        print("來文號碼:",Letter_Num_value)
                        print("來文日期:",Letter_Date_value)
                        print("來文名稱:",Letter_titl_value)
                        print("路徑",file_path)
                        print("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<")
                        
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

    '''
    #輸出Excel檔案
    basename  = os.path.basename(file_path)
    print(basename)
    filename_without_extension = os.path.splitext(basename)[0]
    NewExcelfile = "Done_" +  filename_without_extension + ".xlsx"
    Output_path = os.path.join(folder_path, NewExcelfile)
    #輸出成"Done_HT-D1-CTC-GEL-23-1171.xlsx"(範例)
    df.to_excel(Output_path, index=False)
    '''
    return Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value

#Test_Path = input("請輸入路徑:")
Test_Path = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT\HT-D1-CTC-GEL-23-1188"
Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value = Read_HT_Excel(Test_Path)
print("來文名稱：", Letter_titl_value, "\n版次：" , DrawinVision_value, "\n來文號碼：", Letter_Num_value,"\n來文日期：", Letter_Date_value )
input("Press enter to exit...")
