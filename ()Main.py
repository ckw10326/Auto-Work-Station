#簡化修正版
#cd C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station
#pyinstaller -F ()Main.py
import shutil
import os #用來避免檔案已經存在的異常
import openpyxl
import fnmatch
import pandas as pd
import datetime
import pyxlsb
# TC_abbreviation = ["TPC", "TC(C0))", "CD", "23"]
# HT_abbreviation = ["HT", "D1", "CTC", "GEL", "23"]
Easy_Dict = {"1":"HT", "2":"TC", "99":""}
Plan_Dict = {"HT" : ["HT", "D1", "CTC", "GEL", "23"],
             "TC" : ["TPC", "TC(C0)", "CD", "23"]
             }
Cloud_HT = r"/content/sample_data/HT"
Cloud_TC = r"/content/sample_data/TC"
Home_HT = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT"
Home_TC = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC"
Source_Company_HT = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫"
destny_Company_HT = r"D:\00 興達計劃\05 EPC提供資料\HT"
Source_Company_TC = r"\\10.162.10.58\全處共用區\_Dwg\台中電廠燃氣機組更新計畫"
destny_Company_TC = r"D:\00 台中計劃\05 EPC提供資料\TC"
Source_Customized = ""
destny_Customized = ""

#輸入 計畫名稱、末4號碼
#輸出 計畫代號、完整文號
def Get_DocNo():
    # PlanNo :  1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑
    PlanNo = input("請輸入計畫\n興達請輸入\"1\" \n台中請輸入\"2\" \n自訂計畫請輸入\"99\":")
    DocNo = ""
    if PlanNo == "99" :
        DocNo = input("請輸入您的文號:")
    else:
        Doc_Num = str(input("請輸入號碼(1234):"))
        File_a = Plan_Dict.get(Easy_Dict.get(PlanNo, 'f'))
        DocNo = "-".join(File_a) + "-" + Doc_Num     
    return PlanNo, DocNo
#輸出 來源路徑、最終路徑
def Gen_path(PlanNo, DocNo):
    # PathWay : 1 = 雲端硬碟路徑  # 2 = 家庭硬碟路徑    # 3 = 公司硬碟路徑  # 0 = 指定路徑
    # PlanNo :  1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑
    path_map = {# "代碼組合" : (source_folder, dest_folder)
        "11":(Cloud_HT, Cloud_HT),
        "12":(Cloud_TC, Cloud_TC),
        "21":(Home_HT, Home_HT),
        "22":(Home_TC, Home_TC),
        "31":(Source_Company_HT, destny_Company_HT),
        "32":(Source_Company_TC, destny_Company_TC),
        "0":("",""),}
    PathWay = input("請輸入使用路徑(0 = 自訂, 1 = 雲端, 2 = 家庭, 3 = 公司 ：)")
    print("PathWay:", PathWay)
    source_folder = ""
    dest_folder = ""
    if PathWay == "0":
        source_folder = input("請輸入您的來源路徑 :") + "\\" + DocNo
        dest_folder = input("請輸入您的目的路徑 :") + "\\new" +DocNo
    else:
        source_folder, dest_folder = path_map[PathWay + PlanNo][0] + "\\" + DocNo, path_map[PathWay + PlanNo][1] + "\\" + Easy_Dict.get(PlanNo, 'f') +"_" +DocNo
    return source_folder, dest_folder
#輸出 最後資料夾路徑
def move_docutment(source_folder, dest_folder):
    #確認是否存在Source_folder
    if os.path.exists(source_folder):
        print("已找到指定資料夾")
        #若存在dest_folder，則刪除
        if os.path.exists(dest_folder):
            if int(input("資料夾已存在，請選擇保留 0 或選擇 重新複製 1:")):
                print("您選擇1，重新複製資料夾")
                shutil.rmtree(dest_folder)
                shutil.copytree(source_folder, dest_folder)
            else:
                print("您選擇0，保留原本資料夾，不複製")
        else:
            shutil.copytree(source_folder, dest_folder)
    else:
        print("查「無」指定資料夾")
    return dest_folder
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
                        DrawingNo_value = worksheet["E"+str(2+i)].value
                        Drawing_Title_value = worksheet["F"+str(2+i)].value
                        DrawinVision_value = worksheet["G"+str(2+i)].value
                        Letter_Num_value = worksheet["B2"].value
                        Letter_Date_value = worksheet["L2"].value
                        Letter_titl_value = worksheet["O2"].value
                        
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

    #輸出Excel檔案
    basename  = os.path.basename(file_path)
    filename_without_extension = os.path.splitext(basename)[0]
    NewExcelfile = "Done_" +  filename_without_extension + ".xlsx"
    Output_path = os.path.join(folder_path, NewExcelfile)
    #輸出成"Done_HT-D1-CTC-GEL-23-1171.xlsx"(範例)
    df.to_excel(Output_path, index=False)
    return Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value
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

    #輸出Excel檔案
    basename  = os.path.basename(file_path)
    filename_without_extension = os.path.splitext(basename)[0]
    NewExcelfile = "Done_" +  filename_without_extension + ".xlsx"
    Output_path = os.path.join(folder_path, NewExcelfile)
    #輸出成"Done_HT-D1-CTC-GEL-23-1171.xlsx"(範例)
    df.to_excel(Output_path, index=False)
    return Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value

def Convert_xlsb(File_folder):
    for file in os.listdir(File_folder):
        if fnmatch.fnmatch(file, '*.xlsb'):
            File_Path =os.path.join(File_folder, file)
            print("找到檔案：", File_Path)
    #顯示表格
    df = pd.read_excel(File_Path, sheet_name='Data1', engine='pyxlsb')
    Xlsx_File_Path = os.path.splitext(File_Path)[0] + "_converted.xlsx"
    #輸出檔案
    df.to_excel(Xlsx_File_Path)
    print("輸出檔案：", Xlsx_File_Path)
    print("return ", File_folder)
    input("暫停一下")
    return File_folder

def Text_Generator(Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value, PlanNo):
    print(PlanNo)
    Plan_No = int(PlanNo)
    Doc_content = Letter_titl_value
    Doc_Vision = DrawinVision_value
    Doc_No = Letter_Num_value
    Doc_Date = Letter_Date_value

    # PlanNo :  1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑
    My_company = ["", "南部施工處", "中部施工處", "興達發電廠", "台中發電廠"]
    Consult_Company = ["", "吉興公司", "泰興公司", "GE/CTCI"]
    Pages = str(input("請輸入審查意見頁數:"))

    date_obj = datetime.datetime.strptime(Doc_Date, "%Y/%m/%d")
    month = date_obj.strftime("%m")
    day = date_obj.strftime("%d")
    Plan_Name = "興達" if "CTC" in Doc_No else "台中"

    contents0 = "本文係" + My_company[Plan_No] + "對統包商提送「" + Doc_content + "」" + "Rev." + str(Doc_Vision) +"所提審查意見(共" + Pages + "頁)" + "，未逾合約規範，已電傳" + Consult_Company[Plan_No] + "，擬陳閱後文存。"
    contents1 = "檢送" + Plan_Name + "電廠燃氣機組更新改建計畫" + Doc_content + "Rev." + str(Doc_Vision) + "，" + My_company[Plan_No] + "之審查意見（如附，共" + Pages + "頁）供卓參，請查照。"
    contents2 = "依據GE/CTCI 112年" + month + "月" + day +"日" + Doc_No + "號辦理。"
    contents4 = "本文係統包商提送「" + Doc_content + "」" + "Rev." + str(Doc_Vision) +"，本組無意見，已Email通知" + Consult_Company[Plan_No] + "公司" + "，擬陳閱後文存。"
    print("----------------他單位審查意見簽辦------------------")
    print(contents0)
    print("----------------傳真------------------")
    print(contents1)
    #print(contents2)
    print("----------------主辦簽辦------------------")
    #print(contents4)
    input("暫停")
    return 0

def main():
    PlanNo, DocNo = Get_DocNo()
    print("PlanNo:", PlanNo)
    print("DocNo:", DocNo)
    print(r"----------------------Get_DocNo Done!-------------------------------------------")
    source_folder, dest_folder= Gen_path(PlanNo, DocNo)
    print("source_folder:", source_folder)
    print("dest_folder:", dest_folder) 
    print(r"----------------------Gen_path Done!--------------------------------------------")
    dest_folder = move_docutment(source_folder, dest_folder)
    print("dest_folder:",dest_folder)
    print(r"----------------------move_docutment Done!--------------------------------------")
    #開啟複製好的資料價
    os.startfile(dest_folder)
    if "TC(C0)" in dest_folder:
        Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value = Read_TC_Excel(dest_folder)
    else:
        #若為興達計畫，則額外執行.xlsb轉換.xlsx動作
        Xlsx_File_Path = Convert_xlsb(dest_folder)
        print(r"----------------------Convert_xlsb Done!--------------------------------------")
        Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value = Read_HT_Excel(dest_folder)
    print("來文名稱：", Letter_titl_value, "\n版次：" , DrawinVision_value, "\n來文號碼：", Letter_Num_value,"\n來文日期：", Letter_Date_value )
    print(r"----------------------Read_Excel Done!------------------------------------------")   
    Text_Generator(Letter_titl_value, DrawinVision_value, Letter_Num_value, Letter_Date_value, PlanNo)
    print(r"----------------------Text_Generator Done!------------------------------------------") 
    input("Press enter to exit...")

if __name__ == '__main__':
    main()
