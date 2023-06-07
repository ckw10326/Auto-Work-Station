'''
pyinstaller -F main.py
'''
import sys
import os
import datetime
import shutil
import fnmatch
import openpyxl
import pandas as pd

sys.path.append(r'C:/Users/OXO/OneDrive/01 Book/00 Test program/Auto-Work-Station')

PLAN_LIST = ["", "HT", "TC"]
LOCATION = ["", "家庭", "公司", "雲端"]
DOC_NO_STRUCTURE = {"HT" : ["HT", "D1", "CTC", "GEL", "23"],
                "TC" : ["TPC", "TC(C0)", "CD", "23"],
                "HT2" : ["HT", "D1", "GEI", "GEL", "23"]
                }
#CONSTANT
CLOUD_HT = r"/content/sample_data/HT"
CLOUD_TC = r"/content/sample_data/TC"
HOME_HT = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT"
HOME_TC = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC"
SOURCE_HT = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫"
DESTINY_HT = r"D:\00 興達計劃\05 EPC提供資料\HT"
SOURCE_TC = r"\\\10.162.10.58\全處共用區\_Dwg\台中發電廠新建燃氣機組計畫"
DESTINY_TC = r"D:\00 台中計劃\05 EPC提供資料\TC"
#套印文件 [1興達2台中]+[1家庭 2公司 3雲端]
PRINT_STD_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT套印標準.rtf"
PRINT_STD_FILE12 = r"D:\00 興達計劃\HT套印標準.rtf"
PRINT_STD_FILE13 = r"ckw10326/Auto-Work-Station/Files/HT套印標準.doc"
PRINT_STD_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC套印標準.rtf"
PRINT_STD_FILE22 = r"D:\00 台中計畫\TC套印標準.rtf"
PRINT_STD_FILE23 = r"ckw10326/Auto-Work-Station/Files/TC套印標準.rtf"
#傳真文件
FAX_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT 傳真 sample.doc"
FAX_FILE12 = r"D:\00 興達計劃\HT 傳真 sample.doc"
FAX_FILE13 = r"ckw10326/Auto-Work-Station/Files/HT 傳真 sample.doc"
FAX_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC 傳真 sample.doc"
FAX_FILE22 = r"D:\00 台中計畫\TC 傳真 sample.doc"
FAX_FILE23 = r"ckw10326/Auto-Work-Station/Files/TC 傳真 sample.doc"
#審查意見
COMMENT_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT 審查意見 sample.doc"
COMMENT_FILE12 = r"D:\00 興達計劃\HT 審查意見 sample.doc"
COMMENT_FILE13 = r"ckw10326/Auto-Work-Station/Files/HT 審查意見 sample.doc"
COMMENT_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC 審查意見 sample.doc"
COMMENT_FILE22 = r"D:\00 台中計畫\TC 審查意見 sample.doc"
COMMENT_FILE23 = r"ckw10326/Auto-Work-Station/Files/TC 審查意見 sample.doc"
#variable
source_customized = ""
destiny_customized = ""
plan_no = ""
doc_path = ""
location_no = ""
doc_no = ""
source_folder = ""#資料來源資料夾
dest_folder = ""#存放路徑資料夾
converted_xlsx_path = ""#轉檔路徑
#來文資訊
letter_title = ""#來文資訊
letter_vision = ""#來文資訊
letter_num = ""#來文資訊
letter_date = ""#來文資訊


def file_path_process():
    "收集關鍵字，生成文號，生成路徑"
    file_plan_no = ""
    file_doc_num = ""
    file_plan_no = input("請輸入計畫\n興達請輸入\"1\" \n台中請輸入\"2\" \n自訂計畫請輸入\"99\":")
    if file_plan_no == "99" :
        pass
    else:
        #檢查輸入是否為4位數字
        check_num = True
        while check_num:
            doc_end_num = str(input("請輸入號碼(1234):"))
            if len(doc_end_num) == 4 or len(doc_end_num) == 6 :
                check_num = False
            else:
                print("請輸入4位數字")
        ana_num = DOC_NO_STRUCTURE.get(PLAN_LIST[int(file_plan_no)])
        #pre_path = ["HT", "D1", "CTC", "GEL", "23"]
        file_doc_num = "-".join(ana_num) + "-" + doc_end_num
        #file_doc_num = HT-D1-CTC-GEL-23 + "-1234"
    """
    pathway : #1 = 雲端硬碟路徑, #2 = 家庭硬碟路徑, #3 = 公司硬碟路徑, #0 = 指定路徑\
    plan_no : 1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑\
    代碼組合 : (source_folder, dest_folder)
    """
    path_map = {"11":(CLOUD_HT, CLOUD_HT), "12":(CLOUD_TC, CLOUD_TC),
                "21":(HOME_HT, HOME_HT), "22":(HOME_TC, HOME_TC),
                "31":(SOURCE_HT, DESTINY_HT), "32":(SOURCE_TC, DESTINY_TC),
                "0":("","")
                }
    path_way = input("請輸入使用路徑(0 = 自訂, 1 = 雲端, 2 = 家庭, 3 = 公司 ：)")
    print("pathWay:", path_way)
    file_source_folder = ""
    file_dest_folder = ""
    if path_way == "0":
        file_source_folder = input("請輸入您的來源路徑 :") + "\\" + file_doc_num
        file_dest_folder = input("請輸入您的目的路徑 :") + "\\The_" +file_doc_num
    else:
        file_source_folder = path_map[path_way + file_plan_no][0] + "\\" + file_doc_num
        file_dest_folder = (path_map[path_way + file_plan_no][1] + "\\" +
                           PLAN_LIST[int(file_plan_no)] +"_" +file_doc_num)
    print("來源路徑:", file_source_folder)
    print("目的路徑:", file_dest_folder)
    print("計畫代碼:", file_plan_no)
    print("文件號碼:", file_doc_num)
    print(r"----------------------file_path_process Done!--------------------------------")
    return file_doc_num, file_source_folder, file_dest_folder

def move_docutment(def_source_folder, def_dest_folder):
    "移動資料夾"
    #檢查來源資料夾是否存在
    if os.path.exists(def_source_folder):
        print("已找到指定資料夾：" + def_source_folder)
        #若存在dest_folder，則刪除
        if os.path.exists(def_dest_folder):
            if int(input("資料夾已存在，請選擇保留 0 或選擇 重新複製 1:")):
                print("您選擇1，重新複製資料夾")
                shutil.rmtree(def_dest_folder)
                shutil.copytree(def_source_folder, def_dest_folder)
                print(def_dest_folder,"複製完成")
            else:
                print("您選擇0，保留原本資料夾，不複製")
        else:
            shutil.copytree(def_source_folder, def_dest_folder)
            #開啟複製好的資料價
    else:
        print("查「無」指定資料夾")
        return None
    print(r"----------------------move_docutment Done!--------------------------------------")
    input("按任意鍵結束")
    return None

def text_gen(def_l_title, def_l_vision, def_l_Num, Letter_Date):
    if "HT" in def_l_Num:
        Plan_No = 1
    elif "TC" in def_l_Num:
        Plan_No = 2
    else:
        Plan_No = 3
    print("Plan_No：", Plan_No)
    judge00 = input("是否有意見，有請輸入1，無請輸入0：")
    if int(judge00):
        My_company_Num = int(input("請輸入審查意見填寫單位:\n 1 = 南部施工處 \n 2 = 中部施工處 \n 3 = 興達發電廠 \n 4 =台中發電廠\n"))
        Pages = str(input("請輸入審查意見頁數:"))
        My_company = ["", "南部施工處", "中部施工處", "興達發電廠", "台中發電廠"]
        Consult_Company = ["", "吉興公司", "泰興公司", "GE/CTCI"]
        
        date_obj = datetime.datetime.strptime(Letter_Date, "%Y/%m/%d")
        month = date_obj.strftime("%m")
        day = date_obj.strftime("%d")
        Plan_Name = "興達" if "CTC" in def_l_Num else "台中"

        contents0 = "本文係" + My_company[My_company_Num] + "對統包商提送「" + def_l_title + "」" + "Rev." + str(def_l_vision) +"所提審查意見(共" + Pages + "頁)" + "，未逾合約規範，已電傳" + Consult_Company[Plan_No] + "，擬陳閱後文存。"
        contents1 = "檢送" + Plan_Name + "電廠燃氣機組更新改建計畫" + def_l_title + "Rev." + str(def_l_vision) + "，" + My_company[My_company_Num] + "之審查意見（如附，共" + Pages + "頁）供卓參，請查照。"
        contents2 = "依據GE/CTCI 112年" + month + "月" + day +"日" + def_l_Num + "號辦理。"
        contents4 = "本文係統包商提送「" + def_l_title + "」" + "Rev." + str(def_l_vision) +"，本組無意見，已Email通知" + Consult_Company[Plan_No] + "公司" + "，擬陳閱後文存。"
        print("----------------他單位審查意見簽辦------------------")
        print(contents0)
        print("----------------傳真------------------")
        print(contents1)
        print(contents2)
        print("----------------主辦簽辦------------------")
        print(contents4)
        input("enter to exit")

        '''
        #複製套印文件
        print_judge = input("是否需印套印文件，是請輸入1，否請輸入0：")
        if int(print_judge):
            local_judge = input("請輸入特徵碼\n [1興達2台中]+[1家庭 2公司 3雲端]")
            print("local_judge:", local_judge)
            print_path = "PRINT_STD_FILE" + str(local_judge)
            #會跑出PRINT_STD_FILE11
            shutil.copytree(print_path, def_dest_folder)
        else:
            print("不印套印文件")
            return 0
        Print_page_File = dest_folder + r"\套印_" + def_l_Num + ".rtf"
        print(Pring_page_StandardFile)
        print(dest_folder)
        print(Print_page_File)
        input("plz enter any key")
        shutil.copy(Pring_page_StandardFile, Print_page_File)
        #複製傳真文件
        Fax_PageFile = dest_folder + r"\Fax_" + def_l_Num + ".doc"
        shutil.copy(Source_Fax_PageFile, Fax_PageFile)
        '''
    return 0

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
            #比對副檔名條件
            if file.endswith(file_type) * file.count("-") > 3:
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

def convert_xlsb(folder_path):
    xlsx_path = ""
    "尋找檔案是否存在"
    if os.path.exists(folder_path):
        print("找到目標資料夾：", folder_path)
        for file in os.listdir(folder_path):
            #確認xlsb檔案是否存在
            if fnmatch.fnmatch(file, '*.xlsb'):
                file_path =os.path.join(folder_path, file)
                print("找到目標檔案路徑：", file_path)
                #抓到隱藏檔案(檔名有關鍵字：~$H，不動作
                if "~$" in file_path:
                    print(file_path,"抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
                    return None
                else:
                    xlsx_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
                    print("輸出檔案：", xlsx_path)
                    if os.path.exists(xlsx_path):
                        print("_converted.xlsx" + "檔案已存在，不動作")
                        return None
                    else:
                        data_frame = pd.read_excel(file_path, sheet_name='Data1', engine='pyxlsb')
                        data_frame.to_excel(xlsx_path)
                        print(r"----------------------Convert_xlsb Done!------------------------")
                        input("enter any keys to exit")
                        return xlsx_path
            else:
                print("找不到目標檔案：", file)
                return None
    else:
        "若找不到資料夾，則回傳None"
        print("找不到目標資料夾：", folder_path)
        return None
    
def main():
    "主程式"
    while True:
        global letter_num, source_folder, dest_folder
        global letter_title, letter_vision, letter_num, letter_date
        letter_num, source_folder, dest_folder = file_path_process()
        print("1",letter_num)
        print("2",source_folder)
        print("3",dest_folder)
        move_docutment(source_folder, dest_folder)
        os.startfile(dest_folder)
        if "TC(C0)" in dest_folder:
            print("TC(C0)")
            letter_title, letter_vision, letter_num, letter_date = read_tc_excel(dest_folder)
        elif "HT" in dest_folder:
            converted_xlsx_path = convert_xlsb(dest_folder)
            letter_title, letter_vision, letter_num, letter_date = read_ctc_ht_excel(dest_folder)
        else:
            return None

        text_gen(letter_title, letter_vision, letter_num, letter_date)
        input("Press enter to exit...")

def update_value(new_value1, new_value2, new_value3):
    "更新變數值"
    #global HT_path
    #HT_path = new_value
    return None

def main2():
    """[summary]"""
    return None

if __name__ == '__main__':
    main()
