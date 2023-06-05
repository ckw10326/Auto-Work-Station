"""
請說明內容
"""
import os 
import shutil
import pandas as pd

PLAN_LIST = ["", "HT", "TC"]
DOC_NO_STRUCTURE = {"HT" : ["HT", "D1", "CTC", "GEL", "23"],
                "TC" : ["TPC", "TC(C0)", "CD", "23"]
                }
#CONSTANT
CLOUD_HT = r"/content/sample_data/HT"
CLOUD_HT = r"/content/sample_data/TC"
HOME_HT = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT"
HOME_TC = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC"
SOURCE_HT = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫"
DESTINY_HT = r"D:\00 興達計劃\05 EPC提供資料\HT"
SOURCE_TC = r"\\\10.162.10.58\全處共用區\_Dwg\台中發電廠新建燃氣機組計畫"
DESTINY_TC = r"D:\00 台中計劃\05 EPC提供資料\TC"
#variable
source_customized = ""
destiny_customized = ""
#來文資料
plan_no = ""
doc_no = ""
source_folder = ""#資料來源資料夾
dest_folder = ""#存放路徑資料夾

Xlsb_to_xlsx_Path = ""#轉檔路徑
Letter_Title = ""#來文資訊
Drawing_Vision = ""#來文資訊
Letter_Num = ""#來文資訊
Letter_Date = ""#來文資訊
#Pring_page_StandardFile = r"D:\00 興達計劃\HT套印標準.rtf"#套印檔案路徑
Pring_page_StandardFile = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT套印標準.rtf"
Print_page_File = ""
Source_Fax_PageFile = r"D:\00 台中計畫\05 EPC提供資料\台中傳真 -Sample.doc"
Source_Comment_PageFile = r"D:\00 台中計畫\05 EPC提供資料\台中意見 -Sample.doc"

#輸入 計畫名稱、末4號碼
#輸出 計畫代號、完整文號
def file_path_process():
    "收集關鍵字，生成文號"
    # file_plan_no :  1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑
    file_plan_no = ""
    file_doc_path = ""
    patha = ""
    file_plan_no = input("請輸入計畫\n興達請輸入\"1\" \n台中請輸入\"2\" \n自訂計畫請輸入\"99\":")
    if file_plan_no == "99" :
        file_doc_path = input("請輸入您的路徑:")
    else:
        doc_end_num = str(input("請輸入號碼(1234):"))
        patha = DOC_NO_STRUCTURE.get(PLAN_LIST[int(file_plan_no)])
        file_doc_path = "-".join(patha) + "-" + doc_end_num    
    print("計畫代碼:", file_plan_no)
    print("廠商文號路徑:", file_doc_path)
    print(r"----------------------file_path_process Done!-------------------------------------------")
    return file_plan_no, file_doc_path

def gen_path(plan_no, doc_no):
    "輸入文件相關資料，生成路徑"
    # PathWay : 1 = 雲端硬碟路徑  # 2 = 家庭硬碟路徑    # 3 = 公司硬碟路徑  # 0 = 指定路徑
    # plan_no :  1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑
    path_map = {# "代碼組合" : (source_folder, dest_folder)
        "11":(CLOUD_HT, CLOUD_HT),
        "12":(CLOUD_HT, CLOUD_HT),
        "21":(HOME_HT, HOME_HT),
        "22":(HOME_TC, HOME_TC),
        "31":(SOURCE_HT, DESTINY_HT),
        "32":(SOURCE_TC, DESTINY_TC),
        "0":("",""),}
    PathWay = input("請輸入使用路徑(0 = 自訂, 1 = 雲端, 2 = 家庭, 3 = 公司 ：)")
    print("PathWay:", PathWay)
    source_folder = ""
    dest_folder = ""
    if PathWay == "0":
        source_folder = input("請輸入您的來源路徑 :") + "\\" + doc_no
        dest_folder = input("請輸入您的目的路徑 :") + "\\new" +doc_no
    else:
        source_folder, dest_folder = path_map[PathWay + plan_no][0] + "\\" + doc_no, path_map[PathWay + plan_no][1] + "\\" + PLAN_LIST.get(plan_no, 'f') +"_" +doc_no
    return source_folder, dest_folder
#輸出 最後資料夾路徑
def move_docutment(source_folder, dest_folder):
    "移動資料夾"
    #確認是否存在Source_folder
    if os.path.exists(source_folder):
        print("已找到指定資料夾")
        #若存在dest_folder，則刪除
        if os.path.exists(dest_folder):
            if int(input("資料夾已存在，請選擇保留 0 或選擇 重新複製 1:")):
                print("您選擇1，重新複製資料夾")
                shutil.rmtree(dest_folder)
                shutil.copytree(source_folder, dest_folder)
                print("複製完成")
            else:
                print("您選擇0，保留原本資料夾，不複製")
        else:
            shutil.copytree(source_folder, dest_folder)
    else:
        print("查「無」指定資料夾")
    return None

    for file in os.listdir(File_folder):
        if fnmatch.fnmatch(file, '*.xlsb'):
            File_Path =os.path.join(File_folder, file)
            print("找到檔案：", File_Path)
    #顯示表格
    df = pd.read_excel(File_Path, sheet_name='Data1', engine='pyxlsb')
    Xlsb_to_xlsx_Path = os.path.splitext(File_Path)[0] + "_converted.xlsx"
    #輸出檔案
    df.to_excel(Xlsb_to_xlsx_Path)
    print("輸出檔案：", Xlsb_to_xlsx_Path)
    return None

def main():
    "主要執行內容"
    while True:
        global plan_no
        global doc_no
        plan_no, doc_no = file_path_process()
        # print("plan_no計畫代碼:", plan_no)
        # print("doc_no廠商文號:", doc_no)
        # print(r"----------------------file_path_process Done!-------------------------------------------")

        # source_folder, dest_folder= gen_path(plan_no, doc_no)
        # print("source_folder資料來源資料夾:", source_folder)
        # print("dest_folder存放路徑資料夾:", dest_folder) 
        # print(r"----------------------gen_path Done!--------------------------------------------")

        # move_docutment(source_folder, dest_folder)
        # print(r"----------------------move_docutment Done!--------------------------------------")
        # #開啟複製好的資料價
        # os.startfile(dest_folder)
        # print(r"----------------------Text_Generator Done!------------------------------------------") 
        # input("Press enter to exit...")

if __name__ == '__main__':
    main()
