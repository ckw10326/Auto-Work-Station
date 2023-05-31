#簡化修正版
#cd C:\Users\OXO\OneDrive\01 Book\00 Test program
#pyinstaller -F Pack.py
import shutil
import os #用來避免檔案已經存在的異常
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
    print("PlanNo:", PlanNo)
    print("DocNo:", DocNo)
    print("Get_DocNo() Done!")
    print("-------------------------------------------------------------")         
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
    print("source_folder:", source_folder)
    print("dest_folder:", dest_folder)
    print("Gen_path(PlanNo, DocNo), done!")
    print("-------------------------------------------------------------")   
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
    print("move_docutment(source_folder, dest_folder), done!")
    print("dest_folder:",dest_folder)
    print("-------------------------------------------------------------")   
    return dest_folder

PlanNo, DocNo = Get_DocNo()
source_folder, dest_folder= Gen_path(PlanNo, DocNo)
dest_folder = move_docutment(source_folder, dest_folder)
print("output")
print(dest_folder)
os.startfile(dest_folder)
input("Press enter to exit...")

# PathWay : 1 = 雲端硬碟路徑  # 2 = 家庭硬碟路徑    # 3 = 公司硬碟路徑  # 0 = 指定路徑
# PlanNo :  1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑