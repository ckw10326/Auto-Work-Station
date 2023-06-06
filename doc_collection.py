"""
Doc collection
"""
import os
import shutil
import sys
from test import (CLOUD_HT, CLOUD_TC, DESTINY_HT, DESTINY_TC, DOC_NO_STRUCTURE,
                  HOME_HT, HOME_TC, PLAN_LIST, SOURCE_HT, SOURCE_TC)
sys.path.append(r'C:/Users/OXO/OneDrive/01 Book/00 Test program/Auto-Work-Station')
#輸入 計畫名稱、末4號碼
#輸出 計畫代號、完整文號
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

def main():
    "主要執行內容"
    return None

def test1():
    "測試用"
    return None

if __name__ == '__main__':
    test1()
