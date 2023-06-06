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

Easy_Dict = {"1":"HT", "2":"TC", "99":""}
Plan_Dict = {"HT" : ["HT", "D1", "CTC", "GEL", "23"],
                "TC" : ["TPC", "TC(C0)", "CD", "23"]
                }
Cloud_HT = r"/content/sample_data/HT"
Cloud_TC = r"/content/sample_data/TC"
Home_HT = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT"
Home_TC = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC"
#來源資料夾、目標資料夾
Source_Company_HT = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫"
destny_Company_HT = r"D:\00 興達計劃\05 EPC提供資料\HT"
Source_Company_TC = r"\\\10.162.10.58\全處共用區\_Dwg\台中發電廠新建燃氣機組計畫"
destny_Company_TC = r"D:\00 台中計劃\05 EPC提供資料\TC"
Source_Customized = ""
destny_Customized = ""
PlanNo = ""#計畫代碼
DocNo = ""#廠商文號
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

def Text_Generator(Letter_Title, Drawing_Vision, Letter_Num, Letter_Date, PlanNo):

    Plan_No = int(PlanNo)
    judge00 = input("是否有意見，有請輸入1，無請輸入0：")
    if int(judge00):
        My_company_Num = str(input("請輸入審查意見填寫單位:\n 1 = 南部施工處 \n 2 = 中部施工處 \n 3 = 興達發電廠 \n 4 =台中發電廠"))
        Pages = str(input("請輸入審查意見頁數:"))
        My_company = ["", "南部施工處", "中部施工處", "興達發電廠", "台中發電廠"]
        Consult_Company = ["", "吉興公司", "泰興公司", "GE/CTCI"]
        
        #複製套印文件
        Print_page_File = dest_folder + r"\套印_" + Letter_Num + ".rtf"
        print(Pring_page_StandardFile)
        print(dest_folder)
        print(Print_page_File)
        input("等待一下")
        shutil.copy(Pring_page_StandardFile, Print_page_File)
        #複製傳真文件
        Fax_PageFile = dest_folder + r"\Fax_" + Letter_Num + ".doc"
        shutil.copy(Source_Fax_PageFile, Fax_PageFile)
        
        date_obj = datetime.datetime.strptime(Letter_Date, "%Y/%m/%d")
        month = date_obj.strftime("%m")
        day = date_obj.strftime("%d")
        Plan_Name = "興達" if "CTC" in Letter_Num else "台中"

        contents0 = "本文係" + My_company[My_company_Num] + "對統包商提送「" + Letter_Title + "」" + "Rev." + str(Drawing_Vision) +"所提審查意見(共" + Pages + "頁)" + "，未逾合約規範，已電傳" + Consult_Company[Plan_No] + "，擬陳閱後文存。"
        contents1 = "檢送" + Plan_Name + "電廠燃氣機組更新改建計畫" + Letter_Title + "Rev." + str(Drawing_Vision) + "，" + My_company[My_company_Num] + "之審查意見（如附，共" + Pages + "頁）供卓參，請查照。"
        contents2 = "依據GE/CTCI 112年" + month + "月" + day +"日" + Letter_Num + "號辦理。"
        contents4 = "本文係統包商提送「" + Letter_Title + "」" + "Rev." + str(Drawing_Vision) +"，本組無意見，已Email通知" + Consult_Company[Plan_No] + "公司" + "，擬陳閱後文存。"
        print("----------------他單位審查意見簽辦------------------")
        print(contents0)
        print("----------------傳真------------------")
        print(contents1)
        print(contents2)
        print("----------------主辦簽辦------------------")
        print(contents4)
        input("暫停")

    return 0

def main():
    while True:
        PlanNo, DocNo = Get_DocNo()
        print("PlanNo計畫代碼:", PlanNo)
        print("DocNo廠商文號:", DocNo)
        print(r"----------------------Get_DocNo Done!-------------------------------------------")

        source_folder, dest_folder= Gen_path(PlanNo, DocNo)
        print("source_folder資料來源資料夾:", source_folder)
        print("dest_folder存放路徑資料夾:", dest_folder) 
        print(r"----------------------Gen_path Done!--------------------------------------------")

        move_docutment(source_folder, dest_folder)
        print(r"----------------------move_docutment Done!--------------------------------------")
        #開啟複製好的資料價
        os.startfile(dest_folder)

        if "TC(C0)" in dest_folder:
            Letter_Title, Drawing_Vision, Letter_Num, Letter_Date = Read_TC_Excel(dest_folder)
        else:
            #若為興達計畫，則額外執行.xlsb轉換.xlsx動作
            Convert_xlsb(dest_folder)
            print(r"----------------------Convert_xlsb Done!-----------------------------------")
            Letter_Title, Drawing_Vision, Letter_Num, Letter_Date = Read_HT_Excel(dest_folder)
        print("來文名稱：", Letter_Title, "\n版次：" , Drawing_Vision, "\n來文號碼：", Letter_Num,"\n來文日期：", Letter_Date )
        print(r"----------------------Read_Excel Done!------------------------------------------")   

        Text_Generator(Letter_Title, Drawing_Vision, Letter_Num, Letter_Date, PlanNo)
        print(r"----------------------Text_Generator Done!------------------------------------------") 
        input("Press enter to exit...")

if __name__ == '__main__':
    main()
