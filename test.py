
"""
主程式
前置作業
pip install pyxlsb
pip install openpyxl
"""
import os 
import shutil
import pandas as pd
import doc_collection
import convert_excel
import text_gen

PLAN_LIST = ["", "HT", "TC"]
LOCATION = ["", "家庭", "公司", "雲端"]
DOC_NO_STRUCTURE = {"HT" : ["HT", "D1", "CTC", "GEL", "23"],
                "TC" : ["TPC", "TC(C0)", "CD", "23"],
                "HT2" : ["HT", "D1", "GEI", "GEL", "23"]
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
#套印文件 [1興達2台中]+[1家庭 2公司 3雲端]
HT_PRINT_STD_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT套印標準.rtf"
HT_PRINT_STD_FILE12 = r"D:\00 興達計劃\HT套印標準.rtf"
HT_PRINT_STD_FILE13 = r"ckw10326/Auto-Work-Station/Files/HT套印標準.doc"
TC_PRINT_STD_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC套印標準.rtf"
TC_PRINT_STD_FILE22 = r"D:\00 台中計畫\TC套印標準.rtf"
TC_PRINT_STD_FILE23 = r"ckw10326/Auto-Work-Station/Files/TC套印標準.rtf"
#傳真文件
HT_FAX_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT 傳真 sample.doc"
HT_FAX_FILE12 = r"D:\00 興達計劃\HT 傳真 sample.doc"
HT_FAX_FILE13 = r"ckw10326/Auto-Work-Station/Files/HT 傳真 sample.doc"
TC_FAX_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC 傳真 sample.doc"
TC_FAX_FILE22 = r"D:\00 台中計畫\TC 傳真 sample.doc"
TC_FAX_FILE23 = r"ckw10326/Auto-Work-Station/Files/TC 傳真 sample.doc"
#審查意見
HT_COMMENT_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT 審查意見 sample.doc"
HT_COMMENT_FILE12 = r"D:\00 興達計劃\HT 審查意見 sample.doc"
HT_COMMENT_FILE13 = r"ckw10326/Auto-Work-Station/Files/HT 審查意見 sample.doc"
TC_COMMENT_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC 審查意見 sample.doc"
TC_COMMENT_FILE22 = r"D:\00 台中計畫\TC 審查意見 sample.doc"
TC_COMMENT_FILE23 = r"ckw10326/Auto-Work-Station/Files/TC 審查意見 sample.doc"
#variable
source_customized = ""
destiny_customized = ""
plan_no = ""
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

'''
def main():
    while True:
        plan_no, location_no, doc_no = Get_DocNo()

        source_folder, dest_folder= Gen_path(PlanNo, DocNo)

        move_docutment(source_folder, dest_folder)

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
'''
        
def main2():
    doc_collection.file_path_process()

if __name__ == '__main__':
    main2()
