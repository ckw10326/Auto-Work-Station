"""
引用constant.py內常數
輸出相關文字
"""
import shutil
import datetime
from constants import (HT_PRINT_STD_FILE11, HT_PRINT_STD_FILE12, HT_PRINT_STD_FILE10, TC_PRINT_STD_FILE21, 
                       TC_PRINT_STD_FILE22, TC_PRINT_STD_FILE20, HT_FAX_FILE11, HT_FAX_FILE12, HT_FAX_FILE10, 
                       TC_FAX_FILE21, TC_FAX_FILE22, TC_FAX_FILE20, HT_COMMENT_FILE11, HT_COMMENT_FILE12, 
                       HT_COMMENT_FILE10, TC_COMMENT_FILE21, TC_COMMENT_FILE22, TC_COMMENT_FILE20)

#text_gen(letter_title, letter_vision, letter_num, def_letter_date)
def text_gen(def_letter_title, def_letter_vision, def_letter_num, def_letter_date, dest_folder):
    "內容生成"
    #判斷計畫類別
    if "HT" in def_letter_num:
        plan_no = 1
    elif "TC" in def_letter_num:
        plan_no = 2
    else:
        plan_no = 3
    print("plan_no(1:HT, 2:TC, 3:other)：", plan_no)
    #判斷路徑為[0雲端 1家庭 2公司 ]
    if r"workspaces" in dest_folder:
        location = 0
    elif r"Test program" in  dest_folder:
        location = 1
    elif r"EPC提供資料" in dest_folder:
        location = 2
    else:
        print("路徑可能有點問題")
        return None
    #判斷套印檔案、傳真、comment
    #path_map[keys][0]
    keys = str(plan_no) + str(location)
    path_map = {"10":(HT_PRINT_STD_FILE10, HT_FAX_FILE10, HT_COMMENT_FILE10),
                "11":(HT_PRINT_STD_FILE11, HT_FAX_FILE11, HT_COMMENT_FILE11),
                "12":(HT_PRINT_STD_FILE12, HT_FAX_FILE12, HT_COMMENT_FILE12),
                "20":(TC_PRINT_STD_FILE20, TC_FAX_FILE20, TC_COMMENT_FILE20),
                "21":(TC_PRINT_STD_FILE21, TC_FAX_FILE21, TC_COMMENT_FILE21),
                "22":(TC_PRINT_STD_FILE22, TC_FAX_FILE22, TC_COMMENT_FILE22)}

    judge00 = input("你是否有意見?\n有請輸入1，無請輸入0：")
    if int(judge00):
        print("copy comment file")
    else:
        my_company_num = int(input("請輸入審查意見填寫單位:\n 1 = 南部施工處 \n 2 = 中部施工處 \n 3 = 興達發電廠 \n 4 =台中發電廠\n"))
        pages = str(input("請輸入審查意見頁數:"))
        company_name = ["", "南部施工處", "中部施工處", "興達發電廠", "台中發電廠"]
        consult_company = ["", "吉興公司", "泰興公司", "GE/CTCI"]
        #複製套印文件
        print_dest_file = dest_folder + r"\套印_" + def_letter_num + ".rtf"
        print("套印檔案:", path_map[keys][0])
        print("目標路徑:", dest_folder)
        print("輸出位置:", print_dest_file)
        input("plz enter any key")
        shutil.copy(path_map[keys][0], print_dest_file)
        #複製傳真文件
        fax_dest_file = dest_folder + r"\Fax_" + def_letter_num + ".doc"
        print("傳真檔案:", path_map[keys][1])
        print("目標路徑:", dest_folder)
        print("輸出位置:", fax_dest_file)
        input("plz enter any key")
        shutil.copy(path_map[keys][1], fax_dest_file)
        date_obj = datetime.datetime.strptime(def_letter_date, "%Y/%m/%d")
        month = date_obj.strftime("%m")
        day = date_obj.strftime("%d")
        plan_name = "興達" if "CTC" in def_letter_num else "台中"

        contents0 = ("本文係" + company_name[my_company_num] + "對統包商提送「" + def_letter_title
                    + "」" + "Rev." + str(def_letter_vision)  + "所提審查意見(共" + pages + "頁)"
                    + "，未逾合約規範，已電傳" + consult_company[plan_no] + "，擬陳閱後文存。")
        contents1 = ("檢送" + plan_name + "電廠燃氣機組更新改建計畫"  + def_letter_title + "Rev."
                    + str(def_letter_vision) + "，" + company_name[my_company_num]
                    + "之審查意見（如附，共" + pages + "頁）供卓參，請查照。")
        contents2 = "依據GE/CTCI 112年" + month + "月" + day +"日" + def_letter_num + "號辦理。"
        contents4 = ("本文係統包商提送「" + def_letter_title + "」" + "Rev." + str(def_letter_vision)
                    +"，本組無意見，已Email通知" + consult_company[plan_no] + "公司" + "，擬陳閱後文存。")
        print("----------------他單位審查意見簽辦------------------")
        print(contents0)
        print("----------------傳真------------------------------")
        print(contents1)
        print(contents2)
        print("----------------主辦簽辦--------------------------")
        print(contents4)
        input("enter any kesy to exit")
    return 0

def test2():
    "主程式"
    dest_folder = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT\HT_HT-D1-CTC-GEL-23-1188"
    text_gen('HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab',
             'A', 'TPC-TC(C0)-CD-23-0002', '2023/02/02', dest_folder)
    input("Press enter to exit...")
    return None

def test():
    "測試程式"
    return None

if __name__ == '__main__':
    test2()

