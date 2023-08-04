"""
引用constant.py內常數
輸出相關文字
"""
import shutil
import datetime
import os
import constants

def text_gen(def_letter_title, def_letter_vision, def_letter_num, def_letter_date, dest_folder):
    """內容生成"""
    # 判斷計畫類別
    if "HT" in def_letter_num:
        plan_no = 1
    elif "TC" in def_letter_num:
        plan_no = 2
    else:
        plan_no = 3
    print("plan_no(1:HT, 2:TC, 3:other)：", plan_no)
    # 判斷路徑為[0雲端 1家庭 2公司 ]
    if r"workspaces" in dest_folder:
        location = 0
    elif r"Test program" in dest_folder:
        location = 1
    elif r"EPC提供資料" in dest_folder:
        location = 2
    else:
        print("路徑可能有點問題")
        return None
    # 判斷套印檔案、傳真、comment
    path_map = {"10": (constants.HT_PRINT_STD_FILE10, constants.HT_FAX_FILE10, constants.HT_COMMENT_FILE10),
                "11": (constants.HT_PRINT_STD_FILE11, constants.HT_FAX_FILE11, constants.HT_COMMENT_FILE11),
                "12": (constants.HT_PRINT_STD_FILE12, constants.HT_FAX_FILE12, constants.HT_COMMENT_FILE12),
                "20": (constants.TC_PRINT_STD_FILE20, constants.TC_FAX_FILE20, constants.TC_COMMENT_FILE20),
                "21": (constants.TC_PRINT_STD_FILE21, constants.TC_FAX_FILE21, constants.TC_COMMENT_FILE21),
                "22": (constants.TC_PRINT_STD_FILE22, constants.TC_FAX_FILE22, constants.TC_COMMENT_FILE22)}

    judge00 = input("你是否有意見?\n有請輸入1，無請輸入0：")
    if int(judge00):
        print("copy comment file")
    else:
        my_company_num = int(
            input("請輸入審查意見填寫單位:\n 1 = 南部施工處 \n 2 = 中部施工處 \n 3 = 興達發電廠 \n 4 =台中發電廠\n"))
        pages = str(input("請輸入審查意見頁數:"))
        company_name = ["", "南部施工處", "中部施工處", "興達發電廠", "台中發電廠"]
        consult_company = ["", "吉興公司", "泰興公司", "GE/CTCI"]

        date_obj = datetime.datetime.strptime(def_letter_date, "%Y/%m/%d")
        month = date_obj.strftime("%m")
        day = date_obj.strftime("%d")
        plan_name = "興達" if "CTC" in def_letter_num else "台中"

        contents0 = ("本文係" + company_name[my_company_num] + "對統包商提送「" + def_letter_title
                     + "」" + "Rev." + str(def_letter_vision) +
                     "所提審查意見(共" + pages + "頁)"
                     + "，未逾合約規範，已電傳" + consult_company[plan_no] + "，擬陳閱後文存。")
        contents1 = ("檢送" + plan_name + "電廠燃氣機組更新改建計畫「" + def_letter_title + "」，Rev."
                     + str(def_letter_vision) + "，" +
                     company_name[my_company_num]
                     + "之審查意見（如附，共" + pages + "頁）供卓參，請查照。")
        contents2 = "依據GE/CTCI 112年" + month + "月" + day + "日" + def_letter_num + "號辦理。"
        contents4 = ("本文係統包商提送「" + def_letter_title + "」" + "Rev." + str(def_letter_vision)
                     + "，本組無意見，已Email通知" + consult_company[plan_no] + "公司" + "，擬陳閱後文存。")
        print("----------------套印內容------------------")
        print(contents0)
        print("----------------傳真------------------------------")
        print(contents1)
        print("----------------依據------------------------------")
        print(contents2)
        print("----------------主辦簽辦--------------------------")
        print(contents4)
        input("enter any kesy to exit")
    return 0

def copy_plan_file(filepath):
    """解析路徑"""
    file_folder, file_allname = os.path.split(filepath)
    file_name = file_allname.split(".")[0]

    # 複製套印文件
    print_dest_file = file_folder + r"/套印_" + file_name + ".rtf"
    print("套印檔案:", file_allname)
    print("目標路徑:", file_folder)
    print("輸出套印檔案:", print_dest_file)
    input("plz enter any key")
    if "HT-" in filepath:
        print_file = constants.HT_PRINT_STD_FILE10
    elif "-TC" in filepath:
        print_file = constants.TC_PRINT_STD_FILE20
    else:
        pass

    if os.path.exists(print_file) and not os.path.exists(print_dest_file):
        # 進行文件複製
        shutil.copy(print_file, print_dest_file)
    else:
        # 源文件或目標文件夾不存在
        print("源文件或目標文件夾不存在!")

    # 複製傳真文件
    fax_dest_file = file_folder + r"/Fax_" + file_name + ".doc"
    print("傳真檔案:", fax_dest_file)
    print("目標路徑:", file_folder)
    print("輸出傳真檔案:", fax_dest_file)
    input("plz enter any key")

    if "HT-" in filepath:
        fax_file = constants.HT_FAX_FILE10
    elif "-TC" in filepath:
        fax_file = constants.TC_FAX_FILE20
    else:
        pass

    if os.path.exists(fax_file) and not os.path.exists(fax_dest_file):
        # 進行文件複製
        shutil.copy(fax_file, fax_dest_file)
    else:
        # 源文件或目標文件夾不存在
        print("源文件或目標文件夾不存在!")

def file_path_process():
    """收集關鍵字，生成文號，生成路徑"""
    file_plan_no = ""
    file_doc_num = ""
    file_plan_no = input("請輸入計畫\n興達請輸入\"1\" \n台中請輸入\"2\" \n自訂計畫請輸入\"99\":")
    if file_plan_no == "99":
        pass
    else:
        # 檢查輸入是否為4位數字
        check_num = True
        while check_num:
            doc_end_num = str(input("請輸入號碼(1234):"))
            if len(doc_end_num) == 4 or len(doc_end_num) == 6:
                check_num = False
            else:
                print("請輸入4位數字")
        ana_num = constants.DOC_NO_STRUCTURE.get(
            constants.PLAN_LIST[int(file_plan_no)])
        # pre_path = ["HT", "D1", "CTC", "GEL", "23"]
        file_doc_num = "-".join(ana_num) + "-" + doc_end_num
        # file_doc_num = HT-D1-CTC-GEL-23 + "-1234"

    # pathway : #1 = 雲端硬碟路徑, #2 = 家庭硬碟路徑, #3 = 公司硬碟路徑, #0 = 指定路徑\
    # plan_no : 1 = 興達計畫, 2 = 台中計畫, 99 = 指定路徑\
    # 代碼組合 : (source_folder, dest_folder)
    path_map = {
        "11": (constants.AZURE_SOURCE, constants.AZURE_SOURCE),
        "12": (constants.AZURE_SOURCE, constants.AZURE_SOURCE),
        "21": (constants.HOME_SOURCE_HT, constants.HOME_DESTINY_HT),
        "22": (constants.HOME_SOURCE_TC, constants.HOME_DESTINY_TC),
        "31": (constants.COM_SOURCE_HT, constants.COM_DESTINY_HT),
        "32": (constants.COM_SOURCE_TC, constants.COM_DESTINY_TC),
        "0": ("", "")
    }
    path_way = input("請輸入使用路徑(0 = 自訂, 1 = 雲端, 2 = 家庭, 3 = 公司 ：)")
    print("pathWay:", path_way)
    file_source_folder = ""
    file_dest_folder = ""
    if path_way == "0":
        file_source_folder = input("請輸入您的來源路徑 :") + "/" + file_doc_num
        file_dest_folder = input("請輸入您的目的路徑 :") + "/The_" + file_doc_num
    else:
        file_source_folder = path_map[path_way +
                                      file_plan_no][0] + "/" + file_doc_num
        file_dest_folder = (path_map[path_way + file_plan_no][1] + "/" +
                            constants.PLAN_LIST[int(file_plan_no)] + "_" + file_doc_num)
    print("來源路徑:", file_source_folder)
    print("目的路徑:", file_dest_folder)
    print("計畫代碼:", file_plan_no)
    print("文件號碼:", file_doc_num)
    print(r"----------------------file_path_process Done!--------------------------------")
    return file_doc_num, file_source_folder, file_dest_folder

def main():
    """main"""
    file_path = r"/workspaces/Auto-Work-Station/00dest/HT-D1-CTC-GEL-23-2638.xlsb"
    #copy_stand_file(file_path)
    print(file_path)


if __name__ == '__main__':
    main()
