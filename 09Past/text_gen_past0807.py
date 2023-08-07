"""
引用constant.py內常數
text_gen 輸出相關文字
copy_plan_file 複製套印檔案
file_path_process 輸入關鍵字，生成公司檔案路徑

"""
import shutil
import datetime
import os
import constants

def text_gen(def_letter_title, def_letter_vision, def_letter_num, def_letter_date, dest_folder):
    """五個輸入參數，來文標題、版次、文號、日期、資料夾"""
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
    path_map = {"10": (constants.HT_PRINT_STD_FILE, constants.HT_FAX_FILE, constants.HT_COMMENT_FILE),
                "11": (constants.HT_PRINT_STD_FILE, constants.HT_FAX_FILE, constants.HT_COMMENT_FILE),
                "12": (constants.HT_PRINT_STD_FILE12, constants.HT_FAX_FILE12, constants.HT_COMMENT_FILE12),
                "20": (constants.TC_PRINT_STD_FILE, constants.TC_FAX_FILE, constants.HT_COMMENT_FILE12),
                "21": (constants.TC_PRINT_STD_FILE, constants.TC_FAX_FILE, constants.HT_COMMENT_FILE12),
                "22": (constants.TC_PRINT_STD_FILE22, constants.TC_FAX_FILE22, constants.TC_COMMENT_FILE22)
                }

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
                     + "，本組無意見，已Email通知" + consult_company[plan_no] + "，擬陳閱後文存。")
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

def copy_plan_file(filepath, temp_folder = None):
    """複製套印文件，注意temp_folder必須為資料夾路徑"""
    # 取得檔案夾名稱、檔名(及副檔名)
    file_folder, file_allname = os.path.split(filepath)
    # 取得檔名
    file_name = file_allname.split(".")[0]

    # 複製套印文件
    print_dest_file = os.path.join(file_folder, f"套印_{file_name}.rtf")
    print("套印檔案:", file_allname)
    print("目標路徑:", file_folder)
    print("輸出套印檔案:", print_dest_file)
    input("請按任意鍵繼續")

    # 分析檔案名稱，來決定歸類
    if "HT-" in filepath:
        print_file = constants.HT_PRINT_STD_FILE
    elif "-TC" in filepath:
        print_file = constants.TC_PRINT_STD_FILE
    else:
        print_file = None

    if os.path.exists(print_file):
        # 用來測試的，請勿砍
        if temp_folder:
            print_dest_file = os.path.join(temp_folder , f"套印_{file_name}.rtf")
            # 檢查資料夾是否存在，若不存在則創建
            if not os.path.exists(temp_folder):
                os.makedirs(temp_folder)
                print("資料夾已成功創建")
            else:
                print("資料夾已存在")
            shutil.copy(print_file, print_dest_file)
            input("應該已創立temp_folder請檢查，輸入任意鍵後刪除")
            shutil.rmtree(temp_folder)
            return None    
        print("從﹝", print_file, "﹞複製到﹝", print_dest_file, "﹞")
        shutil.copy(print_file, print_dest_file)
        print("--------複製 套印文件 完成---------------")

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
        fax_file = constants.HT_FAX_FILE
    elif "-TC" in filepath:
        fax_file = constants.TC_FAX_FILE
    else:
        pass

    if os.path.exists(fax_file) and not os.path.exists(fax_dest_file):
        # 進行文件複製
        print("從﹝", fax_file, "﹞複製到﹝", fax_dest_file, "﹞")
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
        "11": (constants.SOURCE_FOLDER, constants.DEST_FOLDER),
        "12": (constants.SOURCE_FOLDER, constants.DEST_FOLDER),
        "21": (constants.SOURCE_FOLDER, constants.DEST_FOLDER),
        "22": (constants.SOURCE_FOLDER, constants.DEST_FOLDER),
        "31": (constants.COM_SOURCE_HT, constants.COM_DEST_HT),
        "32": (constants.COM_SOURCE_TC, constants.COM_DEST_TC),
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

def test():
    """各項測試"""
    def test_text_gen() -> None:
        """測試生成套印文字"""
        dest_folder = r"/workspaces/Auto-Work-Station/00source"
        text_gen('HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab',
                'A', 'TPC-TC(C0)-CD-23-0002', '2023/02/02', dest_folder)
        input("Press enter to exit...")
        return None
    
    def test_copy_plan_file() -> None:
        """測試程式"""
        # 從Sample資料夾取得範例
        file_path = r"/workspaces/Auto-Work-Station/08 Sample/HT-D1-CTC-GEL-23-3046_converted_Done.xlsx"
        # 這個是測試路徑，當有有輸入temp_folder時，優先複製到temp_folder
        temp_folder = r"/workspaces/Auto-Work-Station/()Temp_folder"
        copy_plan_file(file_path, temp_folder)
        """尚未完成"""
    
    def test_file_path_process()-> None:
        t = 123
    
    test_text_gen()
    test_copy_plan_file()
    test_file_path_process()
    
if __name__ == '__main__':
    test()
