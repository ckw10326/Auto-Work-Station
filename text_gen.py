"""
引用constant.py內常數
text_gen 輸出相關文字
copy_plan_file 複製套印檔案
file_path_process 輸入關鍵字，生成公司檔案路徑
20230807
1.完成測試功能
2.刪除dest_folder參數
"""
import shutil
import datetime
import os
import constants

def text_gen(letter_title, letter_rev, letter_num, letter_date) -> None:
    """四個輸入參數，來文標題、版次、文號、日期、資料夾"""
    # 判斷有意見的單位
    my_company_num = int(
        input("請輸入審查意見填寫單位:\n 1 = 南部施工處 \n 2 = 中部施工處 \n 3 = 興達發電廠 \n 4 =台中發電廠\n"))
    pages = str(input("請輸入審查意見頁數:"))
    company_name = ["", "南部施工處", "中部施工處", "興達發電廠", "台中發電廠"]
    # 判斷顧問公司
    consult_company = "泰興公司" if "-TC" in letter_num else "吉興公司"
    # 判斷計畫別
    plan_name = "興達" if "CTC" in letter_num else "台中"
    date_obj = datetime.datetime.strptime(letter_date, "%Y/%m/%d")
    month = date_obj.strftime("%m")
    day = date_obj.strftime("%d")

    contents0 = str("本文係" + company_name[my_company_num] + "對統包商提送「" + letter_title
                 + "」" + "Rev." + str(letter_rev) + "所提審查意見(共" + pages + "頁)"
                 + "，未逾合約規範，已電傳" + consult_company + "，擬陳閱後文存。")
    contents1 = str("檢送" + plan_name + "電廠燃氣機組更新改建計畫「" + letter_title + "」，Rev."
                 + str(letter_rev) + "，" + company_name[my_company_num]
                 + "之審查意見（如附，共" + pages + "頁）供卓參，請查照。")
    contents2 = str("依據GE/CTCI 112年" + month + "月" + day + "日" + letter_num + "號辦理。")
    contents3 = str("本文係統包商提送「" + letter_title + "」" + "Rev." + str(letter_rev)
                 + "，本組無意見，已Email通知" + consult_company + "，擬陳閱後文存。")
    print("----------------套印內容------------------")
    print(contents0)
    print("----------------傳真------------------------------")
    print(contents1)
    print("----------------依據------------------------------")
    print(contents2)
    print("----------------主辦簽辦--------------------------")
    print(contents3)
    input("enter any kesy to exit")
    return contents0, contents1, contents2, contents3

def copy_plan_file(filepath) -> None:
    """複製套印文件到與file_path相同資料夾，注意temp_folder必須為資料夾路徑"""
    # 取得檔案夾名稱、檔名(及副檔名)
    file_folder, file_allname = os.path.split(filepath)
    # 取得檔名
    file_name = file_allname.split(".")[0]

    # 套印、傳真文件名稱
    print_dest_file = os.path.join(file_folder, f"套印_{file_name}.rtf")
    fax_dest_file = os.path.join(file_folder, f"傳真_{file_name}.doc")
    print("套印檔案名稱:", file_allname)
    print("目標資料夾路徑:", file_folder)
    print("輸出套印檔案:", print_dest_file)
    print("輸出傳真檔案:", fax_dest_file)

    # 分析檔案名稱，來決定歸類
    print_file = ""
    fax_file = ""
    if "HT-" in str(filepath):
        print_file = constants.HT_PRINT_STD_FILE
        fax_file = constants.HT_FAX_FILE
    elif "-TC" in str(filepath):
        print_file = constants.TC_PRINT_STD_FILE
        fax_file = constants.TC_FAX_FILE
    else:
        print("---------------檔名不符合關鍵字，不動作-----------------")
        return None
    shutil.copy2(print_file, print_dest_file)
    shutil.copy2(fax_file, fax_dest_file)
    print("--------複製 套印、傳真文件 完成---------------")
    return None

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

def test_text_gen() -> None:
    """測試生成套印文字"""
    text_gen('HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab',
                'A',
                'TPC-TC(C0)-CD-23-0002',
                '2023/02/02',)
    return None

if __name__ == '__main__':
    test_text_gen()
