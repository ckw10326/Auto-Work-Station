"""
用來儲存路徑參數
"""
import os
PLAN_LIST = ["", "HT", "TC", "HT2"]
LOCATION = ["", "家庭", "公司", "雲端"]
DOC_NO_STRUCTURE = {"HT": ["HT", "D1", "CTC", "GEL", "23"],
                    "TC": ["TPC", "TC(C0)", "CD", "23"],
                    "HT2": ["HT", "D1", "GEI", "GEL", "23"]
                    }
# CONSTANT
AZURE_SOURCE = r"/workspaces/Auto-Work-Station/00source"
AZURE_DESTINY = r"/workspaces/Auto-Work-Station/00source"
HOME_SOURCE_HT = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\HT"
HOME_DESTINY_HT = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\HT"
HOME_SOURCE_TC = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\TC"
HOME_DESTINY_TC = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\TC"
COM_SOURCE_HT = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫"
COM_DESTINY_HT = r"D:\00 興達計劃\05 EPC提供資料\HT"
COM_SOURCE_TC = r"\\\10.162.10.58\全處共用區\_Dwg\台中發電廠新建燃氣機組計畫"
COM_DESTINY_TC = r"D:\00 台中計劃\05 EPC提供資料\TC"
# 套印文件 [1興達2台中]+[0雲端 1家庭 2公司 ]
HT_PRINT_STD_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\08Reference_Files\HT套印標準.rtf"
HT_PRINT_STD_FILE12 = r"D:\00 興達計劃\05 EPC提供資料\HT套印標準.rtf"
HT_PRINT_STD_FILE10 = r"/workspaces/Auto-Work-Station/08Reference_Files/HT套印標準.rtf"
TC_PRINT_STD_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\08Reference_Files\TC套印標準.rtf"
TC_PRINT_STD_FILE22 = r"D:\00 台中計劃\05 EPC提供資料\TC套印標準.rtf"
TC_PRINT_STD_FILE20 = r"/workspaces/Auto-Work-Station/08Reference_Files/TC套印標準.rtf"
# 傳真文件
HT_FAX_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\08Reference_Files\HT 傳真 sample.doc"
HT_FAX_FILE12 = r"C:\00 興達計劃\05 EPC提供資料\HT 傳真 sample.doc"
HT_FAX_FILE10 = r"/workspaces/Auto-Work-Station/08Reference_Files/HT 傳真 sample.doc"
TC_FAX_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\08Reference_Files\TC 傳真 sample.doc"
TC_FAX_FILE22 = r"D:\00 台中計劃\05 EPC提供資料\TC 傳真 sample.doc"
TC_FAX_FILE20 = r"/workspaces/Auto-Work-Station/08Reference_Files/TC 傳真 sample.doc"
# 審查意見
HT_COMMENT_FILE11 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\08Reference_Files\HT 審查意見 sample.doc"
HT_COMMENT_FILE12 = r"D:\00 興達計劃\05 EPC提供資料\HT 審查意見 sample.doc"
HT_COMMENT_FILE10 = r"/workspaces/Auto-Work-Station/08Reference_Files/HT 審查意見 sample.doc"
TC_COMMENT_FILE21 = r"C:\Users\OXO\OneDrive\01 Book\00 Test program\Auto-Work-Station\08Reference_Files\TC 審查意見 sample.doc"
TC_COMMENT_FILE22 = r"D:\00 興達計劃\05 EPC提供資料\TC 審查意見 sample.doc"
TC_COMMENT_FILE20 = r"/workspaces/Auto-Work-Station/08Reference_Files/TC 審查意見 sample.doc"
# 暫時用不到的資料
GOOG_HT = r"/content/sample_data/HT"
GOOG_TC = r"/content/sample_data/TC"

# variable
source_customized = ""
destiny_customized = ""
plan_no = ""
doc_path = ""
location_no = ""
doc_no = ""
source_folder = ""  # 資料來源資料夾
dest_folder = ""  # 存放路徑資料夾
converted_xlsx_path = ""  # 轉檔路徑
# 來文資訊
letter_title = ""  # 來文資訊
letter_vision = ""  # 來文資訊
letter_num = ""  # 來文資訊
letter_date = ""  # 來文資訊

def main():
    """測試路徑是否存在"""
    #確認相關路徑是否存在
    if os.path.exists(AZURE_SOURCE):
        print("1. The folder exists!", AZURE_SOURCE)
    else:
        print("1. The folder does not exist!", AZURE_SOURCE)
    if os.path.exists(AZURE_DESTINY):
        print("2. The folder exists!", AZURE_DESTINY)
    else:
        print("2. The folder does not exist!", AZURE_DESTINY)
    #確認套印樣本是否存在
    if os.path.exists(HT_PRINT_STD_FILE10):
        print("3. The file exists!", HT_PRINT_STD_FILE10)
    else:
        print("3. The file does not exist!", HT_PRINT_STD_FILE10)

    if os.path.exists(TC_PRINT_STD_FILE20):
        print("4. The file exists!", TC_PRINT_STD_FILE20)
    else:
        print("4. The file does not exist!", TC_PRINT_STD_FILE20)
    #確認傳真樣本是否存在
    if os.path.exists(HT_FAX_FILE10):
        print("5. The file exists!", HT_FAX_FILE10)
    else:
        print("5. The file does not exist!", HT_FAX_FILE10)
    if os.path.exists(TC_FAX_FILE20):
        print("6. The file exists!", TC_FAX_FILE20)
    else:
        print("6. The file does not exist!", TC_FAX_FILE20)
    #確認審查意見是否存在
    if os.path.exists(HT_COMMENT_FILE10):
        print("3. The file exists!", HT_COMMENT_FILE10)
    else:
        print("3. The file does not exist!", HT_COMMENT_FILE10)
    if os.path.exists(TC_COMMENT_FILE20):
        print("3. The file exists!", TC_COMMENT_FILE20)
    else:
        print("3. The file does not exist!", TC_COMMENT_FILE20)

def test():
    print(type(DOC_NO_STRUCTURE))

if __name__ == '__main__':
    main()
    #test()
