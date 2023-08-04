"""
用來儲存路徑參數
20230804，修改程式，路徑自動抓取
"""
# pylint: disable=invalid-name
# pylint: disable=line-too-long
import os

PLAN_LIST = ["", "HT", "TC", "HT2"]
LOCATION = ["", "家庭", "公司", "雲端"]
DOC_NO_STRUCTURE = {"HT": ["HT", "D1", "CTC", "GEL", "23"],
                    "TC": ["TPC", "TC(C0)", "CD", "23"],
                    "HT2": ["HT", "D1", "GEI", "GEL", "23"]
                    }
# CONSTANT
"""GPT建議如何設定不同電腦不同設定"""
# 獲取當前工作目錄的路徑
current_dir = os.getcwd()
#CONSTANT
SOURCE_FOLDER = os.path.join(current_dir, "00source")
DEST_FOLDER = os.path.join(current_dir, "00dest")
COM_SOURCE_HT = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫"
COM_SOURCE_TC = r"\\\10.162.10.58\全處共用區\_Dwg\台中發電廠新建燃氣機組計畫"
COM_DEST_HT = r"D:\00 興達計劃\05 EPC提供資料\HT"
COM_DEST_TC = r"D:\00 台中計畫\05 EPC提供資料\TC"
# 套印文件 [1興達2台中]+[0雲端 1家庭 2公司 ]
HT_PRINT_STD_FILE =  os.path.join(current_dir, "08Reference_Files/HT套印標準.rtf")
TC_PRINT_STD_FILE = os.path.join(current_dir, "08Reference_Files/TC套印標準.rtf")
HT_PRINT_STD_FILE12 = r"D:\00 興達計劃\05 EPC提供資料\HT套印標準.rtf"
TC_PRINT_STD_FILE22 = r"D:\00 台中計劃\05 EPC提供資料\TC套印標準.rtf"
# 傳真文件
HT_FAX_FILE = os.path.join(current_dir, "08Reference_Files/HT 傳真 sample.doc")
TC_FAX_FILE = os.path.join(current_dir, "08Reference_Files/TC 傳真 sample.doc")
HT_FAX_FILE12 = r"C:\00 興達計劃\05 EPC提供資料\HT 傳真 sample.doc"
TC_FAX_FILE22 = r"D:\00 台中計劃\05 EPC提供資料\TC 傳真 sample.doc"
# 審查意見
HT_COMMENT_FILE = os.path.join(current_dir, "08Reference_Files/HT 審查意見 sample.doc")
TC_COMMENT_FILE = os.path.join(current_dir, "08Reference_Files/TC 審查意見 sample.doc")
HT_COMMENT_FILE12 = r"D:\00 興達計劃\05 EPC提供資料\HT 審查意見 sample.doc"
TC_COMMENT_FILE22 = r"D:\00 興達計劃\05 EPC提供資料\TC 審查意見 sample.doc"

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
    def test_cloud() -> None:
        '''確認相關路徑是否存在'''
        print(f"1. The SOURCE_FOLDER {'exists' if os.path.exists(SOURCE_FOLDER) else ' Dont exist!'}! {SOURCE_FOLDER}")
        print(f"2. The DEST_FOLDER {'exists' if os.path.exists(DEST_FOLDER) else ' Dont exist!'}! {DEST_FOLDER}")
        #確認套印樣本是否存在
        print(f"3. The HT_PRINT_STD_FILE {'exists' if os.path.exists(HT_PRINT_STD_FILE) else ' Dont exist!'}! {HT_PRINT_STD_FILE}")
        print(f"4. The TC_PRINT_STD_FILE {'exists' if os.path.exists(TC_PRINT_STD_FILE) else ' Dont exist!'}! {TC_PRINT_STD_FILE}")
        #確認傳真樣本是否存在
        print(f"5. The HT_FAX_FILE {'exists' if os.path.exists(HT_FAX_FILE) else ' Dont exist!'}! {HT_FAX_FILE}")
        print(f"6. The TC_FAX_FILE {'exists' if os.path.exists(TC_FAX_FILE) else ' Dont exist!'}! {TC_FAX_FILE}")
        #確認審查意見是否存在
        print(f"7. The HT_COMMENT_FILE {'exists' if os.path.exists(HT_COMMENT_FILE) else ' Dont exist!'}! {HT_COMMENT_FILE}")
        print(f"8. The TC_COMMENT_FILE {'exists' if os.path.exists(TC_COMMENT_FILE) else ' Dont exist!'}! {TC_COMMENT_FILE}")

    test_cloud()

def test():
    '''獲取電腦名稱'''
    #linux系統使用
    computer_name = os.uname().nodename
    print(computer_name)

if __name__ == '__main__':
    main()
    test()
