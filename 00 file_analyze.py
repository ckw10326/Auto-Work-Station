'''
測試子程式
測試整合流程
'''
import sys
import os
import shutil
from file_process import files_list, move_document
from read_ht_excel_cloud0712 import convert_xlsb
from read_ht_excel_cloud0712 import read_ctc_ht_excel
from read_tc_excel_cloud0712 import read_tc_excel
from text_gen import file_path_process
from text_gen import text_gen
from text_gen import copy_stand_file

file_doc_num = ""
file_source_folder = "/workspaces/Auto-Work-Station/08 Excel"
file_dest_folder = "/workspaces/Auto-Work-Station/00dest"


def ht_excel_analyze(file_dest_folder):
    """
    # 分析網路空間Excel檔案
    # 1.複製資料夾
    # 2.處理Excel文件
    # 3.檢查內容
    # 4.合併文件
    """
    ht_xlsx_list = files_list(file_dest_folder, ".xlsx")
    for ht_file in ht_xlsx_list:
        read_ctc_ht_excel(ht_file)
    pass


if __name__ == '__main__':
    file_dest_folder = "/workspaces/Auto-Work-Station/07docss"
    ht_excel_analyze(file_dest_folder)

    # test_folder_path()

    # shutil.rmtree(file_dest_folder)
    # test_cloud_ht()

    # shutil.rmtree(file_dest_folder)
    # test_cloud_tc()

    # test_path()

    # test_text_gen()

    # shutil.rmtree(file_dest_folder)
    # cloud_total_run()
else:
    pass
