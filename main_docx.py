'''
2023/9/25 完成台中計畫，套印docx自動帶入文字內容

這個檔案有兩個功能
1.雲端分析檔案，產生套印、傳真資料(2023/8/28完成)
2.分析一個檔案(2023/8/28完成)
3.具有將資料夾【所有符合關鍵字】的檔案分析能力

'''
import os
import sys
import docx
from function_file_process import move_document, files_list, check_plan, del_folder, make_folder
from function_docx import replace_keyword_in_docx, read_docx_file
from read_ht import convert_xlsb
from read_ht import r_ctc_xlsx_sheet
from read_tc import r_tc_xlsm_sheet
from text_gen import copy_plan_file, text_gen

def work_flow():
    '''
    讀取【一個】檔案
    1.設定目錄路徑  2.複製檔案  3.解析檔案(HT TC皆完成)
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    source_folder = os.path.join(root_path, "00source")
    dest_folder = os.path.join(root_path, "00dest")
    # 檢查source_folder是否存在
    if not os.path.exists(source_folder):  # 注意這裡的小寫 "n" in "not"
        print("source_folder不存在，故結束")
        sys.exit()

    # 2.複製檔案
    move_document(source_folder, dest_folder)

    # 3.分析plan檔案，興達返回1 台中返回2
    if str(check_plan(dest_folder)) == "1":
        # 讀取興達
        # 3.1 產生xlsb列表，並轉換成xlsx
        xlsb_file_list = files_list(dest_folder, ".xlsb")
        if xlsb_file_list:
            for xlsb_path in xlsb_file_list:
                convert_xlsb(xlsb_path)

        # 3.2 產生converted.xlsx列表，開始分析內容
        xlsx_file_list = files_list(dest_folder, "converted.xlsx")
        if xlsx_file_list:
            for converted_path in xlsx_file_list:
                draw_title, draw_vision, l_num, letter_date = r_ctc_xlsx_sheet(
                    converted_path)
                print(draw_title, draw_vision, l_num, letter_date)
                text_gen(draw_title, draw_vision, l_num, letter_date)
                # 複製套印、傳真檔案
                print_dest_file, fax_dest_file = copy_plan_file(converted_path)
                print(print_dest_file, "\n", fax_dest_file)

    elif str(check_plan(dest_folder)) == "2":
        # 讀取台中
        # 3.1 產生xlsb列表，並轉換成xlsx
        xlsm_list = files_list(dest_folder, ".xlsm")
        if xlsm_list:
            for xlsm in xlsm_list:
                draw_title, draw_vision, l_num, l_date = r_tc_xlsm_sheet(xlsm)
                print(draw_title, draw_vision, l_num, l_date)
                contents0, contents1, contents2, contents3, contents4 = text_gen(draw_title, draw_vision, l_num, l_date)
                # 複製套印、傳真檔案
                print_dest_file, fax_dest_file = copy_plan_file(xlsm)
                print(print_dest_file, "\n", fax_dest_file)
                #取代docx檔案
                replace_keyword_in_docx(print_dest_file, "123", contents0)

    else:
        print("沒有符合檔案")
        sys.exit()


if __name__ == '__main__':
    del_folder("00dest")
    work_flow()
    del_folder("00source")
    make_folder("00source")

    #新測試流程
    # del_folder("00dest")
    # work_flow()