"""
find all excel file(.xlsx, .xlsb, .xlsm) of path
"""
import os
import ntpath # 解析windows文件
import shutil

"""
有缺陷的複製，僅能複製兩層
找到檔案
1.檔案完整路徑file_path =  /workspaces/Auto-Work-Station/00source/HT-D1-CTC-GEL-23-2068.xlsb
2.檔案資料夾路徑： /workspaces/Auto-Work-Station/00source
3.檔案資料夾名稱： 00source
4.新資料資料夾路徑，er_path /workspaces/Auto-Work-Station/00sdst/00source
5.新檔案完整路徑，new_excel_path /workspaces/Auto-Work-Station/00sdst/00source/HT-D1-CTC-GEL-23-2068.xlsb
6.新資料夾不存在，建立新資料夾
7.Excel檔案不存在，複製檔案
"""
def copy_excel_file(source_path, target_path):
    # 遍歷 09Past 目錄中的檔案及子目錄
    for root, dirs, files in os.walk(source_path):
        for file in files:
            # 如果找到符合條件之檔案
            if ((file.endswith((".xlsx", ".xlsb", ".xlsm")) & ("23" in file)) & (file.count("-") >= 3)):
                print("找到檔案")
                # 分析路徑
                file_path = os.path.join(root, file)
                print("1.檔案完整路徑file_path = ", file_path)
                first_file_path = ntpath.split(file_path)[0]
                print("2.檔案資料夾路徑：", first_file_path)
                second_file_path = ntpath.split(first_file_path)[1]
                print("3.檔案資料夾名稱：", second_file_path)
                make_new_folder_path = target_path + "/" + second_file_path
                print("4.新資料資料夾路徑，er_path", make_new_folder_path)
                new_excel_path = make_new_folder_path + "/" + ntpath.split(file_path)[1]
                print("5.新檔案完整路徑，new_excel_path", new_excel_path)
                #建立資料夾
                if not os.path.exists(make_new_folder_path):
                    print("6.新資料夾不存在，建立新資料夾")
                    os.makedirs(make_new_folder_path)
                # 建立檔案
                if not os.path.exists(new_excel_path):
                    print("7.Excel檔案不存在，複製檔案")
                    shutil.copy(file_path, new_excel_path)

"""
整個結構複製
找到檔案
1.檔案完整路徑file_path =  /workspaces/Auto-Work-Station/00source/HT-D1-CTC-GEL-23-2068.xlsb
2.檔案資料夾路徑： /workspaces/Auto-Work-Station/00source
3.檔案資料夾名稱： 00source
4.新資料資料夾路徑，er_path /workspaces/Auto-Work-Station/00sdst/00source
5.新檔案完整路徑，new_excel_path /workspaces/Auto-Work-Station/00sdst/00source/HT-D1-CTC-GEL-23-2068.xlsb
6.新資料夾不存在，建立新資料夾
7.Excel檔案不存在，複製檔案
"""
def copy_excel_file2(source_path, target_path):
    # 遍歷 09Past 目錄中的檔案及子目錄
    for root, dirs, files in os.walk(source_path):
        for file in files:
            # 如果找到符合條件之檔案
            if ((file.endswith((".xlsx", ".xlsb", ".xlsm")) & ("23" in file)) & (file.count("-") >= 3)):
                print("找到檔案")
                # 分析路徑
                file_path = os.path.join(root, file)
                print("1.檔案完整路徑file_path = ", file_path)
                #end_path 為對比路徑 Ex:/45 copy/HT-D1-CTC-GEL-23-2068_converted.xlsx
                end_path = file_path.replace(source_path, "")
                #new_file_path 為新路徑 target_path + /45 copy/HT-D1-CTC-GEL-23-2068_converted.xlsx
                new_file_path = target_path + end_path
                new_file_folder = ntpath.split(new_file_path)[0]
                #建立資料夾
                if not os.path.exists(new_file_folder):
                    print("2.新資料夾不存在，建立新資料夾")
                    os.makedirs(new_file_folder)
                # 建立檔案
                if not os.path.exists(new_file_path):
                    print("7.Excel檔案不存在，複製檔案")
                    shutil.copy(file_path, new_file_path)

def main():
    source_path = "/workspaces/Auto-Work-Station/00source"
    target_path = "/workspaces/Auto-Work-Station/00sdst"
    #先刪除資料夾內檔案
    if ntpath.exists(target_path):
        shutil.rmtree(target_path)
        input("已經刪除資料夾，plz key any key")
    else:
        pass
    os.makedirs(target_path)
    copy_excel_file2(source_path, target_path)

if __name__ == '__main__':
    main()
