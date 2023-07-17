"""
find all excel file(.xlsx, .xlsb, .xlsm) of path
"""
import os
import ntpath # 解析windows文件
import shutil
PATH1 = r"/workspaces/Auto-Work-Station"
DEST_FOLDER =  r"/workspaces/Auto-Work-Station/new_folder"

def copy_excel_file(source_path, target_path):
    "測試功能"
    source_path = "/workspaces/Auto-Work-Station/00source"
    target_path = "/workspaces/Auto-Work-Station/00sdst"

    # 遍歷 09Past 目錄中的檔案及子目錄
    for root, dirs, files in os.walk(source_path):
        for file in files:
            # 如果找到指定的檔案
            
            if ((file.endswith((".xlsx", ".xlsb", ".xlsm")) & ("23" in file)) & (file.count("-") >= 3):
                print("找到檔案")
                # 分析路徑
                file_path = os.path.join(root, file)
                print("file_path = ", file_path)
                first_file_path = ntpath.split(file_path)[0]
                print("first_file_path：", first_file_path)
                second_file_path = ntpath.split(first_file_path)[1]
                print("second_file_path：", second_file_path)
                make_new_folder_path = target_path + "/" + second_file_path
                print("新Excel資料夾，make_new_folder_path", make_new_folder_path)
                new_excel_path = make_new_folder_path + "/" + ntpath.split(file_path)[1]
                print("新Excel檔案，new_excel_path", new_excel_path)
                #建立資料夾
                if ntpath.exists(make_new_folder_path):
                    print("檔案存在，無須建立資料夾")
                    pass
                else:
                    print("檔案不存在，建立新資料夾")
                    os.makedirs(make_new_folder_path)
                #建立檔案
                if ntpath.exists(new_excel_path):
                    print("Excel檔案已存在，無須動作")
                    pass
                else:
                    shutil.copy(file_path, new_excel_path)

def test():
    "測試功能"
    source_path = "/workspaces/Auto-Work-Station/09Past"
    target_path = "D:\\target_folder"
    search_file = "txt.xlsx"

    # 遍歷 09Past 目錄中的檔案及子目錄
    for root, dirs, files in os.walk(source_path):
        for file in files:
            # 如果找到指定的檔案
            if file == search_file:
                # 分析路徑
                file_path = os.path.join(root, file)
                target_file_path = os.path.join(target_path, os.path.relpath(file_path, source_path))
                target_dir = os.path.dirname(target_file_path)
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir)
                shutil.copy2(file_path, target_file_path)
                print(f"{file_path} copied to {target_file_path}")
                # 只複製一次，跳出迴圈
                break
            else:
                continue

def test2():
    "測試功能2"
    for root, dirs, files in os.walk(PATH1):
        for file in files:
            if file.endswith((".xlsx", ".xlsb", ".xlsm")):
                print("file:", file)
                file_path = os.path.join(root,file)
                print("file_folder", root)
                print("file_path:", file_path)

def copy_excel_files(def_path, def_dest_folder):
    "複製根目錄"
    path = def_path
    dest_folder = def_dest_folder
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(('.xlsx', '.xlsb', '.xlsm')):
                "符合條件後執行以下"
                file_path = os.path.join(root, file)
                print("file_path:",file_path)
                file_root = root
                print("file_root:",file_root)
                shutil.copytree(file_root, def_dest_folder)

    print('Done')

def main():
    source_path = "/workspaces/Auto-Work-Station/"
    target_path = "/workspaces/Auto-Work-Station/new_folder"
    copy_excel_file(source_path, target_path)

if __name__ == '__main__':
    main()
