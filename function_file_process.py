# pylint: disable = w0612
"""
模組提供 20230828，移動work_flow到更適合的位置
1.files_list，表列清單
2.move_document，複製完整結構檔案
"""
import os
import shutil
import sys


def del_folder(keyword):
    """刪除關鍵字資料夾"""
    # 刪除dest00內所有資料
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    dest_foler = os.path.join(root_path, str(keyword))
    if os.path.exists(dest_foler):
        shutil.rmtree(dest_foler)


def make_folder(keyword):
    """創建空白資料夾"""
    root_path = os.path.dirname(os.path.abspath(__file__))
    dest_foler = os.path.join(root_path, str(keyword))
    if not os.path.exists(dest_foler):
        os.makedirs(dest_foler)


def files_list(xpath, search_str=None):
    """GPT 改良版本，輸入路徑、關鍵字"""
    the_file_list = []
    for root, _, files in os.walk(xpath):
        # 遍歷當前文件夾下的所有檔案
        for file in files:
            # 輸出檔案路徑
            thefile = os.path.join(root, file)
            if search_str is None or search_str in thefile:
                print("符合條件檔案", thefile)
                the_file_list.append(thefile)
    return the_file_list


def move_document(source_folder, dest_folder):
    """
    copytree具有缺陷，1.僅複製資料夾內的東西，2.若目標資料夾存在會報錯誤
    改良版本複製來源資料夾，包含資料夾本身
    """
    # 檢查來源資料夾是否存在
    if not os.path.exists(source_folder):
        print("找不到指定資料夾：" + source_folder)
        return None

    # 生成新資料夾路徑 EX:/workspaces/Auto-Work-Station/00source/00dest
    new_dest_folder = os.path.join(
        dest_folder, os.path.basename(source_folder))

    # 檢查目標資料夾是否存在
    if not os.path.exists(new_dest_folder):
        # 若不存在，直接複製資料夾
        shutil.copytree(source_folder, new_dest_folder)
        print("----------------------複製資料夾完成------------------------")
    else:
        choice = 2  # 預設選擇方法二

        if choice == 1:
            # 方法一：刪除後複製
            print("方法一：檢查資料夾", new_dest_folder, "已存在，開始刪除")
            shutil.rmtree(new_dest_folder)
            shutil.copytree(source_folder, new_dest_folder)
        else:
            # 方法二：不刪除直接複製
            print("方法二：檢查資料夾", new_dest_folder, "已存在，但不刪除")
            shutil.copytree(source_folder, new_dest_folder, dirs_exist_ok=True)
    return None


def check_plan(folder_path):
    """
    給予【母資料夾】路徑
    檢查資料夾及其子資料夾中是否存在符合條件的檔案
    興達返回1
    台中返回2
    """
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if "-CTC-" in file:  # 檢查【興達】關鍵字 "-CTC-"
                return 1
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if "TPC-TC" in file:  # 檢查【台中】關鍵字 "-CTC-"
                return 2
    return False


if __name__ == '__main__':
    pass
