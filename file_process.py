"""
模組提供 20230803
1.files_list，表列清單
2.move_document，複製完整結構檔案
3.workflow，將所有流程統一化
"""

import os
import shutil
import sys
import read_ht_excel_cloud0712


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


def work_flow():
    '''
    1.設定模組路徑  2.複製檔案  2.列出xlsb並轉換  3.製作表格存檔
    目前只有HT
    '''
    sample_folder = "/workspaces/Auto-Work-Station/00source"
    destination_dir = "/workspaces/Auto-Work-Station/00dest"

    # 1.使用外部模組
    if 1:
        root_git = r"/workspaces/Auto-Work-Station"
        sys.path.append(root_git)

    # 2.複製檔案
    move_document(sample_folder, destination_dir)

    def check_ht(destination_dir):
        """3.解析.xlsb(興達資料)，判斷是否有興達.xlsb檔案，轉換"""
        xlsb_file_list = files_list(destination_dir, ".xlsb")
        # 展生清單 轉換檔案
        if xlsb_file_list:
            for xlsb_path in xlsb_file_list:
                read_ht_excel_cloud0712.convert_xlsb(xlsb_path)

        # 將.xlsb轉換成.xlsx後，開始分析內容
        xlsx_file_list = files_list(destination_dir, "converted.xlsx")
        if xlsx_file_list:
            for converted_path in xlsx_file_list:
                print("符合興達計畫converted.xlsb路徑:", converted_path)
                # ●●●●●●
                list0 = read_ht_excel_cloud0712.read_ht_make_list(
                    converted_path) if read_ht_excel_cloud0712.read_ht_make_list(converted_path) else 1
                input("按任意鍵退出")
                return list0

    def check_tc(destination_dir):
        """尚未完成功能"""
        print("check_TC功能尚未完成，抱歉", destination_dir)

    check_ht(destination_dir)
    check_tc(destination_dir)

def main():
    """main"""
    dest_folder = "/workspaces/Auto-Work-Station/00dest"
    if os.path.exists(dest_folder):
        shutil.rmtree(dest_folder)
    work_flow()

def test():
    """測試"""
    excel_path = "/workspaces/Auto-Work-Station/00dest/00source/HT-D1-CTC-GEL-23-2710_converted.xlsx"
    old_csv_path = "/workspaces/Auto-Work-Station/01Class/data.csv"
    read_ht_excel_cloud0712.read_ht_make_list(excel_path, old_csv_path)


if __name__ == '__main__':
    main()
