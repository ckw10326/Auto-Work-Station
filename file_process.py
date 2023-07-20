import os
import shutil

#GPT 改良版本
def files_list(xpath, search_str=None):
    the_file_list = []
    for root, dirs, files in os.walk(xpath):
        # 遍歷當前文件夾下的所有檔案
        for file in files:
            # 輸出檔案路徑
            thefile = os.path.join(root, file)
            if search_str is None or search_str in thefile:
                print("符合條件檔案", thefile)
                the_file_list.append(thefile)
    return the_file_list

#我的臃腫版本
def files_list1(xpath, str):
    the_file_list = []
    if str == 0:
        for root, dirs, files in os.walk(xpath):
            # 遍歷當前文件夾下的所有檔案
            for file in files:
                # 輸出檔案路徑
                thefile = os.path.join(root, file)
                print("符合條件檔案", thefile)
                the_file_list.append(thefile)
        #print("----------", str, "檔案清單輸出完成-----------\n")
        return the_file_list
    else:
        for root, dirs, files in os.walk(xpath):
            # 遍歷當前文件夾下的所有檔案
            for file in files:
                # 輸出檔案路徑
                thefile = os.path.join(root, file)
                if str in thefile:
                    print("符合條件檔案", thefile)
                    the_file_list.append(thefile)
        #print("----------", str, "檔案清單輸出完成-----------\n")
        return the_file_list

# copytree具有缺陷，1.僅複製資料夾內的東西，2.若目標資料夾存在會報錯誤
# 複製來源資料夾，包含資料夾本身
def move_document(source_folder, dest_folder):
    # 檢查來源資料夾是否存在
    if not os.path.exists(source_folder):
        print("找不到指定資料夾：" + source_folder)
        return None
    print("source_folder資料夾存在：" + source_folder)

    # 檢查目標資料夾是否存在    
    # 生成新資料夾路徑 EX:/workspaces/Auto-Work-Station/00source/00dest
    new_dest_folder = dest_folder + "/" + source_folder.split("/")[-1]
    if not os.path.exists(new_dest_folder):
        #若不存在，直接複製資料夾
        shutil.copytree(source_folder, new_dest_folder)
    else:
        #choice = 1
        choice = 2
        #若存在，方法一先刪除後複製
        if choice == 1:
            print("方法一：檢查資料夾", new_dest_folder, "已存在，開始刪除")
            shutil.rmtree(new_dest_folder)
            shutil.copytree(source_folder, new_dest_folder)
        else:
            print("方法二：檢查資料夾", new_dest_folder, "已存在，但不刪除")
            shutil.copytree(source_folder, new_dest_folder, dirs_exist_ok=True)
    return None

if __name__ == '__main__':
    x = "/workspaces/Auto-Work-Station/()Merge pdf"
    files_list1(x, ".pdf")
    print("------------000-----------")
    files_list(x, ".pdf")