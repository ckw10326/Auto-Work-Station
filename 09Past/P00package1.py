
def test(x, y=1):
    # 函數程式碼
    print(x,"\n", y)
    pass

def test2():
    x = "符合條件檔案 /workspaces/Auto-Work-Station/()Merge pdf/COMMENT SHEET FOR TC0-1-KDE02-S0113-B.pdf"
    y = 0
    if 0 in x:
        y = 1
    print(y)

import os 
import shutil
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
    #test2()
    source_folder = "/workspaces/Auto-Work-Station/00source"
    dest_folder = "/workspaces/Auto-Work-Station/00dest"
    move_document(source_folder,dest_folder)

    #move_document()