import sys
import os
import shutil

file_dest_folder = "/workspaces/Auto-Work-Station/00dest"

#篩選檔案、資料夾，刪除GEL-TPC" or "GEI-TPC"
def name_filter():
    for root, dirs, files in os.walk(file_dest_folder):
        for dir in dirs:
            thedir = os.path.join(root, dir)
            if ("GEL-TPC" in thedir) or ("GEI-TPC" in thedir):
                print("將刪除",thedir)
                shutil.rmtree(thedir)
            else:
                pass
                
#篩選檔案大小
def size_filter():
    for root, dirs, files in os.walk(file_dest_folder):
        for file in files:
            thefile = os.path.join(root, file)
            if os.path.isfile(thefile) and os.path.getsize(thefile) > 150:  # 大於1MB
                print(thefile)

#刪除空文件夾
def dele_empty_folder():
    for 