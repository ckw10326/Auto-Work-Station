import os
import sys
from file_process import files_list, check_plan

# 1.取目錄路徑
root_path = os.path.dirname(os.path.abspath(__file__))
# 將根目錄路徑添加到 sys.path
sys.path.append(root_path)
source_folder = os.path.join(root_path, "00source")
destination_dir = os.path.join(root_path, "00dest")

print(check_plan(destination_dir))


#f for f in os.listdir(directory) if f.endswith('.' + extension)

# 1. os.listdir(dir_path) 列表出路徑的檔案及資料夾名稱
# print(os.listdir(dir_path))
# 2. "f for f in list" 是一種列表推導式（List comprehension）的用法，用於快速創建新的列表。

# for ss in os.listdir(dir_path):
#     if ss.endswith(".py"):
#         print(ss)