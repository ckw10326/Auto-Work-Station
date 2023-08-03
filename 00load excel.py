'''
快速讀取文件內容
'''
import pandas as pd
import os
import shutil

source_path = "/workspaces/Auto-Work-Station/07docss"
for root, dirs, files in os.walk(source_path):
    for file in files:
        thepath = os.path.join(root, file)
        if  ".xlsx" in thepath:
            # 看一下檔案名稱
            print(thepath)
            df = pd.read_excel(thepath)
            print(df)

#刪除Done
if 0:
    source_path = "/workspaces/Auto-Work-Station/07docss"
    for root, dirs, files in os.walk(source_path):
        for file in files:
            thepath = os.path.join(root, file)
            if  ".xlsx" in thepath and "_Done" in thepath:
                # 看一下檔案名稱
                os.remove(thepath)
else:
    pass