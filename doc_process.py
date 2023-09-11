"""
pip install python-docx
"""
import os
import docx
import sys
import shutil
from file_process import files_list,del_folder,make_folder, move_document

root = os.path.dirname(os.path.abspath(__file__))
# 設定樣本路徑
sample_fax_path = os.path.join(root, r"09Past/test.docx")
# 設定測試資料夾路徑
test_folder_path = os.path.join(root, "00test")

# 進行主程式
# del_folder("00test")
# make_folder("00test")
# shutil.move(sample_fax_path, test_folder_path)

doc_path = "/workspaces/Auto-Work-Station/00test/test.docx"
doc = docx.Document(doc_path)


para = doc.paragraphs[11]
print(para.text + '\n')
print('run數量： ', len(para.runs))

for i in range(0, len(para.runs)):
    print(i, para.runs[i].text)