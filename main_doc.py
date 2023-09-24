"""
pip install python-docx
"""
import os
import docx
import sys
import shutil
import docx_process
from file_process import files_list,del_folder,make_folder, move_document
import docx_process



path = "/workspaces/Auto-Work-Station/08Reference_Files/TC 傳真 sample.docx"
path2 = "/workspaces/Auto-Work-Station/00test/test.docx"
doc = docx.Document(path)


# 遍歷文件中的每個段落
for paragraph in doc.paragraphs:
    # 將 "abc" 取代為 "123"
    print(paragraph.text)
    # if 'abc' in paragraph.text:
    #     paragraph.text = paragraph.text.replace('abc', '123')