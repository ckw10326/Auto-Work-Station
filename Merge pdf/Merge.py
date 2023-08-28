"""
MergePdf 合併資料夾內pdf檔案
from PyPDF2 import PdfMerger

def merge_pdfs(input_files, output_file):
    merger = PdfMerger()
    
    for file in input_files:
        merger.append(file)
    
    merger.write(output_file)
    merger.close()

# 要合併的 PDF 文件列表
input_files = ['file1.pdf', 'file2.pdf']

# 合併後的輸出文件名稱
output_file = 'merged.pdf'

# 呼叫函式進行合併
merge_pdfs(input_files, output_file)
"""
import os
from PyPDF2 import PdfMerger

#取得目前執行的腳本檔案的絕對路徑
dir_path = os.path.dirname( __file__)

def list_files(directory, extension):
    """產生一個附檔名符合【.extension】的列表"""
    # 1. os.listdir(dir_path) 列表出路徑的檔案及資料夾名稱
    # print(os.listdir(dir_path))
    # 2. "f for f in list" 是一種列表推導式（List comprehension）的用法，用於快速創建新的列表。
    return list(f for f in os.listdir(directory) if f.endswith('.' + extension))

def merge_pdfs(input_files, output_file):
    """輸入檔案清單，輸出檔案名稱"""
    merger = PdfMerger()

    for file in input_files:
        merger.append(file)
    merger.write(output_file)
    merger.close()

def main():
    """合併merge.py下所有的pdf檔案"""
    # 要合併的 PDF 文件列表
    pdflist = list_files(dir_path, "pdf")
    for i in range(len(pdflist)):
        pdflist[i] = os.path.join(dir_path, pdflist[i])
        print(pdflist[i])
    # 合併後的輸出文件名稱
    output_file = os.path.join(dir_path, 'merged.pdf')
    merge_pdfs(pdflist, output_file)


if __name__ == '__main__':
    main()
