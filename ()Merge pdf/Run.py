"""
MergePdf 合併資料夾內pdf檔案
"""
# pylint: disable=W0105

from PyPDF2 import PdfMerger
import webbrowser
import os
dir_path = os.path.dirname(os.path.realpath(__file__))
print(dir_path)

#產生以下方法會產生物件，無法疊代
def list_files(directory, extension):
    return (f for f in os.listdir(directory) if f.endswith('.' + extension))
#以下是改進清單
def list_files2(directory, extension):
    return list(f for f in os.listdir(directory) if f.endswith('.' + extension))

pdfs2 = list_files2(dir_path, "pdf")
print("pdfs2清單:",pdfs2)

merger = PdfMerger()
for pdf in pdfs2:
    print("pdf檔名:",pdf)
    newpath = os.path.join(dir_path, pdf) 
    print("pdf路徑:", newpath)
    merger.append(open(newpath, 'rb'))

result_path = os.path.join(dir_path, 'Merge_result.pdf')
print(result_path)
with open(result_path, 'wb') as fout:
    merger.write(fout)
