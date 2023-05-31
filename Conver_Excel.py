#輸入路徑解析xlsb成.xlsx
import os
import fnmatch
import pandas

directory = input("請輸入路徑:")
# search for all files with the .xlsb extension
for file in os.listdir(directory):
    if fnmatch.fnmatch(file, '*.xlsb'):
        File_Path =os.path.join(directory, file)
        print("找到檔案：", File_Path)
#顯示表格
df = pd.read_excel(File_Path, sheet_name='Data1', engine='pyxlsb')

Xlsx_File_Path = os.path.splitext(File_Path)[0] + ".xlsx"
print("輸出檔案：", Xlsx_File_Path)
#輸出檔案
df.to_excel(Xlsx_File_Path)