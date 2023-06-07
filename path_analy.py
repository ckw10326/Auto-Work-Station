import os
import ntpath
A ="09Past/123/HT-D1-CTC-GEL-23-2299.xlsb"
B = "HT-D1-CTC-GEL-23-2299.xlsb"
'''
dirname 取得父路徑
C = os.path.dirname(A)
os.path.split 切割路徑及資料夾
D, E = os.path.split(A) 
C = os.path.dirname(A)
print(C)
D, E = os.path.split(A)
print(D,E)
F = os.path.split(D)
print(F)
'''
THE_EXCEL_FILE_PATH = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫\HT-D1-CTC-GEL-23-2345\HT1-1-CUC01-C5161-1.xlsx"
DESTINY_HT= r"/workspaces/Auto-Work-Station/new_folder"
DESTINY_HT1 = r"D:\00 興達計劃\05 EPC提供資料\HT"
#取得資料夾名稱
#首層資料夾 first_file_path
first_file_path = ntpath.split(THE_EXCEL_FILE_PATH)[0]
print("first_file_path：", first_file_path)
second_file_path = ntpath.split(first_file_path)[1]
print("second_file_path：", second_file_path)
make_new_folder_path = DESTINY_HT + "/" + second_file_path
print("make_new_folder_path", make_new_folder_path)

os.makedirs(make_new_folder_path)
