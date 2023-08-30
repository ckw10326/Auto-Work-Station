import os
import shutil
"""
python功能中，輸入"D:\Scan\123.pdf"，
1.輸出"123.pdf" 2.輸出"123" 3.輸出"D:\Scan"如何執行
"""
path = "D:\Scan\123.pdf"

# 獲取文件名
filename = os.path.basename(path)
print(filename) # 輸出：123.pdf

# 獲取文件名 (不包括擴展名)
name = os.path.splitext(filename)[0]
print(name) # 輸出：123

# 獲取文件路徑
folder = os.path.dirname(path)
print(folder) # 輸出：D:\Scan

"""
請用python寫一個程式驗證字串A中包含三個"-"符號
若為真，則輸出1
若為否，則輸出0
"""
def str_analy():
    str_A = "abc-def-ghi-jkl"
    # 計算"-"符號的個數
    count = str_A.count("-")
    if count == 3:
        print(1)
    else:
        print(0)

def filepath():
    filepath = "/workspaces/Auto-Work-Station/00dest/TPC-TC(C0)-CD-23-2388_Done.xlsx"
    filepath2 = "/workspaces/Auto-Work-Station/00dest"
    file_folder = os.path.split(filepath)[0]
    file_allname = os.path.basename(filepath)
    file_name = file_allname.split(".")[0]
    print(file_folder)
    print(file_allname)
    print(file_name)
    print(filepath2.split("/")[-1])


def copy_directory(src, dst):
    shutil.copytree(src, dst, dirs_exist_ok=True)

if __name__ == '__main__':
    '''
    filepath = "/workspaces/Auto-Work-Station/00source"
    filepath2 = "/workspaces/Auto-Work-Station/00dest"
    shutil.copytree(filepath, filepath2, dirs_exist_ok=True)
    #filepath()
    '''

    # 複製整個資料夾
    source_folder = "/workspaces/Auto-Work-Station/00source"
    dest_folder = "/workspaces/Auto-Work-Station/00dest"
    # 先刪除資料夾內容
    shutil.rmtree(dest_folder)
    input("xdxxxd")
    copy_directory(source_folder, os.path.join(dest_folder, os.path.basename(source_folder)))
