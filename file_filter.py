"""
將資料夾內不符合規格的資料刪除
"""
import os
import shutil

def name_filter(the_path):
    """篩選檔案、資料夾，刪除GEL-TPC or GEI-TPC"""
    for root, dirs, _ in os.walk(the_path):
        for directors in dirs:
            thedir = os.path.join(root, directors)
            if "CTC-GEL" in thedir:
                print("將刪除",thedir)
                shutil.rmtree(thedir)

def size_filter(the_path):
    '''篩選檔案大小'''
    for root, _ , files in os.walk(the_path):
        for file in files:
            thefile = os.path.join(root, file)
            if os.path.isfile(thefile) and os.path.getsize(thefile) > 150:  # 大於1MB
                print(thefile)

def dele_empty_folder():
    """刪除空文件夾"""

def test():
    """測試功能"""
    path1 = r"C:\Users\S\Documents\GitHub\Auto-Work-Station\00dest\00source"
    name_filter(path1)
    #因為name_filter有做修改故測試

if __name__ == '__main__':
    test()
