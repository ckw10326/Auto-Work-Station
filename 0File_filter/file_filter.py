"""
將資料夾內不符合規格的資料刪除
20230804，編寫測試程序
"""
import os
import shutil

def folder_dele(the_path, strs = "CTC-GEL"):
    """刪除符合【關鍵字】資料夾(GEL-TPC or GEI-TPC)"""
    for root, dirs, _ in os.walk(the_path):
        for directors in dirs:
            thedir = os.path.join(root, directors)
            if strs in thedir:
                print("將刪除",thedir)
                shutil.rmtree(thedir)

def file_size_filter(the_path, size):
    '''顯示符合【篩選檔案大小】的資料夾、檔案'''
    for root, _ , files in os.walk(the_path):
        for file in files:
            thefile = os.path.join(root, file)
            file_size = (os.path.getsize(thefile))/1024/1024
            target_size = int(size)*1024*1024
            if os.path.getsize(thefile) > target_size:
                print("路徑thefile", thefile, "的大小為", round(file_size,2), "MB")

def test():
    def test_name_filter():
        '''測試name_filter功能'''
        path2 = "/workspaces/Auto-Work-Station/00dest"
        folder_dele(path2, "package")

    def test_file_size_filter():
        """測試file_size_filter功能"""
        path2 = "/workspaces/Auto-Work-Station/00dest"
        file_size_filter(path2, 5000)

    test_name_filter()
    test_file_size_filter()

def sample():
    path = r"C:\Users\S\Downloads"
    folder_dele(path, "16k")
    path = r"C:\Users\S\Downloads"
    file_size_filter(path, "10")

def main():


if __name__ == '__main__':
    main()
