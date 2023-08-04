"""
將資料夾內不符合規格的資料刪除
20230804，編寫測試程序
"""
import os
import shutil
import file_process

def folder_dele(the_path, strs = "CTC-GEL"):
    """篩選檔案、資料夾，刪除GEL-TPC or GEI-TPC"""
    for root, dirs, _ in os.walk(the_path):
        for directors in dirs:
            thedir = os.path.join(root, directors)
            if strs in thedir:
                print("將刪除",thedir)
                shutil.rmtree(thedir)

def file_size_filter(the_path, size):
    '''給予資料夾路徑，篩選檔案大小'''
    for root, _ , files in os.walk(the_path):
        for file in files:
            thefile = os.path.join(root, file)
            print("路徑thefile", thefile, "的大小為", os.path.getsize(thefile))
            if os.path.getsize(thefile) > size:
                print("這個檔案達於指定大小")
def test():
    """測試功能"""
    def movedoc():
        '''先複製範例檔案到dest00'''
        path1 = "/workspaces/Auto-Work-Station/09Past"
        path2 = "/workspaces/Auto-Work-Station/00dest"
        file_process.move_document(path1, path2)

    def test_name_filter():
        '''測試name_filter功能'''
        path2 = "/workspaces/Auto-Work-Station/00dest"
        folder_dele(path2, "package")

    def test_file_size_filter():
        """測試file_size_filter功能"""
        path2 = "/workspaces/Auto-Work-Station/00dest"
        file_size_filter(path2, 5000)

    movedoc()
    test_name_filter()
    test_file_size_filter()

if __name__ == '__main__':
    test()
