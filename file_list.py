import os

def files_list1(xpath, str):
    the_file_list = []
    if str == 0:
        for root, dirs, files in os.walk(xpath):
            # 遍歷當前文件夾下的所有檔案
            for file in files:
                # 輸出檔案路徑
                thefile = os.path.join(root, file)
                print(thefile)
                the_file_list.append(thefile)
        print("----------", str, "檔案清單輸出完成-----------\n")
        return the_file_list
    else:
        for root, dirs, files in os.walk(xpath):
            # 遍歷當前文件夾下的所有檔案
            for file in files:
                # 輸出檔案路徑
                thefile = os.path.join(root, file)
                if str in thefile:
                    print(thefile)
                    the_file_list.append(thefile)
        print("----------", str, "檔案清單輸出完成-----------\n")
        return the_file_list