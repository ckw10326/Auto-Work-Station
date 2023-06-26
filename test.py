import os

path = r"/workspaces/Auto-Work-Station/00source"

xlsx_path = ""
"尋找檔案是否存在"
if os.path.exists(path):
    print("找到目標資料夾：", path)
    for file in os.listdir(path):
        print(file)


'''
        #確認xlsb檔案是否存在
        if fnmatch.fnmatch(file, '*.xlsb'):
            file_path =os.path.join(folder_path, file)
            print("找到目標檔案路徑：", file_path)
            #抓到隱藏檔案(檔名有關鍵字：~$H，不動作
            if "~$" in file_path:
                print(file_path,"抓到隱藏檔案(檔名有關鍵字：~$H，不動作")
                return None
            else:
                xlsx_path = os.path.splitext(file_path)[0] + "_converted.xlsx"
                print("輸出檔案：", xlsx_path)
                if os.path.exists(xlsx_path):
                    print("_converted.xlsx" + "檔案已存在，不動作")
                    return None
                else:
                    data_frame = pd.read_excel(file_path, sheet_name='Data1', engine='pyxlsb')
                    data_frame.to_excel(xlsx_path)
                    print(r"----------------------Convert_xlsb Done!------------------------")
                    input("enter any keys to exit")
                    return xlsx_path
        else:
            print("找不到目標檔案：", file)
'''