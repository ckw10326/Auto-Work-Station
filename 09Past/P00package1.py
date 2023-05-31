#抓取檔案(開放選擇)
#cd C:\Users\OXO\OneDrive\01 Book\00 Test program
#pyinstaller -F P00package1.py
#加入輸入文號功能
import shutil
import os #用來避免檔案已經存在的異常
def move_docutment():
    """
    興達範本：HT-D1-CTC-GEL-23-0522
    台中範本：TPC-TC(C0)-CD-23-1526
    """
    #a = "TPC", b = "TC(C0)", c = "CD", d = "23", e = "1526"
    #提供計畫資料
        judge0 = input("請輸入計畫(興達請輸入1,台中請輸入2:")
        e = str(input("請輸入號碼(1234):"))
        if judge0 == 1:
            a = "HT"
            b = "D1"
            c = "CTC-GEL"
            d = "23"
        else:
            a = "TPC"
            b = "TC(C0)"
            c = "CD"
            d = "23"
        File = a + "-" + b + "-" + c + "-" + d + "-" + e
        print(File)

    #生成目標路徑
        source_folder = ""
        # 1 = 雲端硬碟路徑  # 2 = 家庭硬碟路徑    # 3 = 家庭硬碟路徑
        judge1 = int(input("請輸入使用路徑(1 = 雲端, 2 = 家庭, 3 = 公司：)"))
        if judge1 == 1:
            source_folder = r"/content/sample_data" + "/" + File
        elif judge1 == 2:
            source_folder = r"C:\Users\OXO\OneDrive\01 Book\00 Test program" + "\\" + File
        elif judge1 == 3:
            source_folder = r"\\10.162.10.58\全處共用區\_Dwg\興達電廠燃氣機組更新計畫" + "/" + File
        else:
            print("error")
        print(source_folder)

        #目標資料夾
        if judge1 == 1:
            dest_folder = r"/content/" + "/" + "new" +  File
        elif judge1 == 2:
            dest_folder = r"C:\Users\OXO\OneDrive\01 Book\00 Test program" + "\\" + "new" + File
        elif judge1 == 3:
            dest_folder = r"D:\00 興達計劃\05 EPC提供資料" + "/" + "new" + File
        else:
            print("error")
        print(dest_folder)

    #確認是否存在Source_folder
        '''
        '''
    
    #確認是否存在dest_folder
        if os.path.exists(dest_folder):
            #若存在則刪除
            shutil.rmtree(dest_folder)
    #開始複製   
        shutil.copytree(source_folder, dest_folder)

        return source_folder, dest_folder

source_folder, dest_folder = move_docutment()
#顯示結果路徑
    print("-----------已完成-------------")
    print("複製路徑如下",source_folder)
    print("輸出路徑如下",dest_folder)
