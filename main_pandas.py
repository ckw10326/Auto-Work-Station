'''
分析EXCEL檔案，【批次】處理
'''
def read_ctc_ht_map(excel_path):
    """測試映射值"""
    field_mapping = {
        'No': ['批次序號'],
        'drawing_no_value': ['圖號:', 'CLIENTDOCNO', 'DOCVERSIONDESC'],
        'drawing_title_value': ['圖名:', 'DESCRIPTION', 'DOCCLASS'],
        'drawing_vision_value': ['版次:', 'DOCVERSIONDESC'],
        'letter_num_value': ['來文號碼:', 'TRANSMITTALNO'],
        'letter_date_value': ['來文日期:', 'REVDATE', 'PLANNEDCLIENTRETURNDATE'],
        'letter_titl_value': ['來文名稱:', 'DESCRIPTION'],
        'file_path': ['路徑']
    }

# 輸入value，輸出key

def work_flow():
    '''
    讀取【一個】檔案
    1.設定目錄路徑  2.複製檔案  3.解析檔案(目前只有HT)
    '''
    # 1.設定目錄路徑
    root_path = os.path.dirname(os.path.abspath(__file__))
    # 將根目錄路徑添加到 sys.path
    sys.path.append(root_path)
    source_folder = os.path.join(root_path, "00source")
    destination_dir = os.path.join(root_path, "00dest")
    # 檢查source_folder是否存在
    if not os.path.exists(source_folder):  # 注意這裡的小寫 "n" in "not"
        print("source_folder不存在，故結束")
        sys.exit()

    # 2.複製檔案
    move_document(source_folder, destination_dir)

    # 3.分析plan檔案，興達返回1 台中返回2
    if str(check_plan(destination_dir)) == "1":
        # 讀取興達
        # 3.1 產生xlsb列表，並轉換成xlsx
        xlsb_file_list = files_list(destination_dir, ".xlsb")
        if xlsb_file_list:
            for xlsb_path in xlsb_file_list:
                convert_xlsb(xlsb_path)

        # 3.2 產生converted.xlsx列表，開始分析內容
        xlsx_file_list = files_list(destination_dir, "converted.xlsx")
        if xlsx_file_list:
            for converted_path in xlsx_file_list:
                list0 = read_ht_excel_cloud0712.read_ht_make_list(
                    converted_path) if read_ht_excel_cloud0712.read_ht_make_list(converted_path) else 1
                input("按任意鍵退出")
                return list0
            
def main():
    """測試"""
    excel_path = "/workspaces/Auto-Work-Station/00dest/00source/HT-D1-CTC-GEL-23-2710_converted.xlsx"
    old_csv_path = "/workspaces/Auto-Work-Station/01Class/data.csv"
    read_ht_excel_cloud0712.read_ht_make_list(excel_path, old_csv_path)

    def test_movedoc():
        '''複製完整結構'''
        path1 = "/workspaces/Auto-Work-Station/09Past"
        path2 = "/workspaces/Auto-Work-Station/00dest"
        move_document(path1, path2)