'''
Table處理相關功能
1. combine_csv
2. combine_csv_list 
3. read_csv 讀取
'''
import pandas as pd


def combine_csv(src, dst):
    """整合單一csv檔案"""
    try:
        # 讀取源檔案
        df_src = pd.read_csv(src)
        # 讀取目標檔案
        df_dst = pd.read_csv(dst)
        # 合併檔案
        combined_df = pd.concat([df_src, df_dst], ignore_index=True)
        # 寫入合併後的檔案
        combined_df.to_csv(dst, index=False)
        print("檔案合併完成")

    except FileNotFoundError:
        print("找不到檔案")

    except Exception as error:
        print("發生錯誤:", str(error))


def combine_csv_list(scr_list, dst):
    """
    整合檔案列表所有csv檔案
    scr需要為完整路徑的列表
    dst需要為完整csv檔案
    """
    if isinstance(scr_list, list):
        for element in scr_list:
            df_element = pd.read_csv(element)
            df_dst = pd.read_csv(dst)
            combined_df = pd.concat([df_element, df_dst], ignore_index=True)
            combined_df.to_csv(dst, index=True)
            print(combined_df)
            print("----------------整合第", element, "的檔案-----------------")
        print("完成")
        return True
    else:
        print("輸入源不是一個列表")
        return False


def read_csv(file_path):
    """讀取CSV文件"""
    try:
        data_frame = pd.read_csv(file_path)
        print(data_frame.to_string())
    except FileNotFoundError:
        print("找不到文件")
    except Exception as error:
        print("發生錯誤:", str(error))
