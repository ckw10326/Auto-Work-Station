'''
Table處理相關功能
1. combine_csv
2. combine_csv_list 
3. read_csv 讀取
'''
import os
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
            #索引列出了"Unnamed: 0.27 Unnamed: 0.26 Unnamed: 0.25 Unnamed: 0.24 Unnamed: 0.23"
            #等未命名的列。這通常是由於將索引列寫入CSV文件時產生的
            #combined_df.to_csv(dst, index=True)
            combined_df.to_csv(dst, index=None)
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
        # 嘗試使用 UTF-8 編碼讀取 CSV 文件
        data_frame = pd.read_csv(file_path, encoding='utf-8')
        print(data_frame.to_string())
    except UnicodeDecodeError:
        try:
            # 若無法使用 UTF-8 編碼，嘗試使用 Latin-1 編碼讀取 CSV 文件
            data_frame = pd.read_csv(file_path, encoding='latin-1')
            print(data_frame.to_string())
        except FileNotFoundError:
            print("找不到文件")
        except Exception as error:
            print("發生錯誤:", str(error))
    except FileNotFoundError:
        print("找不到文件")
    except Exception as error:
        print("發生錯誤:", str(error))

def txt_to_df(txt_path = "08Reference_Files/data.txt"):
    """
    2023/08/30 完成測試
    輸入txt檔案路徑
    輸出dataframe
    """
    root_path = os.path.dirname(os.path.abspath(__file__))
    txt_path = os.path.join(root_path, "08Reference_Files/data.txt")
    # 讀取文字檔案
    with open(txt_path, 'r') as file:
        lines = file.readlines()
    print("每一行的資料如下\n", lines)

    # 解析資料並建立字典
    data = {}

    # .strip() 是一個字串方法，用於移除字串開頭和結尾的空白字符（例如空格、換行符等）
    # .split('\t')，為切片
    columns = lines[0].strip().split('\t')

    # 創造data的key，並將內容設定為空白列表
    for column in columns:
        data[column] = []
    print("轉換成【字典】資料如下\n", data)

    # 將lines的第二行後面的字串先做.strip()移除開頭、結尾的空白
    for line in lines[1:]:
        values = line.strip().split('\t')
        for i, value in enumerate(values):
            data[columns[i]].append(value)

    # 建立 DataFrame
    df = pd.DataFrame(data)
    print("Dataframe資料如下\n", df.to_string())
    return df, data

def clean_df(data_frame, threshold=None, null_num=1):
    '''
    輸入data_frame，輸出dataf_frame
    threshold，用於指定保留至少多少個非空值的行或列
    nullnum，用於保留那些缺少值數量少於2的行
    重新設定表頭
    '''
    # 讀取df
    temp_data_frame = data_frame.copy()
    clean_df = None

    if threshold is not None:
        # 如果 `threshold` 不是 `None`，即有指定閾值
        clean_df = temp_data_frame.dropna(thresh=threshold)

    elif null_num is not None:
        # 只保留缺少值數量少於 null_num 的行
        clean_df = temp_data_frame[temp_data_frame.isnull().sum(
            axis=1) < null_num]

    # 指定第一列為表頭
    # 設定 DataFrame 中df.iloc[0] 用於選取 DataFrame 的第一列數據
    clean_df.columns = clean_df.iloc[0]
    # 刪除第一列
    clean_df = clean_df[1:]
    clean_df = clean_df.reset_index(drop=True)
    return clean_df

    # 將只包含完整資料的資料行寫入新的 CSV 檔案
    # clean_df.to_csv('cleaned_data.csv', index=False)
    # print("已整理並輸出整潔的資料到 cleaned_data.csv 檔案中。")

def add_column_dataframe(data: pd.DataFrame, index: str, content: str):
    """
    新增欄位以及內容
    data，以data_frame格式
    index，str格式
    content，str格式
    """
    try:
        # 新增一個名為 'index' 的欄位，其值為變數 'content'
        data[str(index)] = str(content)
        print(data.to_string())
        return data
    except FileNotFoundError:
        print("找不到文件")
    except Exception as error:
        print("發生錯誤:", str(error))