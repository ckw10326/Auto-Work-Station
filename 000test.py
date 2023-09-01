import pandas as pd

# 原始資料
data = {
    'Unnamed: 0': [6, 7],
    'Unnamed: 1': ['NO.', '1'],
    'Unnamed: 2': ['CUSTOMER DRAWING NO.', 'HT0-1-UMM01-T6933'],
    'Unnamed: 3': ['REV.', 'A'],
    'Unnamed: 4': ['DOCUMENT TITLE', 'SAT procedure for Ammonia Vaporizer and Local ...'],
    'Unnamed: 5': ['CLASS', 'A'],
    'Unnamed: 6': ['CONTRACTOR DOCUMENT NO.', 'HT0-1-UMM01-T6933'],
    'Unnamed: 7': ['REMARK', 'NaN']
}

# 轉換為DataFrame
df = pd.DataFrame(data)
print(df,"\n\n\n")

# 指定第一列為表頭
df.columns = df.iloc[0]

# 刪除第一列
df = df[1:]

# 顯示DataFrame
print(df)