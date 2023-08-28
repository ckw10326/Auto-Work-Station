import pandas as pd

# 假設有兩個資料框 df1 和 df2，並且它們都有一個共同的欄位 'key'
df1 = pd.DataFrame({'key': ['A', 'B', 'C'], 'value1': [1, 2, 3]})
df2 = pd.DataFrame({'key': ['B', 'C', 'D'], 'value2': [4, 5, 6]})

# 使用 pd.merge() 函數根據 'key' 欄位將兩個資料框水平合併
merged_df = pd.merge(df1, df2, on='key')

print(merged_df)