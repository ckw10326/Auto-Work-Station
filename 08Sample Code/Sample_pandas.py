import os
import sys
import pandas as pd


# 假設有兩個資料框 df1 和 df2，並且它們都有一個共同的欄位 'key'
df1 = pd.DataFrame({'key': ['A', 'B', 'C'], 'value1': [1, 2, 3]})
df2 = pd.DataFrame({'key': ['B', 'C', 'D'], 'value2': [4, 5, 6]})

print(df1)
print(df2)
