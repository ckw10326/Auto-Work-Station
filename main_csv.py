import sys, os
import pandas as pd

# 1.設定目錄路徑
root_path = os.path.dirname(os.path.abspath(__file__))
# 將根目錄路徑添加到 sys.path
sys.path.append(root_path)
# 指定資料夾路徑
STANDARD_FILE = os.path.join(root_path, "00dest/stander_csv.csv")

data_frame = pd.read_csv(STANDARD_FILE)
print(data_frame)

data_frame2 = data_frame.dropna(axis=1)
print(data_frame2)