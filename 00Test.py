'''
import re

# 讀取RTF文件
with open('/workspaces/Auto-Work-Station/00source/套印_TPC-TC(C0)-CD-23-2172.rtf', 'rt') as file:
    data = file.read()

# 在RTF中尋找“Test”並替換為“文係中部施工處對”
updated_data = re.sub(r'Circuit', ' For Evaporation-Crystallization System of A。', data)

# 將更新的數據寫入新的文件中
with open('/workspaces/Auto-Work-Station/00source/套印_TPC-TC(C0)-CD-23-2172_updated.rtf', 'wt') as file:
    file.write(updated_data)
'''
import re

from pyth.plugins.rtf15.reader import Rtf15Reader
from pyth.plugins.plaintext.writer import PlaintextWriter

# 開啟 RTF 檔案，使用 pyth 解析內容
with open('/workspaces/Auto-Work-Station/00source/套印_TPC-TC(C0)-CD-23-2172.rtf', 'rb') as file:
    doc = Rtf15Reader.read(file)

# 將 "circuit" 取代成 "電路板"
for p in doc.content:
    for i in range(len(p.content)):
        if p.content[i].is_text and "circuit" in p.content[i].text:
            p.content[i].text = p.content[i].text.replace("circuit", "電路板")

# 將處理過的內容輸出成純文字
output = PlaintextWriter.write(doc).getvalue()

# 印出處理好的內容
print(output)