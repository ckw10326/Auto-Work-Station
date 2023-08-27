'''
pytest -s Test_module.py
cd /workspaces/Auto-Work-Station/Test
'''
import sys
import os

# 取得上層資料夾的絕對路徑
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)

# 將上層資料夾的路徑加入到 sys.path 列表中
sys.path.append(parent_dir)

# 可以在這裡進行模組的導入
import text_gen