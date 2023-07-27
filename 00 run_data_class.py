from class_format import DocItem
# 讀取資料
list1 = ['HT0-1-UMM01-T6933', 
         'SAT procedure  and LCP for AmmoniaVaporizer 蒸發器SAT程序與控制面板[Ammonia Storage & Supply System]', 
         'A', 
         'HT-D1-CTC-GEL-23-2867', 
         '2023/07/19', 
         'SAT procedure  and LCP for AmmoniaVaporizer 蒸發器SAT程序與控制面板[Ammonia Storage & Supply System]', 
         '/workspaces/Auto-Work-Station/00source/HT-D1-CTC-GEL-23-2867.xlsb']
doc_item = DocItem(*list1)
doc_item.process_data0()
