import os

csv_folder_path = r"/workspaces/Auto-Work-Station/00dest/collectTC20230728/D-D1-GEA-GEL-22-1084"
path_company = r"//10.162.10.58/全處共用區/_Dwg/台中發電廠新建燃氣機組計畫/"
path_cloud = csv_folder_path.replace(r"/workspaces/Auto-Work-Station/00dest/collectTC20230728/", "")
dest_path = os.path.join(path_company, path_cloud)
