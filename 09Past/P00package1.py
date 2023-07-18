import os
filepath = "/workspaces/Auto-Work-Station/00dest/TPC-TC(C0)-CD-23-2388_Done.xlsx"
file_folder = os.path.split(filepath)[0]
file_allname = os.path.basename(filepath)
file_name = file_allname.split(".")[0]
print(file_folder)
print(file_allname)
print(file_name)
