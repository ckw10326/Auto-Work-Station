"""
find all excel file(.xlsx, .xlsb, .xlsm) of path
"""

path = "ckw10326/Auto-Work-Station/"
for root, dirs, files in os.walk(def_dest_folder):
    for file in files:
        print(file)