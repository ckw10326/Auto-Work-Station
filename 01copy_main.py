import os
import copy_file
#source = "/workspaces/Auto-Work-Station/00source"
#dest = "/workspaces/Auto-Work-Station/00dest"

source = input("plz enter ur source")
dest = input("plz enter ur dest")

if not os.path.exists(dest):
    os.makedirs(dest)
copy_file.copy_the_excel_file2(source, dest)
input("plz enter any key")